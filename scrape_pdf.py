#!/usr/bin/env python3
import argparse
import json
import re
from typing import List, Dict, Tuple, Set, Optional

import pdfplumber
from tqdm.auto import tqdm

# How many front pages to skip when chunking (table of contents etc.)
SKIP_FIRST_PAGES = 9

# Numbered heading patterns
CHAPTER_RE = re.compile(r'^(?P<num>[1-9]\d?)\s+(?P<title>.+)$')          # 1 .. 99
SUBCHAPTER_RE = re.compile(r'^(?P<num>\d+\.\d+)\s+(?P<title>.+)$')      # 1.1, 6.3
SECTION_RE = re.compile(r'^(?P<num>\d+\.\d+\.\d+)\s+(?P<title>.+)$')    # 1.2.1, 6.3.1

# For subsection heading heuristics
SUBSECTION_EXCLUDE_PREFIXES = [
    "Download a Microsoft Word version",
    "Last updated:",
]
COMMON_LOWER_WORDS = {
    "and", "or", "of", "the", "to", "a", "an", "in", "for",
    "on", "at", "with", "without", "by", "as", "is", "are",
    "be", "from",
}

# Regex used to parse lines from the table of contents
TOC_HEADING_RE = re.compile(
    r'^(?P<num>\d+(?:\.\d+){0,2})\s+(?P<title>.+?)(?:\s+\.{2,}\s*\d+)?$'
)


def normalize_heading(num: Optional[str], title: Optional[str]) -> str:
    """Return '1.2.1 – Title' or just 'Title'."""
    if num and title:
        return f"{num} – {title}"
    if title:
        return title
    return ""


def convert_line_to_markdown(line: str) -> str:
    """
    Convert a single PDF text line into markdown-ish text:

    - bullet characters -> '- item'
    - "1.  Text" -> "1. Text"
    - "(a) Text" -> indented ordered item
    """
    s = line.rstrip()
    if not s:
        return ""

    stripped = s.lstrip()

    # Bullet points: PDF often uses bullets like '•', '', etc.
    if stripped.startswith(("\uf0b7", "•", "")):
        content = stripped[1:].lstrip()
        return f"- {content}"

    # Numbered list: "1.  Text" -> "1. Text"
    m = re.match(r"^(\d+)\.\s{0,3}(.*)$", stripped)
    if m:
        num, rest = m.group(1), m.group(2)
        return f"{num}. {rest}"

    # Sublist like "(a) Text"
    m = re.match(r"^\(([a-z])\)\s+(.*)$", stripped)
    if m:
        rest = m.group(2)
        # Represent as indented ordered item
        return f"   1. {rest}"

    return s


def looks_like_subsection_heading(line: str) -> bool:
    """
    Heuristic to detect unnumbered subsection headings like:
      - Information to Provide to Jurors
      - Excusing Jurors
      - Jury Directions and Statutory Language
      - Inciting Conduct
      - Duration of the trial
      - Other Reasons
    """
    line = line.strip()
    if not line:
        return False

    # Must start with a capital letter
    if not line[0].isupper():
        return False

    # Exclude clearly numbered or section-style lines
    if re.match(r"^\d+[\.\)]\s", line):
        return False
    if re.match(r"^\d+\.\d+(\.\d+)*\s", line):
        return False
    if re.match(r"^\d+\s+\S", line):
        return False

    lower = line.lower()
    for pref in SUBSECTION_EXCLUDE_PREFIXES:
        if lower.startswith(pref.lower()):
            return False

    # Do not end with sentence punctuation
    if line[-1] in ".?!;":
        return False

    # Length / word heuristics
    words = [w for w in re.split(r"\s+", line) if w]
    if len(words) == 0 or len(words) > 8:  # allow up to 8 words
        return False

    # Check proportion of "title-like" words: capitalised or common connectors
    caps = 0
    for w in words:
        if w[0].isupper() or w[0].isdigit() or w.lower() in COMMON_LOWER_WORDS:
            caps += 1
    ratio = caps / len(words)

    if ratio < 0.5:
        # Also allow pattern: First word capitalised, remaining 1–3 lower-case words
        cleaned = re.sub(r"[^A-Za-z\s]", "", line)
        if re.fullmatch(r"[A-Z][a-z]+(?:\s+[a-z]+){0,3}", cleaned):
            return True
        return False

    return True


def footnote_lines_to_markdown(lines: List[str]) -> str:
    """
    Convert raw footnote lines for a page into ordered-list markdown:

      1. Text of footnote 1 ...
      2. Text of footnote 2 ...
    """
    items: List[str] = []
    current: Optional[str] = None

    for ln in lines:
        s = ln.rstrip()
        stripped = s.strip()
        if not stripped:
            continue

        m = re.match(r"^(\d+)\s+(.*)$", stripped)
        if m:
            # New footnote
            if current is not None:
                items.append(current.strip())
            num, rest = m.group(1), m.group(2)
            current = f"{num}. {rest}"
        else:
            # Continuation of current footnote
            if current is None:
                current = stripped
            else:
                current += " " + stripped

    if current is not None:
        items.append(current.strip())

    return "\n".join(items)


# ----------------------------------------------------------------------
#   Heading catalog from table of contents
# ----------------------------------------------------------------------

def build_heading_catalog(pdf_path: str, max_toc_pages: int = 10) -> Dict[str, str]:
    """
    Scan the first few pages (table of contents) and build a mapping:
      '1' -> 'Preliminary Directions'
      '1.1' -> 'Introductory Remarks'
      '1.2.1' -> 'Charge: Jury Empanelment'
      etc.
    """
    catalog: Dict[str, str] = {}

    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages_to_scan = min(max_toc_pages, len(pdf.pages))
            for idx in range(pages_to_scan):
                page = pdf.pages[idx]
                text = page.extract_text() or ""
                for line in text.splitlines():
                    s = line.strip()
                    m = TOC_HEADING_RE.match(s)
                    if not m:
                        continue
                    num = m.group("num")
                    title = m.group("title").strip().rstrip(".")
                    # Only keep up to three levels  (1, 1.1, 1.1.1)
                    if num.count(".") <= 2:
                        catalog[num] = title
    except Exception:
        # If anything goes wrong, just return an empty catalog
        pass

    return catalog


# ----------------------------------------------------------------------
#   PDF → linear text entries + per-page footnotes
# ----------------------------------------------------------------------

def extract_pages_for_document(
    pdf_path: str,
    skip_first_pages: int = SKIP_FIRST_PAGES,
) -> Tuple[List[Dict], Dict[int, List[str]]]:
    """
    Read the PDF with pdfplumber, skip the front matter, and return:
      - doc_entries: [{text, page, line_index}, ...]
      - footnotes_by_page: {logical_page_number: [raw footnote lines]}
    """
    doc_entries: List[Dict] = []
    footnotes_by_page: Dict[int, List[str]] = {}

    with pdfplumber.open(pdf_path) as pdf:
        total_pages = len(pdf.pages)

        for idx in tqdm(
            range(skip_first_pages, total_pages),
            desc="Reading pages",
        ):
            page = pdf.pages[idx]
            text = page.extract_text()
            if not text:
                continue

            lines = text.splitlines()

            # Strip header lines with the book title
            while lines and "Criminal Charge Book" in lines[0]:
                lines.pop(0)

            # Leading / trailing blank lines
            while lines and not lines[0].strip():
                lines.pop(0)
            while lines and not lines[-1].strip():
                lines.pop()

            # Remove trailing printed page number (digits only) if present
            page_number = None
            if lines and re.fullmatch(r"\d+", lines[-1].strip()):
                page_number = int(lines.pop().strip())

            logical_page = page_number if page_number is not None else (idx + 1)

            # Clean trailing blanks again
            while lines and not lines[-1].strip():
                lines.pop()

            # Detect footnotes near the bottom:
            #   - lines starting with "1 " "2 " ... but NOT "1." numbered paragraphs
            #   - only in bottom ~40% of the page
            n = len(lines)
            fn_indices: List[int] = []
            for i, ln in enumerate(lines):
                s = ln.strip()
                if re.match(r"^\d+\s+\S", s) and not re.match(r"^\d+\.\s", s):
                    if i >= int(n * 0.6):
                        fn_indices.append(i)

            if fn_indices:
                first_fn = fn_indices[0]
                footnotes_by_page[logical_page] = lines[first_fn:]
                lines = lines[:first_fn]

            # Record remaining lines as body text
            for line_idx, ln in enumerate(lines):
                doc_entries.append(
                    {
                        "text": ln.rstrip(),
                        "page": logical_page,
                        "line_index": line_idx,
                    }
                )

    return doc_entries, footnotes_by_page


# ----------------------------------------------------------------------
#   Heading classification (chapter / subchapter / section)
# ----------------------------------------------------------------------

def classify_heading_entry(
    entry: Dict,
    heading_catalog: Dict[str, str],
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Decide if a line is a chapter / subchapter / section heading.
    Uses both text and line_index (position on page), plus TOC catalog.
    """
    s = entry["text"].strip()
    li = entry["line_index"]

    # SECTION (e.g. 6.3.1 Charge: Incitement (Victoria))
    m = SECTION_RE.match(s)
    if m:
        num = m.group("num")
        # Use TOC title if available, otherwise raw title
        title_raw = m.group("title").strip()
        title = heading_catalog.get(num, title_raw)
        return "section", num, title

    # SUBCHAPTER (e.g. 6.3 Incitement (Victoria))
    m = SUBCHAPTER_RE.match(s)
    if m:
        num = m.group("num")
        title_raw = m.group("title").strip()
        title = heading_catalog.get(num, title_raw)
        return "subchapter", num, title

    # CHAPTER (e.g. 6 Inchoate Offences) – only near the top of the page
    if li <= 3:
        m = CHAPTER_RE.match(s)
        if m:
            num = m.group("num")
            title_raw = m.group("title").strip()

            # Hard filters to reject case citations (e.g. "1 VR 290)", "140 CLR 342).")
            #  - Only allow 1–2 digit chapter numbers (regex already restricts 1–99)
            #  - Require at least one lowercase letter in the raw title
            #  - Reject common law-report patterns like "VR", "CLR", "HCA", "VSCA"
            if not re.search(r"[a-z]", title_raw):
                return None, None, None
            if re.search(r"\b(VR|CLR|VSC|VSCA|HCA)\b", title_raw):
                return None, None, None

            # TOC-based canonical title if available
            title = heading_catalog.get(num, title_raw)
            return "chapter", num, title

    return None, None, None


# ----------------------------------------------------------------------
#   Chunking logic
# ----------------------------------------------------------------------

def chunk_document_entries(
    doc_entries: List[Dict],
    footnotes_by_page: Dict[int, List[str]],
    heading_catalog: Dict[str, str],
) -> List[Dict]:
    """
    Turn linear pdf text into structured chunks:
      - chapter / subchapter / section / subsection
      - markdown 'text'
      - markdown 'citations'
      - page_numbers
    """
    chunks: List[Dict] = []

    current_chapter: Optional[str] = None
    current_subchapter: Optional[str] = None
    current_section: Optional[str] = None
    current_subsection: Optional[str] = None

    current_lines: List[str] = []
    current_pages: Set[int] = set()
    chunk_id = 0

    def flush_chunk():
        nonlocal chunk_id, current_lines, current_pages, current_subsection

        if not current_lines:
            return

        # Derive subsection title:
        #  - explicit unnumbered heading if present
        #  - otherwise fall back to *section* title text (without the number)
        #  - otherwise "Overview"
        if current_subsection:
            subsection_title = current_subsection
        elif current_section:
            # Take the part after "–" if present
            if "–" in current_section:
                subsection_title = current_section.split("–", 1)[1].strip()
            else:
                subsection_title = current_section
        else:
            subsection_title = "Overview"

        # Heading stack at the top of each chunk
        lines_md: List[str] = []
        if current_chapter:
            lines_md.append(f"# {current_chapter}")
        if current_subchapter:
            lines_md.append(f"## {current_subchapter}")
        if current_section:
            lines_md.append(f"### {current_section}")
        if subsection_title:
            lines_md.append(f"#### {subsection_title}")
        lines_md.append("")  # blank line before body

        for l in current_lines:
            md_line = convert_line_to_markdown(l)
            lines_md.append(md_line)

        text_md = "\n".join(lines_md).rstrip()

        # Citations: combine footnotes from all pages spanned by this chunk
        citations_parts: List[str] = []
        for p in sorted(current_pages):
            fns = footnotes_by_page.get(p)
            if fns:
                citations_parts.append(footnote_lines_to_markdown(fns))
        citations_text = "\n".join(citations_parts)

        chunk = {
            "id": f"chunk-{chunk_id}",
            "chapter": current_chapter or "",
            "subchapter": current_subchapter or "",
            "section": current_section or "",
            "subsection": subsection_title,
            "text": text_md,
            "citations": citations_text,
            "page_numbers": sorted(current_pages),
        }
        chunks.append(chunk)
        chunk_id += 1

        current_lines = []
        current_pages = set()
        # Do NOT reset current_subsection here; it is reset when a new heading appears

    # Main pass over all lines
    for entry in tqdm(doc_entries, desc="Chunking lines"):
        line = entry["text"]
        page = entry["page"]
        s = line.strip()

        # Preserve blank lines as paragraph breaks inside a chunk
        if not s:
            if current_lines and current_lines[-1] != "":
                current_lines.append("")
                current_pages.add(page)
            continue

        # Skip metadata
        lower = s.lower()
        if lower.startswith("last updated:") or "download a microsoft word version" in lower:
            continue

        # Heading classification (chapter / subchapter / section)
        kind, num, title = classify_heading_entry(entry, heading_catalog)
        if kind == "chapter":
            flush_chunk()
            current_chapter = normalize_heading(num, title)
            current_subchapter = None
            current_section = None
            current_subsection = None
            continue

        if kind == "subchapter":
            flush_chunk()
            current_subchapter = normalize_heading(num, title)
            current_section = None
            current_subsection = None
            continue

        if kind == "section":
            flush_chunk()
            current_section = normalize_heading(num, title)
            current_subsection = None
            continue

        # Unnumbered subsection heading?
        if looks_like_subsection_heading(s):
            flush_chunk()
            current_subsection = s
            continue

        # Normal body text
        current_lines.append(line)
        current_pages.add(page)

    # Flush last chunk
    flush_chunk()
    return chunks


def write_jsonl(chunks: List[Dict], out_path: str) -> None:
    with open(out_path, "w", encoding="utf-8") as f:
        for ch in chunks:
            json.dump(ch, f, ensure_ascii=False)
            f.write("\n")


def main():
    parser = argparse.ArgumentParser(
        description="Chunk Criminal Charge Book PDF into chapter/subchapter/section/subsection JSONL."
    )
    parser.add_argument("pdf_path", help="Path to ccb.pdf")
    parser.add_argument("out_path", help="Path to output JSONL")
    parser.add_argument(
        "--skip-pages",
        type=int,
        default=SKIP_FIRST_PAGES,
        help="Number of pages at start of PDF to skip when chunking (default: 9)",
    )
    parser.add_argument(
        "--toc-pages",
        type=int,
        default=10,
        help="Maximum number of pages to scan for the table of contents (default: 10)",
    )
    args = parser.parse_args()

    # 1) Build heading catalog from the table of contents
    heading_catalog = build_heading_catalog(args.pdf_path, max_toc_pages=args.toc_pages)

    # 2) Extract linear text & per-page footnotes (skipping TOC pages for body)
    doc_entries, footnotes_by_page = extract_pages_for_document(
        args.pdf_path, skip_first_pages=args.skip_pages
    )

    # 3) Chunk into structured JSON
    chunks = chunk_document_entries(doc_entries, footnotes_by_page, heading_catalog)

    # 4) Write JSONL
    write_jsonl(chunks, args.out_path)


if __name__ == "__main__":
    main()
