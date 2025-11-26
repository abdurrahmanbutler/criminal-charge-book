#!/usr/bin/env python3
import argparse
import json
import os
import re
from typing import Dict, List, Tuple

import pypandoc


FOOTNOTE_DEF_RE = re.compile(r'^\[\^([^\]]+)\]:\s*(.*)')


def docx_to_markdown(path: str) -> str:
    """
    Convert a .docx (or .doc that is actually OOXML) file to Markdown using pandoc.
    We force .doc and .docx to be treated as 'docx' so pandoc doesn't try 'doc'.
    If pandoc is missing, we download it once via pypandoc.
    """
    ext = os.path.splitext(path)[1].lower()

    # Force .doc and .docx to be treated as 'docx' by pandoc
    input_format = None
    if ext in (".doc", ".docx"):
        input_format = "docx"

    try:
        return pypandoc.convert_file(
            path,
            "markdown",
            format=input_format,
            extra_args=["--wrap=none"],
        )
    except OSError as e:
        # Handle "No pandoc was found" by downloading it once
        if "No pandoc was found" in str(e):
            print("Pandoc not found. Downloading pandoc via pypandoc...")
            pypandoc.download_pandoc()
            return pypandoc.convert_file(
                path,
                "markdown",
                format=input_format,
                extra_args=["--wrap=none"],
            )
        raise


def clean_markdown_artifacts(md: str) -> str:
    """
    Remove annoying pandoc artifacts like:

        ```{=html}
        <!-- -->
        ```

    and standalone `{=html}` / `<!-- -->` lines.
    """
    lines = md.splitlines()
    out_lines: List[str] = []
    in_artifact_block = False

    for line in lines:
        stripped = line.strip()

        # Start of artifact code block
        if stripped.startswith("```") and "{=html}" in stripped:
            in_artifact_block = True
            continue

        # Inside artifact code block
        if in_artifact_block:
            if stripped.startswith("```"):
                in_artifact_block = False
            # Skip everything inside
            continue

        # Skip stray artifact / comment lines
        if stripped in ("{=html}", "<!-- -->"):
            continue

        out_lines.append(line)

    return "\n".join(out_lines)


def split_markdown_and_footnotes(md: str) -> Tuple[str, Dict[str, str]]:
    """
    Separate the main markdown content from footnote definitions.

    Returns:
        content_md: markdown without any footnote definition blocks
        footnotes: dict mapping footnote id -> full markdown block, e.g.
                   "[^1]: Footnote text..."
                   (including any indented continuation lines)
    """
    lines = md.splitlines()
    content_lines: List[str] = []
    footnotes: Dict[str, List[str]] = {}

    current_id = None
    current_lines: List[str] = []

    for line in lines:
        # New footnote definition?
        m = FOOTNOTE_DEF_RE.match(line)
        if m:
            # flush previous footnote, if any
            if current_id is not None:
                footnotes[current_id] = current_lines

            current_id = m.group(1)
            current_lines = [line]
            continue

        # Continuation of a footnote block (indented line)
        if current_id is not None and (line.startswith("    ") or line.startswith("\t")):
            current_lines.append(line)
            continue

        # End of a footnote block
        if current_id is not None:
            footnotes[current_id] = current_lines
            current_id = None
            current_lines = []

        # Regular content line
        content_lines.append(line)

    # Flush last footnote
    if current_id is not None:
        footnotes[current_id] = current_lines

    # Join blocks
    footnote_blocks = {fid: "\n".join(block) for fid, block in footnotes.items()}
    content_md = "\n".join(content_lines)

    return content_md, footnote_blocks


def parse_headings(content_md: str) -> List[Tuple[int, int, str]]:
    """
    Find markdown headings in the content (after removing footnote defs).

    Returns a list of (line_index, level, title).
    """
    lines = content_md.splitlines()
    headings: List[Tuple[int, int, str]] = []

    heading_re = re.compile(r'^(#{1,6})\s+(.+?)\s*$')

    for idx, line in enumerate(lines):
        m = heading_re.match(line)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()
            headings.append((idx, level, title))

    return headings


def looks_like_list_item(line: str) -> bool:
    """Return True if the line looks like a numbered/bulleted list item."""
    stripped = line.lstrip()
    # numbered: "1. ", "1) ", "i) ", "(a) "
    if re.match(r'^(\d+[\).]|\(?[a-zA-Z]\)|[ivxlcdm]+\.)\s', stripped):
        return True
    # bullets: "- ", "* ", "• "
    if stripped.startswith(('-', '*', '•')):
        return True
    return False


def is_pseudo_heading(line: str, prev_line: str, next_line: str) -> bool:
    """
    Heuristic: treat a plain line as a "heading" when it looks like the
    subheadings in the Criminal Charge Book docs (e.g. 'Overview', etc.),
    including numbered titles like '7.5.5.8 Checklist: ...'.
    """
    line = line.strip()
    prev_line = prev_line.strip()
    next_line = next_line.strip()

    if not line:
        return False

    # Ignore code fences / comments
    if line.startswith("```") or line.startswith("<!--"):
        return False
    # Ignore footnote defs
    if re.match(r'^\[\^[^\]]+\]:', line):
        return False
    # Ignore potential link/image refs
    if line.startswith("["):
        return False

    # Ignore obvious list items
    if looks_like_list_item(line):
        return False

    # Ignore "Last updated: 2 March 2015" footer etc
    if line.lower().startswith("last updated"):
        return False

    # We only promote lines that are visually separated:
    # require a blank line before
    if prev_line != "":
        return False

    # Basic sanity on length / characters
    words = line.split()
    if len(words) < 1 or len(words) > 16:
        return False
    if not re.search(r"[A-Za-z]", line):
        return False

    # Allow either:
    #  - starts with uppercase letter, OR
    #  - starts with a multi-level number then an uppercase word, e.g. "7.5.5.8 Checklist: ..."
    if re.match(r"^[A-Z]", line):
        first_char_ok = True
    else:
        m = re.match(r"^(\d+(?:\.\d+)*)\s+(.+)$", line)
        if not m:
            return False
        title_part = m.group(2).lstrip()
        first_char_ok = bool(title_part and re.match(r"[A-Z]", title_part[0]))

    if not first_char_ok:
        return False

    # Headings usually don't end with full stop / colon / etc.
    if re.search(r"[.!?;:]$", line):
        return False

    return True


def find_pseudo_headings(lines: List[str], start_idx: int = 0) -> List[Tuple[int, str]]:
    """
    Scan through lines[start_idx:] and return a list of (line_index, title)
    for lines that look like headings but are not marked with '#'.
    """
    pseudo: List[Tuple[int, str]] = []
    for i in range(start_idx, len(lines)):
        line = lines[i]
        prev_line = lines[i - 1] if i > 0 else ""
        next_line = lines[i + 1] if i + 1 < len(lines) else ""
        if is_pseudo_heading(line, prev_line, next_line):
            pseudo.append((i, line.strip()))
    return pseudo


def detect_checklist_section_title(lines: List[str], max_scan: int = 80) -> Tuple[int, str]:
    """
    Look near the top of the document for a line that looks like a 'Checklist'
    title, e.g. '7.5.5.8 Checklist: Aggravated Burglary Combined Bases'.

    Returns (line_index, title) or (-1, "") if none found.
    """
    for i, line in enumerate(lines[:max_scan]):
        s = line.strip()
        if not s:
            continue
        lower = s.lower()
        if "checklist" not in lower:
            continue

        # Ignore code fences / comments / footnotes
        if s.startswith("```") or s.startswith("<!--"):
            continue
        if re.match(r'^\[\^[^\]]+\]:', s):
            continue

        # This is good enough as a "checklist" title
        return i, s

    return -1, ""


def chunk_by_subheadings(content_md: str, fallback_section_title: str):
    """
    Chunk the markdown by subheadings.

    Rules:
      - If a plausible "checklist" title is found near the top, we treat
        the *entire document* as a single chunk using that title.
      - Otherwise:
        * section-title = the first heading in the document (if any),
                          otherwise the first pseudo-heading (if any),
                          otherwise fallback_section_title.
        * Prefer real Markdown headings (#, ##, etc.) when there are >= 2.
        * If there is only one real heading (the section title) or none,
          use heuristic "pseudo-headings" like 'Overview', etc., as chunk
          boundaries.
        * Each (sub)heading becomes a chunk_title, and the chunk_text is all
          markdown between that heading and the next heading.
    """
    lines = content_md.splitlines()

    # --- Checklist special case: detect title and don't chunk
    checklist_idx, checklist_title = detect_checklist_section_title(lines)
    if checklist_idx != -1 and checklist_title:
        section_title = checklist_title
        full_text = content_md.strip()
        if not full_text:
            return section_title, []
        # One big chunk only
        return section_title, [(section_title, full_text)]

    # --- Normal (non-checklist) behaviour
    headings = parse_headings(content_md)  # true '#' headings

    # Start by assuming section-title is the first markdown heading, if any.
    if headings:
        section_title = headings[0][2]
        body_start = headings[0][0] + 1
    else:
        section_title = fallback_section_title
        body_start = 0

    # Case 1: we have multiple true Markdown headings -> use them directly
    if len(headings) >= 2:
        chunks = []
        for i in range(1, len(headings)):
            line_idx, level, title = headings[i]
            next_idx = headings[i + 1][0] if i + 1 < len(headings) else len(lines)
            body_lines = lines[line_idx + 1: next_idx]
            chunk_text = "\n".join(body_lines).strip()
            if chunk_text:
                chunks.append((title, chunk_text))

        if chunks:
            # Already have nice heading-based chunks, no need for heuristics
            return section_title, chunks

    # Case 2: only one or zero true headings -> fall back to pseudo-headings
    pseudo = find_pseudo_headings(lines, start_idx=body_start)

    # Special handling when there were NO markdown headings at all:
    # treat the first pseudo-heading as the section-title, and only chunk
    # on subsequent pseudo-headings.
    if not headings and pseudo:
        section_title = pseudo[0][1]
        pseudo_sub = pseudo[1:]
    else:
        pseudo_sub = pseudo

    chunks: List[Tuple[str, str]] = []

    if not pseudo_sub:
        # No pseudo-headings either -> one big chunk
        body_text = "\n".join(lines[body_start:]).strip()
        if not body_text:
            return section_title, []
        return section_title, [(section_title, body_text)]

    # Build chunks from pseudo-subheadings
    for idx, (line_idx, title) in enumerate(pseudo_sub):
        start = line_idx + 1
        end = pseudo_sub[idx + 1][0] if idx + 1 < len(pseudo_sub) else len(lines)
        chunk_text = "\n".join(lines[start:end]).strip()
        if chunk_text:
            chunks.append((title, chunk_text))

    return section_title, chunks


def extract_footnotes_for_chunk(chunk_text: str, footnote_defs: Dict[str, str]) -> str:
    """
    From a chunk's text, find referenced footnote ids (e.g. [^1]) and
    return a markdown block containing only the corresponding definitions.
    """
    ids = set(re.findall(r'\[\^([^\]]+)\]', chunk_text))
    if not ids:
        return ""

    def sort_key(x: str):
        return (0, int(x)) if x.isdigit() else (1, x)

    lines: List[str] = []
    for fid in sorted(ids, key=sort_key):
        block = footnote_defs.get(fid)
        if block:
            lines.append(block)

    return "\n".join(lines).strip()


def is_flowchart_document(markdown: str, doc_name: str) -> bool:
    """
    Heuristic to skip flowchart-only documents.

    - If the filename contains 'flowchart' -> skip.
    - Or if the first non-empty line contains 'flowchart' / 'flow chart'
      and there is very little additional text.
    """
    name_lower = doc_name.lower()
    if "flowchart" in name_lower or "flow chart" in name_lower:
        return True

    lines = [l.strip() for l in markdown.splitlines() if l.strip()]
    if not lines:
        return False

    first_line = lines[0].lower()
    if "flowchart" in first_line or "flow chart" in first_line:
        non_title_text = " ".join(lines[1:])
        # Tiny amount of body text -> almost certainly just the diagram doc
        if len(non_title_text) < 400:
            return True

    return False


def process_docx_file(path: str, doc_name: str, chunk_id_start: int):
    """
    Process a single DOCX/DOC file:
      - convert to markdown
      - clean pandoc artifacts
      - split off footnotes
      - skip flowchart-only docs
      - chunk by subheadings (including heuristics + checklist detection)
      - build JSON-able dicts for each chunk

    Returns:
      records: list of dicts for JSONL
      next_chunk_id: next available global id counter
    """
    print(f"Converting {path} to markdown...")
    markdown = docx_to_markdown(path)
    markdown = clean_markdown_artifacts(markdown)

    # Skip flowchart-only docs
    if is_flowchart_document(markdown, doc_name):
        print(f"Skipping flowchart document: {doc_name}")
        return [], chunk_id_start

    content_md, footnote_defs = split_markdown_and_footnotes(markdown)
    fallback_section_title = os.path.splitext(doc_name)[0]

    section_title, chunks = chunk_by_subheadings(content_md, fallback_section_title)

    records: List[dict] = []
    current_id = chunk_id_start

    for raw_chunk_title, chunk_text in chunks:
        current_id += 1
        # Remove leading zeros: just use the integer as a string
        chunk_id = str(current_id)

        # build semantic chunk_title: "Section title – Chunk title" (EN DASH)
        if raw_chunk_title == section_title or not raw_chunk_title:
            combined_title = section_title
        else:
            combined_title = f"{section_title} – {raw_chunk_title}"

        footnotes_md = extract_footnotes_for_chunk(chunk_text, footnote_defs)

        record = {
            "id": chunk_id,
            "section-title": section_title,
            "chunk_title": combined_title,
            "text": chunk_text,
            "footnotes": footnotes_md,
            "doc_name": doc_name,
        }
        records.append(record)

    return records, current_id


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Chunk Word documents by subheading and write a JSONL "
            "with id, section-title, chunk_title, text, footnotes, doc_name."
        )
    )
    parser.add_argument("input_dir", help="Directory containing individual .docx/.doc files")
    parser.add_argument("output_jsonl", help="Path to output JSONL file")
    parser.add_argument(
        "--skip-combined",
        default="combined.docx",
        help="Filename to skip (e.g. the merged document) [default: combined.docx]",
    )

    args = parser.parse_args()

    input_dir = args.input_dir
    output_path = args.output_jsonl
    skip_combined = args.skip_combined.lower()

    all_records: List[dict] = []
    chunk_id_counter = 0

    for fname in sorted(os.listdir(input_dir)):
        lower = fname.lower()
        if not (lower.endswith(".docx") or lower.endswith(".doc")):
            continue
        if lower == skip_combined:
            continue
        if lower.startswith("~$"):  # skip temp/lock files
            continue

        doc_path = os.path.join(input_dir, fname)
        print(f"Processing {doc_path} ...")

        records, chunk_id_counter = process_docx_file(doc_path, fname, chunk_id_counter)
        all_records.extend(records)

    print(f"Writing {len(all_records)} chunks to {output_path} ...")
    with open(output_path, "w", encoding="utf-8") as f:
        for rec in all_records:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    print("Done.")


if __name__ == "__main__":
    main()
