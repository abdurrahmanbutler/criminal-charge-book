#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Chunk downloaded Word documents into markdown sections based on subheadings
(font size 14) and save all chunks into a single JSONL file.

- Skips files whose *filename* contains "flow chart" (case-insensitive)
- Uses a JSONL metadata file from the previous step with fields:
    order, chapter, filename, page_number, url
- For each document, splits into chunks based on subheadings (font size 14),
  only creating a new chunk when its text length would be >= 100 characters.
- Each chunk is written as one JSON object per line with fields:
    id, chunk_title, text, footnotes, order, chapter, filename,
    page_number, url

ID format (unique across all docs):
    "<unique_doc_number>-c<chunk_index>"

Where:
    - base_doc_number is taken from the start of the filename, e.g. "7.3.3A.1"
    - unique_doc_number is:
        1st time we see that base_doc_number  -> "7.3.3A.1"
        2nd time                                -> "B7.3.3A.1"
        3rd time                                -> "C7.3.3A.1"
        etc.

chunk_title rules:
    - If the heading starts with a number (e.g. "7.3.3A.1 Charge..."),
      use the heading as-is.
    - Else, if base_doc_number exists, use "<base_doc_number> <heading>".
    - Else fall back to filename-based titles.
    - Asterisks '*' are stripped from chunk_title so markdown markers don't leak into titles.
    - If the first token is duplicated at the start (e.g. "7.4.6.2 7.4.6.2 Checklist"),
      it is collapsed to a single instance ("7.4.6.2 Checklist").
"""

import argparse
import json
import os
import re
from typing import Dict, List, Optional, Set, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT


# --------------------------
# Metadata helpers
# --------------------------

def load_metadata(meta_path: str) -> List[dict]:
    records: List[dict] = []
    with open(meta_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            records.append(json.loads(line))
    return records


def extract_doc_number_from_filename(filename: str) -> Optional[str]:
    """
    Extract the leading legal-style id from a filename, including letters
    and multiple dot-separated segments.

    Examples:
        "7.3.3A.1 Charge.docx"      -> "7.3.3A.1"
        "7.2.1B Some Title.docx"    -> "7.2.1B"
        "6.2.2.3.1 Extra.docx"      -> "6.2.2.3.1"
        "1 Introductory Remarks"    -> "1"

    Pattern: first token starting with a digit, then segments separated by '.',
    where each segment is alphanumeric (digits and/or letters).
    """
    base = os.path.basename(filename)
    name, _ = os.path.splitext(base)
    m = re.match(r"\s*([0-9]+(?:\.[0-9A-Za-z]+)*)", name)
    if not m:
        return None
    return m.group(1)


def heading_starts_with_number(heading: Optional[str]) -> bool:
    """
    True if the heading text starts with a digit (after any leading whitespace
    and common invisible characters).
    """
    if not heading:
        return False
    # Strip common invisible/whitespace chars
    s = heading.lstrip("\u200e\u200f\ufeff \t\r\n")
    return bool(s and s[0].isdigit())


def sanitize_chunk_title(title: Optional[str]) -> Optional[str]:
    """
    Sanitize the chunk title:
      - Remove asterisks (markdown markers)
      - Collapse repeated first token:
        e.g. "7.4.6.2 7.4.6.2 Checklist" -> "7.4.6.2 Checklist"
    """
    if title is None:
        return None
    title = title.replace("*", "").strip()

    # Collapse repeated first token
    m = re.match(r"^(\S+)\s+\1\b(.*)$", title)
    if m:
        title = (m.group(1) + m.group(2)).strip()

    return title


# --------------------------
# Word document helpers
# --------------------------

def get_paragraph_font_size(paragraph) -> Optional[float]:
    """
    Return approximate font size (in points) of a paragraph by inspecting runs
    and, if needed, its style. Returns None if no size can be determined.
    """
    # Try runs first
    for run in paragraph.runs:
        size = run.font.size
        if size is not None:
            try:
                return size.pt  # type: ignore[attr-defined]
            except Exception:
                # Fallback to EMU value if needed
                return float(size) / 12700.0  # 1pt = 12700 EMU

    # Fall back to paragraph style
    style = getattr(paragraph, "style", None)
    if style is not None and getattr(style, "font", None) is not None:
        size = style.font.size
        if size is not None:
            try:
                return size.pt  # type: ignore[attr-defined]
            except Exception:
                return float(size) / 12700.0

    return None


def is_subheading_paragraph(paragraph) -> bool:
    """
    Subheading paragraphs are defined as those with font size ~14 pt and
    non-empty text.
    """
    size = get_paragraph_font_size(paragraph)
    if size is None:
        return False
    if abs(size - 14.0) > 0.6:
        return False
    return bool(paragraph.text.strip())


def get_heading_level(paragraph, is_first: bool, is_subheading: bool) -> Optional[int]:
    """
    Determine a heading level for markdown.

    Rules:
      - The very first paragraph in the document is always a top-level heading: "#"
      - Subheading paragraphs (14pt) are "##"
      - As a fallback, we also inspect built-in Word heading styles.
    """
    if is_first:
        return 1
    if is_subheading:
        return 2

    # Fallback based on style name
    style_name = (getattr(paragraph.style, "name", "") or "").lower()
    if "heading 1" in style_name:
        return 1
    if "heading 2" in style_name:
        return 2
    if "heading 3" in style_name:
        return 3

    # Fallback based on font size (rarely needed now)
    size = get_paragraph_font_size(paragraph)
    if size is None:
        return None
    if size > 16.0:
        return 1
    if size > 13.4:
        return 2
    return None


def build_numbering_cache(doc) -> Dict[Tuple[int, int], dict]:
    """
    Build a cache mapping (numId, ilvl) -> {fmt, lvlText}
    so we know whether a given level is decimal, bullet, lowerLetter, etc.
    """
    cache: Dict[Tuple[int, int], dict] = {}
    try:
        num_root = doc.part.numbering_part.element
    except Exception:
        return cache

    ns = num_root.nsmap

    # abstractNumId -> {ilvl -> {fmt, lvlText}}
    abstract_map: Dict[str, Dict[int, dict]] = {}

    for abs_el in num_root.findall("w:abstractNum", ns):
        abs_id = abs_el.get(qn("w:abstractNumId"))
        if abs_id is None:
            continue
        levels: Dict[int, dict] = {}
        for lvl_el in abs_el.findall("w:lvl", ns):
            ilvl_val = lvl_el.get(qn("w:ilvl")) or "0"
            try:
                ilvl = int(ilvl_val)
            except Exception:
                ilvl = 0
            fmt_el = lvl_el.find("w:numFmt", ns)
            text_el = lvl_el.find("w:lvlText", ns)
            fmt = fmt_el.get(qn("w:val")) if fmt_el is not None else None
            lvl_text = text_el.get(qn("w:val")) if text_el is not None else None
            levels[ilvl] = {"fmt": fmt, "lvlText": lvl_text}
        abstract_map[abs_id] = levels

    for num_el in num_root.findall("w:num", ns):
        num_id = num_el.get(qn("w:numId"))
        if num_id is None:
            continue
        abs_ref = num_el.find("w:abstractNumId", ns)
        if abs_ref is None:
            continue
        abs_id = abs_ref.get(qn("w:val"))
        if abs_id not in abstract_map:
            continue
        for ilvl, info in abstract_map[abs_id].items():
            try:
                key = (int(num_id), ilvl)
            except Exception:
                continue
            cache[key] = info

    return cache


def get_list_info(
    paragraph, numbering_cache
) -> Tuple[Optional[str], Optional[int], Optional[str], Optional[str], Optional[int]]:
    """
    Determine whether a paragraph is part of a list and, if so,
    return (list_type, list_level, num_fmt, lvl_text, num_id).

    list_type: 'ul' for unordered, 'ol' for ordered, None otherwise
    list_level: 0 for top-level list, 1+ for nested lists, None if not a list
    num_fmt: Word numbering format (decimal, lowerLetter, bullet, etc.) or None
    lvl_text: raw lvlText pattern like "%1." or "%1)" or None
    num_id: the Word numId (or None if unknown)
    """
    style_name = (getattr(paragraph.style, "name", "") or "").lower()
    list_type: Optional[str] = None
    num_fmt: Optional[str] = None
    lvl_text: Optional[str] = None
    level: Optional[int] = None
    num_id: Optional[int] = None
    ilvl: Optional[int] = None

    pPr = paragraph._p.pPr

    if pPr is not None and pPr.numPr is not None:
        numId_el = pPr.numPr.numId
        ilvl_el = pPr.numPr.ilvl
        if numId_el is not None and numId_el.val is not None:
            try:
                num_id = int(numId_el.val)
            except Exception:
                num_id = None
        if ilvl_el is not None and ilvl_el.val is not None:
            try:
                ilvl = int(ilvl_el.val)
            except Exception:
                ilvl = 0

        if num_id is not None and ilvl is not None:
            info = numbering_cache.get((num_id, ilvl))
            if info:
                num_fmt = info.get("fmt")
                lvl_text = info.get("lvlText")

        if num_fmt == "bullet":
            list_type = "ul"
        elif num_fmt is not None:
            list_type = "ol"

    # Fallback: style-based detection if numbering info is missing
    if list_type is None:
        if "bullet" in style_name:
            list_type = "ul"
        elif "number" in style_name or "list" in style_name:
            list_type = "ol"

    if list_type is None:
        return None, None, None, None, None

    level = ilvl if ilvl is not None else 0
    return list_type, level, num_fmt, lvl_text, num_id


def run_to_markdown(r_el) -> str:
    """
    Convert a w:r element (run) into markdown, preserving bold/italic.
    """
    ns = r_el.nsmap
    texts: List[str] = []
    for t in r_el.findall(".//w:t", namespaces=ns):
        if t.text:
            texts.append(t.text)
    text = "".join(texts)
    if not text:
        return ""

    rPr = r_el.find("w:rPr", namespaces=ns)
    bold = False
    italic = False
    if rPr is not None:
        if rPr.find("w:b", namespaces=ns) is not None:
            bold = True
        if rPr.find("w:i", namespaces=ns) is not None:
            italic = True

    if bold and italic:
        return f"***{text}***"
    if bold:
        return f"**{text}**"
    if italic:
        return f"*{text}*"
    return text


def hyperlink_to_markdown(h_el, doc) -> str:
    """
    Convert a w:hyperlink element into markdown [text](url).
    """
    ns = h_el.nsmap
    r_id = h_el.get(qn("r:id"))
    url = None
    if r_id:
        rel = doc.part.rels.get(r_id)
        if rel is not None:
            url = rel.target_ref

    # Collect the text of the hyperlink
    text_parts: List[str] = []
    for r_el in h_el.findall(".//w:r", namespaces=ns):
        text_parts.append(run_to_markdown(r_el))
    text = "".join(text_parts).strip()
    if not text:
        return ""

    if url:
        return f"[{text}]({url})"
    return text


def collect_inline_markdown(
    p_el, doc, used_footnote_ids: Optional[Set[int]] = None
) -> Tuple[str, Set[int]]:
    """
    Convert the inline content of a w:p element into markdown, capturing
    any footnote references encountered.
    """
    ns = p_el.nsmap
    parts: List[str] = []
    ids: Set[int] = set()

    for child in p_el.iterchildren():
        tag = child.tag
        if tag == qn("w:r"):
            # Normal run, possibly with a footnote reference
            parts.append(run_to_markdown(child))
            # Handle any footnote references inside this run
            for fn_ref in child.findall(".//w:footnoteReference", namespaces=ns):
                fn_id_str = fn_ref.get(qn("w:id"))
                if fn_id_str is not None:
                    try:
                        fn_id = int(fn_id_str)
                        ids.add(fn_id)
                        parts.append(f"[^{fn_id}]")
                    except ValueError:
                        pass
        elif tag == qn("w:hyperlink"):
            parts.append(hyperlink_to_markdown(child, doc))
        elif tag == qn("w:br"):
            parts.append("\n")
        # ignore other child types for now

    if used_footnote_ids is not None:
        used_footnote_ids.update(ids)

    text = "".join(parts).strip()
    return text, ids


def extract_footnotes_markdown(doc) -> Dict[int, str]:
    """
    Extract footnote texts from the document and return them as a mapping
    {footnote_id: markdown_text}.

    Uses the underlying part's XML (.blob) so it works even on python-docx
    versions where the footnotes part doesn't expose `.element`.
    If there is no footnotes part or parsing fails, returns {}.
    """
    notes: Dict[int, str] = {}
    try:
        footnotes_part = doc.part.part_related_by(RT.FOOTNOTES)
    except KeyError:
        # Document has no footnotes part
        return notes

    # Parse the raw XML of the footnotes part
    try:
        fn_el = parse_xml(footnotes_part.blob)
    except Exception:
        # If anything goes wrong, just skip footnotes rather than crashing
        return notes

    ns = fn_el.nsmap or {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    }

    for fn in fn_el.findall("w:footnote", namespaces=ns):
        id_attr = fn.get(qn("w:id"))
        if id_attr is None:
            continue
        try:
            fn_id = int(id_attr)
        except ValueError:
            continue

        # Skip separator/continuation etc. which have negative IDs
        if fn_id < 0:
            continue

        lines: List[str] = []
        for p_el in fn.findall("w:p", namespaces=ns):
            text, _ = collect_inline_markdown(p_el, doc, used_footnote_ids=None)
            if text.strip():
                lines.append(text.strip())

        notes[fn_id] = "\n".join(lines)

    return notes


def build_blocks(doc) -> List[dict]:
    """
    Build a list of blocks from the document paragraphs.
    Each block contains:
        markdown (inline, no heading/list markers yet)
        heading_level (1, 2, 3, or None)
        is_subheading (bool)
        list_type ('ul', 'ol', or None)
        list_level (int or None)
        num_fmt (Word numbering format or None)
        lvl_text (raw lvlText pattern)
        num_id (numId or None)
        list_key (hashable key for this logical list)
        used_footnote_ids (set[int])
    Empty paragraphs are skipped.
    """
    blocks: List[dict] = []
    numbering_cache = build_numbering_cache(doc)

    for idx, p in enumerate(doc.paragraphs):
        if not p.text and not p._p.getchildren():
            continue

        text, used_ids = collect_inline_markdown(p._p, doc, used_footnote_ids=set())
        if not text and not used_ids:
            # skip purely empty paragraphs
            continue

        is_sub = is_subheading_paragraph(p)
        heading_level = get_heading_level(
            p, is_first=(idx == 0), is_subheading=is_sub
        )
        list_type, list_level, num_fmt, lvl_text, num_id = get_list_info(
            p, numbering_cache
        )

        # Construct a list key so we can keep numbering across sublists, etc.
        if list_type is not None:
            if num_id is not None:
                list_key = (num_id, list_level, list_type)
            else:
                style_name = (getattr(p.style, "name", "") or "").lower()
                list_key = (list_type, list_level, style_name)
        else:
            list_key = None

        blocks.append(
            {
                "markdown": text,
                "heading_level": heading_level,
                "is_subheading": is_sub,
                "list_type": list_type,
                "list_level": list_level,
                "num_fmt": num_fmt,
                "lvl_text": lvl_text,
                "num_id": num_id,
                "list_key": list_key,
                "used_footnote_ids": used_ids,
            }
        )

    return blocks


def get_chunk_boundaries(blocks: List[dict]) -> List[Tuple[int, int]]:
    """
    Determine (start_idx, end_idx) for each chunk.

    - Chunks are anchored on subheadings (blocks where is_subheading=True).
    - The very first block (document title) starts the first chunk.
    - Text before the first subheading (if any) is included in the first chunk.
    - A new chunk is only created when its text length would be >= 100 chars.
    """
    n = len(blocks)
    if n == 0:
        return []

    # Indices of subheading blocks (excluding block 0 if it happens to look like a subheading)
    sub_indices = [i for i, b in enumerate(blocks) if b["is_subheading"] and i != 0]

    # If there are no subheadings (beyond the first title), treat the entire document as a single chunk
    if not sub_indices:
        return [(0, n - 1)]

    def span_length(start: int, end: int) -> int:
        return sum(len(blocks[i]["markdown"]) for i in range(start, end + 1))

    current_start = 0  # always start at first block
    i_sub = 0
    boundaries: List[Tuple[int, int]] = []

    while i_sub < len(sub_indices):
        # Candidate span from current_start until just before the *next* subheading (if any)
        if i_sub < len(sub_indices) - 1:
            candidate_end = sub_indices[i_sub + 1] - 1
        else:
            candidate_end = n - 1

        temp_end = candidate_end
        temp_i_sub = i_sub

        # Extend span until it reaches >= 100 characters, or we run out of subheadings
        while True:
            length = span_length(current_start, temp_end)
            if length >= 100:
                break
            if temp_i_sub + 1 >= len(sub_indices):
                # No further subheadings to merge; accept the span even if <100
                break
            # Merge the next subheading segment as well
            if temp_i_sub + 1 < len(sub_indices) - 1:
                next_end_idx = sub_indices[temp_i_sub + 2] - 1
            else:
                next_end_idx = n - 1
            temp_end = next_end_idx
            temp_i_sub += 1

        boundaries.append((current_start, temp_end))

        # Prepare for next chunk
        if temp_i_sub + 1 >= len(sub_indices):
            break
        current_start = sub_indices[temp_i_sub + 1]
        i_sub = temp_i_sub + 1

    # If we never covered the end of the document, add a final segment
    last_end = boundaries[-1][1]
    if last_end < n - 1:
        boundaries.append((last_end + 1, n - 1))

    return boundaries


def build_chunk_markdown(
    blocks: List[dict], start: int, end: int
) -> Tuple[str, Set[int], Optional[str]]:
    """
    Build markdown text and collect used footnote IDs for blocks[start:end+1].
    Returns (text, used_footnote_ids, heading_title)
    where heading_title is the first heading text in the chunk (if any).

    List handling:
      - 4 spaces per indent level for nested lists
      - Ordered list numbering is tracked per list_key (numId+level+type),
        so 1. 2. 3. ... sublist ... 4. 5. 6. stays correct.
    """
    lines: List[str] = []
    used_footnotes: Set[int] = set()
    heading_title: Optional[str] = None

    prev_is_list = False
    prev_list_type: Optional[str] = None
    prev_list_level: Optional[int] = None

    # Track ordered-list counters by logical list_key
    ol_counters_by_key: Dict[Tuple, int] = {}

    for idx in range(start, end + 1):
        block = blocks[idx]
        text = block["markdown"]
        if not text:
            continue

        used_footnotes.update(block["used_footnote_ids"])
        heading_level = block["heading_level"]
        list_type = block["list_type"]
        list_level = block["list_level"] if block["list_level"] is not None else 0
        list_key = block["list_key"]

        # Non-list paragraph
        if list_type is None:
            if prev_is_list and lines and lines[-1] != "":
                lines.append("")
            prev_is_list = False
            prev_list_type = None
            prev_list_level = None

            if heading_level is not None:
                # Heading line
                if lines and lines[-1] != "":
                    lines.append("")
                prefix = "#" * heading_level
                lines.append(f"{prefix} {text}")
                if heading_title is None:
                    heading_title = text
            else:
                # Normal paragraph
                lines.append(text)
            continue

        # List paragraph
        indent = "    " * list_level  # 4 spaces per level

        # Decide if this is a "new list block" for spacing purposes
        if not prev_is_list:
            if lines and lines[-1] != "":
                lines.append("")
        elif list_type != prev_list_type and list_level == 0:
            # Changing list type at top level → visually separate
            if lines and lines[-1] != "":
                lines.append("")

        if list_type == "ol":
            # We want numbering continuity for each logical list (list_key)
            key = list_key if list_key is not None else ("ol", list_level)
            current = ol_counters_by_key.get(key, 0) + 1
            ol_counters_by_key[key] = current
            prefix = f"{indent}{current}. "
        else:
            # Bullet
            prefix = f"{indent}- "

        lines.append(prefix + text)

        prev_is_list = True
        prev_list_type = list_type
        prev_list_level = list_level

    full_text = "\n\n".join(lines).strip()
    return full_text, used_footnotes, heading_title


def build_footnotes_markdown(
    footnotes_map: Dict[int, str], used_ids: Set[int]
) -> str:
    """
    Build markdown footnote definitions (one per line) for the used IDs.
    """
    if not used_ids:
        return ""
    lines: List[str] = []
    for fn_id in sorted(used_ids):
        body = footnotes_map.get(fn_id, "")
        if body:
            lines.append(f"[^{fn_id}]: {body}")
    return "\n".join(lines)


# --------------------------
# Document processing
# --------------------------

def process_document(
    doc_path: str,
    meta: dict,
    base_doc_number: Optional[str],
    id_doc_number: Optional[str],
) -> List[dict]:
    """
    Process a single .docx/.doc document into chunks with metadata.
    base_doc_number: the "pure" x.x.x.x... derived from filename (for titles)
    id_doc_number:   the unique version (with B/C/etc prefix if needed) used in IDs
    """
    doc = Document(doc_path)
    blocks = build_blocks(doc)
    if not blocks:
        return []

    footnotes_map = extract_footnotes_markdown(doc)
    boundaries = get_chunk_boundaries(blocks)

    chunks: List[dict] = []
    chunk_idx = 0

    base_filename = os.path.basename(meta["filename"])
    base_name_no_ext, _ = os.path.splitext(base_filename)

    for start, end in boundaries:
        chunk_idx += 1
        text, used_fn_ids, heading_title = build_chunk_markdown(blocks, start, end)

        # Enforce 100-character rule defensively (boundaries should already
        # honor it)
        if len(text) < 100 and len(boundaries) > 1:
            # Skip creating extremely small chunks if multiple chunks exist
            continue

        footnotes_md = build_footnotes_markdown(footnotes_map, used_fn_ids)

        # --- ID construction (uses unique id_doc_number) ---
        if id_doc_number:
            chunk_id = f"{id_doc_number}-c{chunk_idx}"
        else:
            chunk_id = f"{base_name_no_ext}-c{chunk_idx}"

        # --- Title construction (uses base_doc_number for human readability) ---
        if heading_title:
            stripped_heading = heading_title.strip()
            if heading_starts_with_number(stripped_heading):
                # Heading already has numbering; prefer it as-is
                chunk_title = stripped_heading
            elif base_doc_number:
                chunk_title = f"{base_doc_number} {stripped_heading}"
            else:
                chunk_title = f"{base_name_no_ext} - {stripped_heading}"
        else:
            chunk_title = base_doc_number if base_doc_number else base_name_no_ext

        # Sanitize title (remove *, collapse repeated first token)
        chunk_title = sanitize_chunk_title(chunk_title)

        result = {
            "id": chunk_id,
            "chunk_title": chunk_title,
            "text": text,
            "footnotes": footnotes_md,
        }

        # Attach original document metadata at chunk level
        for key in ("order", "chapter", "filename", "page_number", "url"):
            if key in meta:
                result[key] = meta[key]

        chunks.append(result)

    return chunks


# --------------------------
# Main
# --------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Chunk Word documents into markdown sections and save as JSONL."
    )
    parser.add_argument(
        "--docs-dir",
        required=True,
        help="Directory containing the downloaded Word files.",
    )
    parser.add_argument(
        "--meta",
        required=True,
        help="Path to JSONL metadata file (from previous download step).",
    )
    parser.add_argument(
        "--out-jsonl",
        required=True,
        help="Path to output JSONL file containing all chunks.",
    )

    args = parser.parse_args()

    docs_dir = args.docs_dir
    meta_path = args.meta
    out_path = args.out_jsonl

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    metadata_records = load_metadata(meta_path)

    total_chunks = 0

    # Track how many times we've seen each base_doc_number
    doc_number_counts: Dict[str, int] = {}

    with open(out_path, "w", encoding="utf-8") as out_f:
        for meta in metadata_records:
            filename = meta.get("filename")
            if not filename:
                continue

            # Skip "flow chart" files based on filename
            if "flow chart" in filename.lower():
                print(f"Skipping flow chart file: {filename}")
                continue

            doc_path = os.path.join(docs_dir, filename)
            if not os.path.exists(doc_path):
                print(f"WARNING: File not found, skipping: {doc_path}")
                continue

            # --- Compute base_doc_number and unique id_doc_number ---
            base_doc_number = extract_doc_number_from_filename(filename)
            if base_doc_number:
                count = doc_number_counts.get(base_doc_number, 0)
                if count == 0:
                    id_doc_number = base_doc_number
                else:
                    # For 2nd, 3rd, ... occurrences add B, C, ...
                    if count < 26:
                        letter = chr(ord("A") + count)  # 1 -> 'B', 2 -> 'C', ...
                        id_doc_number = f"{letter}{base_doc_number}"
                    else:
                        # Fallback if somehow > 26 repeats
                        id_doc_number = f"{count+1}-{base_doc_number}"
                doc_number_counts[base_doc_number] = count + 1
            else:
                base_doc_number = None
                id_doc_number = None

            print(f"Processing: {doc_path} (base_doc_number={base_doc_number}, id_doc_number={id_doc_number})")
            chunks = process_document(doc_path, meta, base_doc_number, id_doc_number)
            for chunk in chunks:
                out_f.write(json.dumps(chunk, ensure_ascii=False) + "\n")
            total_chunks += len(chunks)

    print(f"Done. Wrote {total_chunks} chunk(s) to {out_path}")


if __name__ == "__main__":
    main()
