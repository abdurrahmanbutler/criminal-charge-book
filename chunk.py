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

Chunk ID format:
    "<doc_number>-c<chunk_index>"
Where <doc_number> is taken from the start of the filename (e.g. "1.1.3")
and <chunk_index> is 1-based.

chunk_title format:
    "<doc_number> <heading of that chunk>"
If doc_number is missing, it falls back to filename-based titles.
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
    Extract the leading "x.x.x..." pattern from a filename like:
        "1.1 Something.docx" -> "1.1"
        "3.2.4 Some Title.docx" -> "3.2.4"
    Returns None if no such pattern exists.
    """
    base = os.path.basename(filename)
    name, _ = os.path.splitext(base)
    m = re.match(r"\s*(\d+(?:\.\d+)*)", name)
    if not m:
        return None
    return m.group(1)


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


def get_list_info(paragraph, numbering_cache) -> Tuple[Optional[str], Optional[int], Optional[str], Optional[str]]:
    """
    Determine whether a paragraph is part of a list and, if so,
    return (list_type, list_level, num_fmt, lvl_text).

    list_type: 'ul' for unordered, 'ol' for ordered, None otherwise
    list_level: 0 for top-level list, 1+ for nested lists, None if not a list
    num_fmt: Word numbering format (decimal, lowerLetter, bullet, etc.) or None
    lvl_text: raw lvlText pattern like "%1." or "%1)" or None
    """
    style_name = (getattr(paragraph.style, "name", "") or "").lower()
    list_type: Optional[str] = None
    num_fmt: Optional[str] = None
    lvl_text: Optional[str] = None
    level: Optional[int] = None

    pPr = paragraph._p.pPr
    num_id = None
    ilvl = None

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
        return None, None, None, None

    level = ilvl if ilvl is not None else 0
    return list_type, level, num_fmt, lvl_text


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


def collect_inline_markdown(p_el, doc, used_footnote_ids: Optional[Set[int]] = None) -> Tuple[str, Set[int]]:
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

    ns = fn_el.nsmap or {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

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
        heading_level = get_heading_level(p, is_first=(idx == 0), is_subheading=is_sub)
        list_type, list_level, num_fmt, lvl_text = get_list_info(p, numbering_cache)

        blocks.append(
            {
                "markdown": text,
                "heading_level": heading_level,
                "is_subheading": is_sub,
                "list_type": list_type,
                "list_level": list_level,
                "num_fmt": num_fmt,
                "lvl_text": lvl_text,
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
      - contiguous list blocks stay together
    """
    lines: List[str] = []
    used_footnotes: Set[int] = set()
    heading_title: Optional[str] = None

    prev_is_list = False
    prev_list_type: Optional[str] = None
    ol_counters: Dict[int, int] = {}

    for idx in range(start, end + 1):
        block = blocks[idx]
        text = block["markdown"]
        if not text:
            continue

        used_footnotes.update(block["used_footnote_ids"])
        heading_level = block["heading_level"]
        list_type = block["list_type"]
        list_level = block["list_level"] if block["list_level"] is not None else 0

        # Non-list paragraph
        if list_type is None:
            if prev_is_list and lines and lines[-1] != "":
                lines.append("")
            prev_is_list = False
            prev_list_type = None
            ol_counters = {}

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

        # If we are starting a new list block, reset counters and add a blank line
        if not prev_is_list or list_type != prev_list_type:
            if lines and lines[-1] != "":
                lines.append("")
            ol_counters = {}

        if list_type == "ol":
            current = ol_counters.get(list_level, 0) + 1
            ol_counters[list_level] = current
            prefix = f"{indent}{current}. "
        else:
            prefix = f"{indent}- "

        lines.append(prefix + text)

        prev_is_list = True
        prev_list_type = list_type

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


def process_document(
    doc_path: str, meta: dict
) -> List[dict]:
    """
    Process a single .docx/.doc document into chunks with metadata.
    Returns a list of chunk dicts.
    """
    doc = Document(doc_path)
    blocks = build_blocks(doc)
    if not blocks:
        return []

    footnotes_map = extract_footnotes_markdown(doc)
    boundaries = get_chunk_boundaries(blocks)

    chunks: List[dict] = []
    doc_number = extract_doc_number_from_filename(meta["filename"])
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

        # Chunk ID
        if doc_number:
            chunk_id = f"{doc_number}-c{chunk_idx}"
        else:
            chunk_id = f"{base_name_no_ext}-c{chunk_idx}"

        # Chunk title: doc number + heading (preferred), with fallbacks
        if doc_number and heading_title:
            chunk_title = f"{doc_number} {heading_title}"
        elif doc_number:
            chunk_title = doc_number
        elif heading_title:
            chunk_title = f"{base_name_no_ext} - {heading_title}"
        else:
            chunk_title = base_name_no_ext

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

            print(f"Processing: {doc_path}")
            chunks = process_document(doc_path, meta)
            for chunk in chunks:
                out_f.write(json.dumps(chunk, ensure_ascii=False) + "\n")
            total_chunks += len(chunks)

    print(f"Done. Wrote {total_chunks} chunk(s) to {out_path}")


if __name__ == "__main__":
    main()
