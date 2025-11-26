#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract and download all Word documents linked from a PDF, where
Word links are known to match this pattern:

    https://www.judicialcollege.vic.edu.au/media/<digits>/file

Saves:
  - Word files into a directory
  - A JSONL log with fields:
      order, chapter, filename, page_number, url

Usage:
    python extract_word_links_jcv_jsonl.py ccb.pdf \
        --out-dir word_docs \
        --log docx_links.jsonl
"""

import argparse
import json
import os
import re
from urllib.parse import urlparse, unquote

import fitz  # PyMuPDF
import requests


# Only these links are considered "Word links"
WORD_LINK_PATTERN = re.compile(
    r"^https://www\.judicialcollege\.vic\.edu\.au/media/\d+/file/?$"
)

# Chapter mapping based on the FIRST integer in the filename
CHAPTER_TITLES = {
    1: "1 – Preliminary Directions",
    2: "2 – Directions in Running",
    3: "3 – Final Directions",
    4: "4 – Evidentiary Directions",
    5: "5 – Complicity",
    6: "6 – Conspiracy, Incitement and Attempts",
    7: "7 – Victorian Offences",
    8: "8 – Victorian Defences",
    9: "9 – Commonwealth Offences",
    10: "10 – Unfitness to Stand Trial",
    11: "11 – Factual Questions and Integrated Directions",
}


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def get_filename_from_content_disposition(cd_header):
    """
    Try to pull a filename from a Content-Disposition header.
    Handles simple `filename=` and RFC 5987 `filename*=` forms.
    """
    if not cd_header:
        return None

    parts = [p.strip() for p in cd_header.split(";")]
    for part in parts:
        lower = part.lower()
        if lower.startswith("filename*="):
            # e.g. filename*=UTF-8''My%20File.docx
            _, value = part.split("=", 1)
            value = value.strip().strip('"')
            # strip charset and language if present
            try:
                _, _, encoded = value.split("'", 2)
            except ValueError:
                encoded = value
            return unquote(encoded)
        if lower.startswith("filename="):
            _, value = part.split("=", 1)
            return value.strip().strip('"')

    return None


def guess_filename_from_url(url):
    path = urlparse(url).path
    base = os.path.basename(path)
    if not base:
        return None
    return unquote(base)


def make_unique_filename(directory, filename):
    """
    If `filename` already exists in `directory`, append _1, _2, ... before the extension.
    """
    root, ext = os.path.splitext(filename)
    candidate = filename
    counter = 1
    while os.path.exists(os.path.join(directory, candidate)):
        candidate = f"{root}_{counter}{ext}"
        counter += 1
    return candidate


def ensure_word_extension(filename, content_type):
    """
    Make sure the filename ends with .doc or .docx in line with the content-type,
    adding an extension if necessary.
    """
    lower = filename.lower()
    if lower.endswith(".doc") or lower.endswith(".docx"):
        return filename

    if content_type:
        ct = content_type.split(";")[0].strip().lower()
        if ct == "application/msword":
            return filename + ".doc"
        if ct == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            return filename + ".docx"

    # Fallback: assume modern .docx
    return filename + ".docx"


def download_word_file(url, dest_dir, link_index):
    """
    Download the URL (treated as a Word document because it matches the JCV pattern).
    Returns (saved_filename or None, content_type, status).
    Status is 'ok' or 'error:...'
    """
    try:
        resp = requests.get(url, stream=True, allow_redirects=True, timeout=30)
    except Exception as e:
        return None, "", f"error:{type(e).__name__}"

    content_type = (resp.headers.get("Content-Type") or "").strip()

    # Determine filename
    cd_header = resp.headers.get("Content-Disposition")
    filename = get_filename_from_content_disposition(cd_header)

    if not filename:
        filename = guess_filename_from_url(resp.url or url)

    if not filename:
        filename = f"document_{link_index}"

    filename = ensure_word_extension(filename, content_type)
    filename = make_unique_filename(dest_dir, filename)

    dest_path = os.path.join(dest_dir, filename)

    try:
        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if not chunk:
                    continue
                f.write(chunk)
    except Exception as e:
        resp.close()
        return None, content_type, f"error:{type(e).__name__}"

    resp.close()
    return filename, content_type, "ok"


def extract_matching_links_from_pdf(pdf_path):
    """
    Generator yielding (page_number_1_based, uri) for every external link
    in the PDF that matches WORD_LINK_PATTERN.
    """
    doc = fitz.open(pdf_path)
    for page_index in range(len(doc)):
        page = doc[page_index]
        for link in page.get_links():
            uri = link.get("uri")
            if not uri:
                continue
            if WORD_LINK_PATTERN.match(uri):
                yield page_index + 1, uri  # 1-based page number


def get_chapter_from_filename(filename):
    """
    From a filename like '4.3.1 Something.docx' or '11 Some Title.docx',
    extract the FIRST integer and map it to the chapter title.
    """
    base = os.path.basename(filename)
    name_without_ext, _ = os.path.splitext(base)

    # Look for an integer at the start (possibly followed by .x.x or space)
    m = re.match(r"\s*(\d+)", name_without_ext)
    if not m:
        return None

    num = int(m.group(1))
    return CHAPTER_TITLES.get(num)


def main():
    parser = argparse.ArgumentParser(
        description="Download Word docs linked from a PDF (JCV pattern) and log to JSONL."
    )
    parser.add_argument("pdf", help="Path to the PDF file")
    parser.add_argument(
        "--out-dir",
        default="word_docs",
        help="Directory where Word files will be saved (default: word_docs)",
    )
    parser.add_argument(
        "--log",
        default=None,
        help="JSONL file to record metadata (default: <out-dir>/docx_links.jsonl)",
    )

    args = parser.parse_args()

    pdf_path = args.pdf
    out_dir = args.out_dir
    log_path = args.log or os.path.join(out_dir, "docx_links.jsonl")

    ensure_dir(out_dir)
    ensure_dir(os.path.dirname(log_path) or ".")

    # order = 1, 2, 3, ... only for successfully downloaded files
    order_index = 0
    # link_index = index among all matching links (used for fallback filenames)
    link_index = 0

    print(f"Scanning PDF: {pdf_path}")

    records = []

    for page_number, url in extract_matching_links_from_pdf(pdf_path):
        link_index += 1
        print(f"[link {link_index}] page {page_number} URL: {url}")

        filename, content_type, status = download_word_file(url, out_dir, link_index)

        if status != "ok" or not filename:
            print(f"  -> error downloading ({status})")
            continue

        order_index += 1  # only increment for successfully downloaded files
        chapter = get_chapter_from_filename(filename)

        print(
            f"  -> saved as {filename} "
            f"(order {order_index}, chapter={chapter!r})"
        )

        # Only save data for files we ACTUALLY downloaded
        record = {
            "order": order_index,
            "chapter": chapter,
            "filename": filename,
            "page_number": page_number,
            "url": url,
        }
        records.append(record)

    # Write JSONL log
    with open(log_path, "w", encoding="utf-8") as f:
        for rec in records:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")

    print()
    print(f"Done. Successfully downloaded {len(records)} file(s).")
    print(f"Saved files in: {out_dir}")
    print(f"Metadata written to (JSONL): {log_path}")


if __name__ == "__main__":
    main()
