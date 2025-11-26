#!/usr/bin/env python3
import argparse
import copy
import os
import re
import sys
from urllib.parse import urlparse

import fitz  # PyMuPDF: pip install pymupdf
import requests  # pip install requests
from docx import Document  # pip install python-docx
from docxcompose.composer import Composer  # NEW

PHRASE = "Download a Microsoft Word version of this"


def extract_word_urls(pdf_path: str) -> list[str]:
    """
    Scan the PDF for pages containing the Word-download phrase and
    return all matching judicialcollege /media/<id>/file URLs
    in the order they appear in the document.
    """
    doc = fitz.open(pdf_path)
    urls: list[str] = []
    seen: set[str] = set()

    media_pattern = re.compile(r"^/media/\d+/file/?$")

    for page_index in range(doc.page_count):
        page = doc.load_page(page_index)
        text = page.get_text("text")

        # Only look at pages that have the download phrase
        if PHRASE not in text:
            continue

        for link in page.get_links():
            uri = link.get("uri")
            if not uri:
                continue

            if "judicialcollege.vic.edu.au/media" not in uri:
                continue

            parsed = urlparse(uri)
            if not media_pattern.match(parsed.path):
                continue

            # Normalise to https and strip trailing slash
            scheme = "https"
            netloc = parsed.netloc or "www.judicialcollege.vic.edu.au"
            norm = f"{scheme}://{netloc}{parsed.path.rstrip('/')}"

            if norm not in seen:
                seen.add(norm)
                urls.append(norm)

    return urls


def guess_extension(content_type: str) -> str:
    """
    Guess a reasonable file extension from the HTTP Content-Type.
    Defaults to .docx if we can't tell.
    """
    if not content_type:
        return ".docx"

    ct = content_type.split(";")[0].strip().lower()
    if ct == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return ".docx"
    if ct == "application/msword":
        return ".doc"
    return ".docx"


def download_all(urls: list[str], out_dir: str) -> list[str]:
    """
    Download each URL into out_dir, naming files in the order they appear
    in the PDF: 001.docx, 002.docx, etc. Returns a list of paths in order.
    """
    os.makedirs(out_dir, exist_ok=True)
    session = requests.Session()

    downloaded_paths: list[str] = []
    total = len(urls)

    for idx, url in enumerate(urls, 1):
        try:
            resp = session.get(url, stream=True, timeout=30)
            resp.raise_for_status()
        except Exception as e:
            print(f"[{idx}/{total}] ERROR fetching {url}: {e}", file=sys.stderr)
            continue

        ext = guess_extension(resp.headers.get("Content-Type", ""))
        filename = f"{idx:03d}{ext}"
        dest = os.path.join(out_dir, filename)

        with open(dest, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        downloaded_paths.append(dest)
        print(f"[{idx}/{total}] Saved {dest}")

    return downloaded_paths


def merge_docs(doc_paths: list[str], output_path: str) -> None:
    """
    Merge all DOCX files in doc_paths (in order) into a single DOCX file
    using docxcompose for robust merging. Non-DOCX files are ignored.
    """
    docx_paths = [p for p in doc_paths if p.lower().endswith(".docx")]

    if not docx_paths:
        print("No .docx files to merge; skipping merge.", file=sys.stderr)
        return

    print(f"Merging {len(docx_paths)} .docx files into {output_path} ...")

    # Use the first document as the base
    master = Document(docx_paths[0])
    composer = Composer(master)

    # Append the rest in order
    for path in docx_paths[1:]:
        doc = Document(path)
        composer.append(doc)

    composer.save(output_path)
    print(f"Merged document saved to {output_path}")



def main():
    parser = argparse.ArgumentParser(
        description=(
            "Download all 'Microsoft Word version of this ...' files "
            "linked from the Criminal Charge Book PDF and merge them."
        )
    )
    parser.add_argument("pdf", help="Path to ccb.pdf")
    parser.add_argument("out_dir", help="Directory to save downloaded Word files")
    parser.add_argument(
        "--merged",
        default="combined.docx",
        help="Filename for the merged Word document (inside out_dir)",
    )
    args = parser.parse_args()

    urls = extract_word_urls(args.pdf)
    print(f"Found {len(urls)} Word download links.")
    if not urls:
        return

    downloaded = download_all(urls, args.out_dir)

    if downloaded:
        merged_path = os.path.join(args.out_dir, args.merged)
        merge_docs(downloaded, merged_path)


if __name__ == "__main__":
    main()
