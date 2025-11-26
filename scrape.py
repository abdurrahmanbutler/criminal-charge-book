from pathlib import Path
from urllib.parse import urlparse
import time
import re

import requests
from playwright.sync_api import sync_playwright

START_URL = "https://resources.judicialcollege.vic.edu.au/article/1053858"

# Where to save the Word files
OUTPUT_DIR = Path("ccb_word_versions_playwright")
OUTPUT_DIR.mkdir(exist_ok=True)

# Be polite not to hammer the site
PER_PAGE_DELAY = 0.5

# Restrict to these domains when crawling article links
ALLOWED_ARTICLE_DOMAINS = {
    "resources.judicialcollege.vic.edu.au",
}

# Word files themselves live on this host
ALLOWED_WORD_DOMAINS = {
    "judicialcollege.vic.edu.au",
    "www.judicialcollege.vic.edu.au",
}

session = requests.Session()
session.headers["User-Agent"] = "VictorianCCB-WordDownloader/1.0 (personal use)"


def is_allowed_article_url(url: str) -> bool:
    host = urlparse(url).netloc.lower()
    return host in ALLOWED_ARTICLE_DOMAINS


def is_allowed_word_url(url: str) -> bool:
    host = urlparse(url).netloc.lower()
    return host in ALLOWED_WORD_DOMAINS


def guess_filename_from_url_or_headers(resp: requests.Response, url: str) -> str:
    cd = resp.headers.get("Content-Disposition") or resp.headers.get("content-disposition", "")
    filename = None

    if "filename=" in cd:
        _, _, tail = cd.partition("filename=")
        filename = tail.strip().strip('";\'')
        filename = filename.split("/")[-1]

    if not filename:
        path = urlparse(url).path
        last = Path(path).name
        if last and last not in {"file", ""}:
            filename = last
        else:
            m = re.search(r"/media/(\d+)/", path)
            if m:
                filename = f"ccb_{m.group(1)}.docx"
            else:
                filename = "ccb_document.docx"

    return filename


def collect_article_urls(page) -> set[str]:
    """
    Look through all frames on the page and collect links that look like
    article URLs (e.g. /article/10538xx, etc.)
    """
    article_urls: set[str] = set()

    for frame in page.frames:
        # This selector is deliberately broad: any link containing '/article/'
        # You can tighten it later if needed.
        try:
            urls = frame.eval_on_selector_all(
                "a[href*='/article/']",
                "els => els.map(a => a.href)"
            )
        except Exception:
            continue

        for u in urls:
            if is_allowed_article_url(u):
                article_urls.add(u)

    return article_urls


def collect_word_urls_from_current_page(page) -> set[str]:
    """
    In the *currently loaded article page*, look through all frames and
    collect any <a> elements that appear to be 'Download a Microsoft Word version...'
    links.
    """
    word_urls: set[str] = set()

    for frame in page.frames:
        try:
            urls = frame.eval_on_selector_all(
                "a[href]",
                """
                els => els
                    .filter(a => {
                        const text = (a.textContent || "").toLowerCase();
                        const href = (a.getAttribute("href") || "").toLowerCase();
                        // Heuristics based on your example snippet
                        return text.includes("microsoft word version")
                               || text.includes("word version")
                               || (href.includes("/media/") && href.endsWith("/file"));
                    })
                    .map(a => a.href);
                """
            )
        except Exception:
            continue

        for u in urls:
            if is_allowed_word_url(u):
                word_urls.add(u)

    return word_urls


def download_word_file(url: str):
    print(f"   [download] {url}")
    try:
        with session.get(url, stream=True, timeout=60) as resp:
            resp.raise_for_status()
            filename = guess_filename_from_url_or_headers(resp, url)
            out_path = OUTPUT_DIR / filename

            base = out_path
            counter = 1
            while out_path.exists():
                out_path = base.with_name(f"{base.stem}_{counter}{base.suffix}")
                counter += 1

            with open(out_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)

        print(f"      -> saved to {out_path}")
    except Exception as e:
        print(f"      !! failed to download {url}: {e}")


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # set True once it's working
        context = browser.new_context()
        page = context.new_page()

        print(f"[open] {START_URL}")
        page.goto(START_URL, wait_until="networkidle")
        time.sleep(1.0)

        # 1) Grab all the article URLs we can see from this page (across frames)
        article_urls = collect_article_urls(page)
        print(f"[info] found {len(article_urls)} article URLs from the main page")

        # Safety net: always include the start URL in case it also has a Word link
        article_urls.add(START_URL)

        all_word_urls: set[str] = set()

        # 2) Visit each article URL and collect Word download links
        for idx, article_url in enumerate(sorted(article_urls), start=1):
            print(f"[article {idx}/{len(article_urls)}] {article_url}")
            page.goto(article_url, wait_until="networkidle")
            time.sleep(PER_PAGE_DELAY)

            word_urls_here = collect_word_urls_from_current_page(page)
            print(f"   -> found {len(word_urls_here)} word link(s) on this page")
            all_word_urls.update(word_urls_here)

        browser.close()

    print(f"[summary] total unique Word URLs found: {len(all_word_urls)}")

    # 3) Download all the Word files we discovered
    for url in sorted(all_word_urls):
        download_word_file(url)


if __name__ == "__main__":
    main()
