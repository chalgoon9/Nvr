import sys
import os
from pathlib import Path
import csv
import re

try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None


def summarize_html(text: str):
    if not text:
        return {"img_count": 0, "heading_count": 0, "text_block_count": 0, "has_wrapper": False}
    if BeautifulSoup is None:
        t = text.lower()
        return {
            "img_count": t.count("<img"),
            "heading_count": t.count("<h1") + t.count("<h2") + t.count("<h3"),
            "text_block_count": t.count("<p"),
            "has_wrapper": ("<html" in t) or ("<head" in t),
        }
    soup = BeautifulSoup(text, "html.parser")
    img_count = len(soup.find_all("img"))
    heading_count = sum(len(soup.find_all(h)) for h in ("h1", "h2", "h3"))
    text_block_count = len(soup.find_all("p"))
    has_wrapper = bool(soup.find("html") or soup.find("head"))
    return {
        "img_count": img_count,
        "heading_count": heading_count,
        "text_block_count": text_block_count,
        "has_wrapper": has_wrapper,
    }


def main():
    # Usage: python debug/index_output_html.py [dir]
    # Default to env OUTPUT_HTML_DIR or repo debug/content_outputs
    base = (
        Path(sys.argv[1])
        if len(sys.argv) > 1
        else Path(os.getenv("OUTPUT_HTML_DIR", "debug/content_outputs"))
    )
    base = base.resolve()
    if not base.exists():
        print(f"Target folder not found: {base}")
        sys.exit(1)

    files = sorted(base.rglob("*.html"))
    if not files:
        print(f"No HTML files under: {base}")
        return

    rows = []
    for f in files:
        try:
            text = f.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        s = summarize_html(text)
        name = f.name
        m_code = re.search(r"(\d+)", name)
        m_label = re.search(r"_(cleaned|html4|fallback_gallery)_", name)
        rows.append(
            {
                "file": str(f),
                "name": name,
                "product_code": m_code.group(1) if m_code else "",
                "label": m_label.group(1) if m_label else "",
                **s,
            }
        )

    out_csv = base / "index_summary.csv"
    with out_csv.open("w", newline="", encoding="utf-8") as fp:
        w = csv.DictWriter(
            fp,
            fieldnames=[
                "file",
                "name",
                "product_code",
                "label",
                "img_count",
                "heading_count",
                "text_block_count",
                "has_wrapper",
            ],
        )
        w.writeheader()
        for r in rows:
            w.writerow(r)
    print(f"Wrote summary: {out_csv}")

    # Simple console top-10 preview by modified time
    latest = sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)[:10]
    print("Latest files:")
    for p in latest:
        print(" -", p)


if __name__ == "__main__":
    main()

