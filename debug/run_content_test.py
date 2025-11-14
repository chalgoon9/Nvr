import os
from pathlib import Path

from bs4 import BeautifulSoup, Tag

# Self-contained constants mirroring main script behavior
REPO_ROOT = Path(__file__).resolve().parents[1]
DUMP_DIR = REPO_ROOT / "debug" / "content_outputs"

BRANDING_IMAGE_URLS = {
    "https://axh2eqadoldy.compat.objectstorage.ap-chuncheon-1.oraclecloud.com/bucket-20230610-0005/upload/top.png",
    "https://axh2eqadoldy.compat.objectstorage.ap-chuncheon-1.oraclecloud.com/bucket-20230610-0005/upload/bottom.png",
    "https://coudae.s3.ap-northeast-2.amazonaws.com/A00412936/cloud/7290.png",
}

GRAY_LINE_REPLACEMENTS = {
    "https://rapid-up.s3.ap-northeast-2.amazonaws.com/dev/gray-line.png":
        "https://axh2eqadoldy.compat.objectstorage.ap-chuncheon-1.oraclecloud.com/bucket-20230610-0005/upload/gray-line.png"
}

BLOCKED_IMAGE_PREFIXES = (
    "https://rapid-up.s3.ap-northeast-2.amazonaws.com",
    "https://cdn.heyseller.kr",
    "https://ai.esmplus.com/",
)

WRAP_CONTENT_HTML = (os.getenv("WRAP_CONTENT_HTML", "0").lower() in {"1","true","yes"})


def dump_content_html(product_code: str, html_text: str, label: str) -> None:
    if not html_text:
        return
    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    ts = __import__("time").strftime("%Y%m%d_%H%M%S")
    path = DUMP_DIR / f"{product_code}_{label}_{ts}.html"
    path.write_text(html_text, encoding="utf-8")
    # Root mirrors for quick IDE comparison
    if label == "html4":
        (REPO_ROOT / "html4.txt").write_text(html_text, encoding="utf-8")
    else:
        (REPO_ROOT / "output_html.txt").write_text(html_text, encoding="utf-8")


def build_html4_compatible_html_from_raw(raw_html: str, product_code: str = "TEST") -> str | None:
    if not raw_html:
        return None
    soup = BeautifulSoup(raw_html, "html.parser")
    # Normalize lazy images
    for img in soup.find_all("img", attrs={"data-src": True}):
        img["src"] = img["data-src"]
        try:
            del img["data-src"]
        except Exception:
            pass
    filtered = BeautifulSoup("", "html.parser")
    container = filtered.new_tag("div")
    filtered.append(container)
    seen_imgs: set[str] = set()
    allowed_headings = {"h1","h2","h3"}
    for node in soup.descendants:
        if not isinstance(node, Tag):
            continue
        name = node.name
        if name == "img":
            src = (node.get("src") or "").strip()
            if not src:
                continue
            src = GRAY_LINE_REPLACEMENTS.get(src, src)
            if any(src.startswith(pref) for pref in BLOCKED_IMAGE_PREFIXES):
                continue
            if src in BRANDING_IMAGE_URLS:
                continue
            if src in seen_imgs:
                continue
            seen_imgs.add(src)
            img = filtered.new_tag("img", src=src)
            img["style"] = "display:block;margin:0 auto 10px auto;"
            container.append(img)
            continue
        if name in allowed_headings:
            text = node.get_text(" ", strip=True)
            if text:
                h = filtered.new_tag(name)
                h.string = text
                container.append(h)
            continue
        if name == "p":
            text = node.get_text(" ", strip=True)
            if text:
                p = filtered.new_tag("p")
                p.string = text
                container.append(p)
            continue
    if not container.contents:
        return None
    return str(filtered)


def build_cleaned_variant_from_raw(raw_html: str, product_code: str = "TEST") -> str:
    if not raw_html:
        return ""

    soup = BeautifulSoup(raw_html, "html.parser")

    # Normalize lazy images
    for img in soup.find_all("img", attrs={"data-src": True}):
        img["src"] = img["data-src"]
        try:
            del img["data-src"]
        except Exception:
            pass

    # Build a minimal variant while preserving document order, similar to main pipeline
    minimal = BeautifulSoup("", "html.parser")
    parts: list[str] = []
    allowed_headings = {"h1","h2","h3"}
    seen_imgs: set[str] = set()
    for node in soup.descendants:
        if not isinstance(node, Tag):
            continue
        name = node.name
        if name == "img":
            src = (node.get("src") or "").strip()
            if not src:
                continue
            src = GRAY_LINE_REPLACEMENTS.get(src, src)
            if any(src.startswith(pref) for pref in BLOCKED_IMAGE_PREFIXES):
                continue
            if src in BRANDING_IMAGE_URLS:
                continue
            if src in seen_imgs:
                continue
            seen_imgs.add(src)
            img = minimal.new_tag("img", src=src)
            img["style"] = "display:block;margin:0 auto 10px auto;"
            parts.append(str(img))
            parts.append("<br>")
        elif name in allowed_headings:
            text = node.get_text(" ", strip=True)
            if text:
                h = minimal.new_tag(name)
                h.string = text
                parts.append(str(h))
        elif name == "p":
            text = node.get_text(" ", strip=True)
            if text:
                p = minimal.new_tag("p")
                p.string = text
                parts.append(str(p))
    final_html = "".join(parts)
    if WRAP_CONTENT_HTML:
        final_html = (
            "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body>"
            + final_html +
            "</body></html>"
        )
    return final_html


def main():
    # Honor env var or default to a bundled sample
    sample_path = (REPO_ROOT / Path(os.getenv("CONTENT_TEST_FILE", "debug/product_sample.html"))).resolve()
    product_code = os.getenv("CONTENT_TEST_CODE", "TEST")
    if not sample_path.exists():
        print(f"Sample file not found: {sample_path}")
        return 1

    raw_html = sample_path.read_text(encoding="utf-8", errors="ignore")

    # Always dump both to compare
    os.environ["DUMP_CONTENT_HTML"] = os.environ.get("DUMP_CONTENT_HTML", "1") or "1"

    html4_html = build_html4_compatible_html_from_raw(raw_html, product_code)
    if html4_html:
        dump_content_html(product_code, html4_html, "html4")

    # For parity with main pipeline’s comparison mode, store cleaned as the html4-ordered variant
    cleaned_html = html4_html or build_cleaned_variant_from_raw(raw_html, product_code)
    if cleaned_html:
        dump_content_html(product_code, cleaned_html, "cleaned")

    print("Saved cleaned and html4 variants. See debug/content_outputs and root html files.")
    print(f" - Root mirror cleaned: output_html.txt")
    print(f" - Root mirror html4  : html4.txt")
    print(f" - Index log         : debug/content_outputs/_index.log")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
