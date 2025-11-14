import os
import time
from pathlib import Path
from bs4 import BeautifulSoup, Tag

REPO_ROOT = Path(__file__).resolve().parents[1]
DUMP_DIR = REPO_ROOT / "debug" / "content_outputs"


def dump(product_code: str, html_text: str, label: str) -> None:
    if not html_text:
        return
    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    path = DUMP_DIR / f"{product_code}_{label}_{ts}.html"
    path.write_text(html_text, encoding="utf-8")
    if label == "html4":
        (REPO_ROOT / "html4.txt").write_text(html_text, encoding="utf-8")
    else:
        (REPO_ROOT / "output_html.txt").write_text(html_text, encoding="utf-8")


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


def build_html4_with_table(raw_html: str) -> str | None:
    if not raw_html:
        return None
    soup = BeautifulSoup(raw_html, "html.parser")
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
    allowed_headings = {"h1", "h2", "h3"}
    def simplify_table(table_node: Tag):
        tbl = filtered.new_tag("table", style="width:100%;border-collapse:collapse;margin:0 auto;")
        for tr in table_node.find_all("tr", recursive=True):
            new_tr = filtered.new_tag("tr")
            for td in tr.find_all(["td","th"], recursive=False):
                new_td = filtered.new_tag("td", style="width:50%;vertical-align:top;text-align:center;border:1px solid #000;padding:4px;")
                for img in td.find_all("img"):
                    src = (img.get("src") or img.get("data-src") or "").strip()
                    if not src:
                        continue
                    src = GRAY_LINE_REPLACEMENTS.get(src, src)
                    if any(src.startswith(pref) for pref in BLOCKED_IMAGE_PREFIXES):
                        continue
                    if src in BRANDING_IMAGE_URLS:
                        continue
                    if src in seen_imgs:
                        pass
                    seen_imgs.add(src)
                    clean = filtered.new_tag("img", src=src)
                    clean["style"] = "display:block;margin:0 auto 10px auto;"
                    new_td.append(clean)
                text = td.get_text(" ", strip=True)
                if text:
                    p = filtered.new_tag("p"); p.string = text; new_td.append(p)
                new_tr.append(new_td)
            if new_tr.find(["td","th"]):
                tbl.append(new_tr)
        return tbl
    handled_tables: set[Tag] = set()
    for node in soup.descendants:
        if not isinstance(node, Tag):
            continue
        name = node.name
        if name == "table":
            classes = " ".join(node.get("class", [])).lower()
            if "se-table-content" in classes or "se-table" in classes:
                tbl = simplify_table(node)
                container.append(tbl)
                handled_tables.add(node)
                continue
        parent_table = node.find_parent("table")
        if parent_table is not None and parent_table in handled_tables:
            continue
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
            clean_img = filtered.new_tag("img", src=src)
            clean_img["style"] = "display:block;margin:0 auto 10px auto;"
            container.append(clean_img)
            continue
        if name in allowed_headings:
            text = node.get_text(" ", strip=True)
            if text:
                h = filtered.new_tag(name); h.string = text; container.append(h)
            continue
        if name == "p":
            text = node.get_text(" ", strip=True)
            if text:
                p = filtered.new_tag("p"); p.string = text; container.append(p)
            continue
    if not container.contents:
        return None
    return str(filtered)


def main():
    path = Path(os.getenv("FILE", "debug/tile_table_sample.html"))
    code = os.getenv("CODE", "TABLETEST")
    if not path.exists():
        print(f"File not found: {path}")
        return 1
    raw = path.read_text(encoding="utf-8", errors="ignore")
    html4 = build_html4_with_table(raw)
    if html4:
        dump(code, html4, "html4"); dump(code, html4, "cleaned")
        print("Saved html4 and cleaned (identical) outputs from file.")
        return 0
    print("No content built."); return 2


if __name__ == "__main__":
    raise SystemExit(main())

