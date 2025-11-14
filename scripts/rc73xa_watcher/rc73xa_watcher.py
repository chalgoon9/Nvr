#!/usr/bin/env python3
# RC73XA Coupang Return watcher (Ubuntu/server friendly)
# - Headless Firefox via Playwright
# - Telegram via env: TG_BOT_TOKEN, TG_CHAT_ID
# - Optional heartbeat and /status command

from playwright.sync_api import sync_playwright
from urllib.parse import urljoin, quote_plus
import os, re, time, json, urllib.request, urllib.parse
from datetime import datetime, timedelta

try:
    from dotenv import load_dotenv  # optional
    load_dotenv()
except Exception:
    pass


# ===== Settings (env overrides) =====
TARGET_MODEL = os.getenv("TARGET_MODEL", r"rc73xa\s*-?\s*nh011w")
EXCLUDES = json.loads(os.getenv("EXCLUDES", "[\"rc71l\\s*-?\\s*nh001w\"]"))
RETURN_KEYWORDS = json.loads(os.getenv("RETURN_KEYWORDS", "[\"\\uBC18\\uD488\", \"\\uB9AC\\uD37C\"]"))

QUERY = os.getenv("QUERY", "RC73XA-NH011W")
SEARCH_URL = f"https://www.coupang.com/np/search?q={quote_plus(QUERY)}"

TG_BOT_TOKEN = os.getenv("TG_BOT_TOKEN", "").strip()
TG_CHAT_ID = os.getenv("TG_CHAT_ID", "").strip()
TG_COMMANDS_ENABLED = os.getenv("TG_COMMANDS_ENABLED", "1").lower() in {"1", "true", "yes"}
TG_ALLOWED_CHAT_ID = os.getenv("TG_ALLOWED_CHAT_ID", TG_CHAT_ID).strip()
HEARTBEAT_MINUTES = int(os.getenv("HEARTBEAT_MINUTES", "360"))  # 6h default
TG_OFFSET_PATH = os.getenv("TG_OFFSET_PATH", os.path.join(os.path.dirname(__file__), "tg_offset.json"))

STATE_FILE = os.getenv("STATE_FILE", os.path.join(os.path.dirname(__file__), "seen_links.json"))
MIN_INTERVAL = int(os.getenv("MIN_INTERVAL", "3"))
MAX_INTERVAL = int(os.getenv("MAX_INTERVAL", "7"))
MAX_RETURN_PRICE = int(os.getenv("MAX_RETURN_PRICE", str(1_000_000)))
HEADLESS = os.getenv("HEADLESS", "1").lower() not in {"0", "false", "no"}

UA_DESK = (
    os.getenv(
        "USER_AGENT",
        "Mozilla/5.0 (X11; Linux x86_64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36",
    )
)


# ===== Utils =====
def log(*a):
    print(datetime.now().strftime("[%Y-%m-%d %H:%M:%S]"), *a, flush=True)


def random_interval_minutes():
    span = max(0, MAX_INTERVAL - MIN_INTERVAL)
    b = int.from_bytes(os.urandom(1), "big")
    return MIN_INTERVAL + (b * span) // 255


def clean(t):
    return re.sub(r"\s+", " ", (t or "").strip())


def name_matches_target(name: str) -> bool:
    n = (name or "").lower()
    if not re.search(TARGET_MODEL, n, re.I):
        return False
    for pat in EXCLUDES:
        if re.search(pat, n, re.I):
            return False
    return True


def is_return_item(text: str) -> bool:
    t = (text or "").lower()
    return any(kw.lower() in t for kw in RETURN_KEYWORDS)


def parse_price(text: str):
    digits = re.sub(r"[^\d]", "", text or "")
    if not digits:
        return None
    try:
        return int(digits)
    except ValueError:
        return None


def wait_stable(page, timeout_ms=30000):
    page.wait_for_load_state("domcontentloaded", timeout=timeout_ms)
    try:
        page.wait_for_function("document.readyState === 'complete'", timeout=timeout_ms)
    except Exception:
        pass
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms // 2)
    except Exception:
        pass


def soft_scroll(page, steps=10, pause=0.25):
    for _ in range(steps):
        try:
            page.mouse.wheel(0, 1200)
        except Exception:
            wait_stable(page, 5000)
        time.sleep(pause)


def collect_candidates(page, limit=80):
    items, seen = [], set()
    anchors = page.locator("a[href*='/vp/products/']")
    cnt = min(anchors.count(), limit * 3)
    for i in range(cnt):
        a = anchors.nth(i)
        href = a.get_attribute("href") or ""
        if "/vp/products/" not in href:
            continue
        link = urljoin("https://www.coupang.com", href)
        if link in seen:
            continue
        seen.add(link)

        name = clean(a.inner_text()) or clean(
            a.get_attribute("aria-label") or a.get_attribute("title") or ""
        )
        parent = a
        if not name:
            for _ in range(5):
                parent = parent.parent_element()
                if not parent:
                    break
                cand = parent.locator(
                    "div.ProductUnit_productNameV2__cV9cw, div.name, .name, .title"
                )
                if cand.count():
                    name = clean(cand.first.inner_text())
                    if name:
                        break
        if name:
            context = name
            if parent:
                try:
                    context = clean(parent.inner_text()) or context
                except Exception:
                    pass
            items.append({"name": name, "link": link, "context": context})
        if len(items) >= limit:
            break
    return items


def fetch_return_price(ctx, link: str, timeout_ms=60000):
    detail = ctx.new_page()
    try:
        detail.goto(link, wait_until="load", timeout=timeout_ms)
        wait_stable(detail, timeout_ms)
        price_text = ""
        selectors = [
            "xpath=(//span[contains(normalize-space(.), '최저')]//strong)[1]",
            "xpath=(//span[contains(normalize-space(.), '최저')])[1]",
        ]
        for selector in selectors:
            try:
                loc = detail.locator(selector)
                if loc.count():
                    price_text = clean(loc.first.inner_text())
                    if price_text:
                        break
            except Exception:
                continue
        price = parse_price(price_text)
        if price is not None:
            return price
        try:
            body_text = detail.inner_text("body")
            for line in body_text.splitlines():
                if "최저" in line:
                    price = parse_price(line)
                    if price is not None:
                        return price
        except Exception:
            pass
        return None
    except Exception as e:
        log("[WARN] return price fetch failed:", e)
        return None
    finally:
        try:
            detail.close()
        except Exception:
            pass


# ===== Telegram helpers =====
def tg_send(text: str):
    if not TG_BOT_TOKEN or not TG_CHAT_ID:
        log("[TG] skipped (env TG_BOT_TOKEN/TG_CHAT_ID not set)")
        return False
    try:
        url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
        data = urllib.parse.urlencode(
            {"chat_id": TG_CHAT_ID, "text": text, "disable_web_page_preview": True}
        ).encode("utf-8")
        with urllib.request.urlopen(url, data=data, timeout=15) as r:
            ok = r.status == 200
        log("[TG]", "sent" if ok else f"status {r.status}")
        return ok
    except Exception as e:
        log("[TG] failed:", e)
        return False


def tg_get_updates(offset=None, timeout=0):
    if not TG_BOT_TOKEN or not TG_COMMANDS_ENABLED:
        return []
    try:
        url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/getUpdates"
        params = {"timeout": timeout}
        if offset is not None:
            params["offset"] = offset
        url = url + "?" + urllib.parse.urlencode(params)
        with urllib.request.urlopen(url, timeout=20) as r:
            data = json.loads(r.read().decode("utf-8"))
        return data.get("result", []) if data.get("ok") else []
    except Exception:
        return []


def tg_reply(chat_id, text):
    try:
        url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
        data = urllib.parse.urlencode(
            {"chat_id": chat_id, "text": text, "disable_web_page_preview": True}
        ).encode("utf-8")
        with urllib.request.urlopen(url, data=data, timeout=15) as r:
            return r.status == 200
    except Exception:
        return False


def tg_offset_load():
    try:
        if os.path.exists(TG_OFFSET_PATH):
            return json.load(open(TG_OFFSET_PATH, "r", encoding="utf-8")).get("offset")
    except Exception:
        pass
    return None


def tg_offset_save(offset):
    try:
        json.dump({"offset": offset}, open(TG_OFFSET_PATH, "w", encoding="utf-8"))
    except Exception:
        pass


# ===== State =====
def load_seen():
    try:
        if os.path.exists(STATE_FILE):
            return set(json.load(open(STATE_FILE, "r", encoding="utf-8")))
    except Exception:
        pass
    return set()


def save_seen(s: set):
    try:
        json.dump(sorted(s), open(STATE_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    except Exception as e:
        log("[WARN] save_seen failed:", e)


def check_once():
    with sync_playwright() as p:
        br = p.firefox.launch(headless=HEADLESS)
        ctx = br.new_context(
            locale="ko-KR",
            timezone_id="Asia/Seoul",
            ignore_https_errors=True,
            viewport={"width": 1366, "height": 900},
            extra_http_headers={"Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8"},
            user_agent=UA_DESK,
        )
        pg = ctx.new_page()

        log("[INFO] goto:", SEARCH_URL)
        pg.goto(SEARCH_URL, wait_until="load", timeout=120_000)
        wait_stable(pg, 30000)
        try:
            pg.wait_for_selector(
                "ul.search-product-list, .search-product, li:has(a[href*='/vp/products/'])",
                timeout=15000,
            )
        except Exception:
            pass
        soft_scroll(pg, steps=10, pause=0.25)

        cands = collect_candidates(pg, limit=80)
        log(f"[INFO] candidates: {len(cands)}")

        seen = load_seen()
        found = []
        for it in cands:
            if not name_matches_target(it["name"]):
                continue
            context = it.get("context") or it["name"]
            if not is_return_item(context):
                continue
            if it["link"] in seen:
                continue
            price = fetch_return_price(ctx, it["link"])
            if price is None:
                log("[INFO] 가격 확인 실패, 건너뜀:", it["link"])
                continue
            if price > MAX_RETURN_PRICE:
                log(f"[INFO] 상한 초과({price:,}원)로 건너뜀:", it["link"])
                continue
            it["price"] = price
            found.append(it)
            seen.add(it["link"])

        if found:
            for f in found:
                msg = (
                    f"✅ 타겟 품목 추가!\n{f['name']}\n반품가: {f['price']:,}원\n{f['link']}"
                )
                log(msg.replace("\n", " | "))
                tg_send(msg)
            save_seen(seen)
        else:
            log("[INFO] 타겟 품목 미발견")

        br.close()
        return len(found)


def process_tg_commands(start_time, last_cycle_ok, last_found):
    if not TG_COMMANDS_ENABLED or not TG_BOT_TOKEN:
        return None
    offset = tg_offset_load()
    updates = tg_get_updates(offset=offset, timeout=0)
    max_update_id = offset
    for upd in updates:
        uid = upd.get("update_id")
        msg = upd.get("message") or {}
        chat = msg.get("chat") or {}
        chat_id = str(chat.get("id")) if chat.get("id") is not None else None
        text = (msg.get("text") or "").strip()
        if uid is not None and (max_update_id is None or uid > max_update_id):
            max_update_id = uid
        if not text or not chat_id:
            continue
        if TG_ALLOWED_CHAT_ID and chat_id != str(TG_ALLOWED_CHAT_ID):
            continue
        if text.lower() in {"/status", "status", "/ping", "ping"}:
            up = datetime.now() - start_time
            last_ok = last_cycle_ok.isoformat(timespec="seconds") if last_cycle_ok else "N/A"
            last_f = last_found.isoformat(timespec="seconds") if last_found else "N/A"
            reply = (
                "RC73XA watcher OK\n"
                f"uptime: {str(up).split('.')[0]}\n"
                f"last cycle: {last_ok}\n"
                f"last found: {last_f}\n"
                f"interval: {MIN_INTERVAL}-{MAX_INTERVAL} min"
            )
            tg_reply(chat_id, reply)
    if max_update_id is not None:
        tg_offset_save(max_update_id + 1)


def main_loop():
    log("[START] RC73XA watcher (every 3–7 min). Ctrl+C to stop.")
    start_time = datetime.now()
    last_cycle_ok = None
    last_found = None
    last_heartbeat = datetime.min
    while True:
        try:
            n = check_once()
            last_cycle_ok = datetime.now()
            if n and n > 0:
                last_found = last_cycle_ok
        except Exception as e:
            log("[ERROR]", e)
        if TG_COMMANDS_ENABLED:
            try:
                process_tg_commands(start_time, last_cycle_ok, last_found)
            except Exception:
                pass
        if HEARTBEAT_MINUTES > 0 and (datetime.now() - last_heartbeat) >= timedelta(minutes=HEARTBEAT_MINUTES):
            hb = (
                f"[HB] alive. uptime={str(datetime.now()-start_time).split('.')[0]} "
                f"last_cycle={last_cycle_ok.isoformat(timespec='seconds') if last_cycle_ok else 'N/A'}"
            )
            tg_send(hb)
            last_heartbeat = datetime.now()
        mins = random_interval_minutes()
        log(f"[SLEEP] {mins}분 대기")
        try:
            time.sleep(mins * 60)
        except KeyboardInterrupt:
            log("[EXIT] stopped by user")
            break


if __name__ == "__main__":
    main_loop()

