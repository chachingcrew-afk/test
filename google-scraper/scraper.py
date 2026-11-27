#!/usr/bin/env python3
"""
Google.nl Scraper – Organic + Shopping
Hollandia, Playwright + Webshare + Supabase
"""

import asyncio
import json
import os
import random
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd
from playwright.async_api import async_playwright, Browser, BrowserContext, Page
from playwright_stealth import stealth_async
from supabase import create_client, Client

CONFIG_PATH = Path("config.json")
DEBUG_DIR = Path("debug")

# Google EU consent megkerülése ("Alles afwijzen" állapot)
SOCS_COOKIE = "CAESHAgBEhJnd3NfMjAyMzA4MTAtMF9SQzIaAmVuIAEaBgiAo_CmBg"

CAPTCHA_MARKERS = [
    "unusual traffic",
    "ongebruikelijk verkeer",     # holland Google captcha szöveg
    "id=\"captcha-form\"",
    "g-recaptcha",
]

MAX_CONSECUTIVE_CAPTCHA = 3


# ---------------------------------------------------------------- config

def load_config() -> dict:
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def get_supabase(config: dict) -> Client:
    url = os.getenv("SUPABASE_URL") or config.get("supabase", {}).get("url")
    key = os.getenv("SUPABASE_KEY") or config.get("supabase", {}).get("key")
    if not url or not key:
        raise ValueError("SUPABASE_URL és SUPABASE_KEY hiányzik")
    return create_client(url, key)


def normalize_ean(value) -> str:
    """Excelből jövő EAN tisztítása (float, '.0', tudományos forma, szóközök)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if not s or s.lower() == "nan":
        return ""
    if re.fullmatch(r"\d+(\.\d+)?[eE]\+?\d+", s):          # 8.71234E+12
        s = str(int(float(s)))
    s = re.sub(r"\.0+$", "", s)                              # 8712345678901.0
    s = re.sub(r"\D", "", s)
    return s


COLUMN_ALIASES = {
    "code": ["Code", "code", "SKU", "sku"],
    "name": ["Product", "product", "name", "Name"],
    "ean": ["EAN", "ean", "Ean", "barcode"],
    "brand": ["Merk", "merk", "brand", "Brand"],
}


def _col(row, key):
    for c in COLUMN_ALIASES[key]:
        if c in row.index and pd.notna(row[c]):
            return row[c]
    return None


def resolve_products_file(setting) -> str:
    candidates = setting if isinstance(setting, list) else [setting]
    for c in candidates:
        if Path(c).exists():
            return c
    raise FileNotFoundError(f"Nincs meg egyik termékfájl sem: {candidates}")


def load_products(path: str) -> list[dict]:
    # dtype=str: különben a pandas floatként olvassa az EAN-t
    df = pd.read_excel(path, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    products, seen = [], set()
    for _, row in df.iterrows():
        ean = normalize_ean(_col(row, "ean"))
        if len(ean) < 8 or ean in seen:
            continue
        seen.add(ean)
        products.append({
            "code": str(_col(row, "code") or "").strip(),
            "name": str(_col(row, "name") or "").strip() or ean,
            "ean": ean,
            "brand": str(_col(row, "brand") or "").strip(),
        })
    return products


def get_already_scraped_eans(supabase: Client, hours: int = 40) -> set[str]:
    """Az utolsó X órában sikeresen scrapelt EAN-ek (lapozva, mert a Supabase 1000 soronként ad vissza)."""
    since = (datetime.now(timezone.utc) - timedelta(hours=hours)).isoformat()
    eans: set[str] = set()
    page_size, start = 1000, 0
    try:
        while True:
            res = (
                supabase.table("google_searches")
                .select("ean")
                .gte("scraped_at", since)
                .in_("status", ["ok", "no_results"])
                .order("id")
                .range(start, start + page_size - 1)
                .execute()
            )
            eans.update(r["ean"] for r in res.data if r.get("ean"))
            if len(res.data) < page_size:
                break
            start += page_size
    except Exception as e:
        print(f"Figyelem: nem sikerült lekérni a korábbi kereséseket: {e}")
    return eans


# ---------------------------------------------------------------- helpers

def clean_price(text: str) -> Optional[float]:
    if not text:
        return None
    m = re.search(r"\d[\d.\s]*(,\d{1,2})?", text)
    if not m:
        return None
    cleaned = m.group(0).replace(" ", "")
    # Holland formátum: 1.234,56 vagy 12,95
    if "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


async def is_captcha(page: Page) -> bool:
    if "/sorry/" in page.url:
        return True
    content = (await page.content()).lower()
    return any(m in content for m in CAPTCHA_MARKERS)


async def handle_consent(page: Page):
    """Ha a cookie ellenére a consent oldalra dob, nyomjuk meg az elutasítást."""
    if "consent.google" not in page.url:
        return
    for label in ["Alles afwijzen", "Reject all", "Alles accepteren", "Accept all"]:
        btn = page.get_by_role("button", name=label)
        if await btn.count():
            await btn.first.click()
            await page.wait_for_load_state("domcontentloaded")
            await asyncio.sleep(1)
            return


async def dump_debug(page: Page, ean: str, tag: str):
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        base = DEBUG_DIR / f"{ean}_{tag}"
        (base.with_suffix(".html")).write_text(await page.content(), encoding="utf-8")
        await page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
    except Exception as e:
        print(f"     debug mentés hiba: {e}")


def build_proxy(proxy_cfg: dict) -> dict:
    """Webshare username: {username}-{country}-rotate (ugyanúgy, mint a simple-scraperben)."""
    user = (os.getenv("WEBSHARE_USER") or proxy_cfg.get("username") or "").strip()
    password = (os.getenv("WEBSHARE_PASS") or proxy_cfg.get("password") or "").strip()
    if not user or not password:
        raise SystemExit("HIBA: nincs WEBSHARE_USER / WEBSHARE_PASS")
    country = str(proxy_cfg.get("country") or "").strip().lower()
    base = re.sub(r"-rotate$", "", user)
    if country:
        base = re.sub(rf"-{re.escape(country)}$", "", base, flags=re.I)
    user = base + (f"-{country}" if country else "") + ("-rotate" if proxy_cfg.get("rotate", True) else "")
    host = proxy_cfg.get("host", "p.webshare.io")
    port = proxy_cfg.get("port", 80)
    return {"server": f"http://{host}:{port}", "username": user, "password": password}


async def create_context(browser: Browser, config: dict) -> BrowserContext:
    proxy_cfg = config.get("proxy", {})
    settings = config.get("settings", {})

    context_options = {
        "locale": settings.get("locale", "nl-NL"),
        "timezone_id": settings.get("timezone_id", "Europe/Amsterdam"),
        "viewport": {"width": 1366, "height": 768},
        "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    }

    if proxy_cfg.get("enabled"):
        context_options["proxy"] = build_proxy(proxy_cfg)

    context = await browser.new_context(**context_options)

    # Proxy adatforgalom spórolás: képek, fontok, videók nem kellenek
    async def block_heavy(route):
        if route.request.resource_type in ("image", "media", "font"):
            await route.abort()
        else:
            await route.continue_()
    await context.route("**/*", block_heavy)

    await context.add_cookies([
        {"name": "SOCS", "value": SOCS_COOKIE, "domain": ".google.nl", "path": "/"},
        {"name": "CONSENT", "value": "PENDING+987", "domain": ".google.nl", "path": "/"},
    ])
    return context


# ---------------------------------------------------------------- extract

ORGANIC_JS = r"""
(max) => {
  const out = [], seen = new Set();
  const anchors = document.querySelectorAll('#rso a:has(h3), #search a:has(h3)');
  for (const a of anchors) {
    const href = a.href || '';
    if (!href.startsWith('http')) continue;
    try { if (new URL(href).hostname.includes('google.')) continue; } catch (e) { continue; }
    if (seen.has(href)) continue;
    seen.add(href);
    const block = a.closest('div.MjjYud, div.g, div[data-hveid]');
    const sn = block && block.querySelector('div.VwiC3b, div[data-sncf], div[style*="-webkit-line-clamp"]');
    out.push({
      title: a.querySelector('h3').innerText.trim(),
      link: href,
      snippet: sn ? sn.innerText.trim() : null,
    });
    if (out.length >= max) break;
  }
  return out;
}
"""

# Google Shopping DOM-ja gyakran változik: először ismert kártya-szelektorok,
# ha nincs találat, heurisztika: ár-szöveget tartalmazó, linkes, rövid blokkok.
SHOPPING_JS = r"""
(max) => {
  const priceRe = /€\s?\d[\d.]*(,\d{2})?|\d[\d.]*,\d{2}\s?€/;
  const txt = el => (el && el.innerText || '').trim();
  const first = (root, sels) => { for (const s of sels) { const e = root.querySelector(s); if (e && txt(e)) return e; } return null; };

  const parse = (card) => {
    const titleEl = first(card, ['h3', '[role="heading"]', '.tAxDx', '.Xjkr3b', '.gkQHve', '.EI11Pd']);
    let priceText = null;
    for (const s of card.querySelectorAll('span, b, div')) {
      if (s.children.length === 0 && priceRe.test(txt(s))) { priceText = txt(s); break; }
    }
    const merchantEl = first(card, ['.aULzUe', '.IuHnof', '.WggmJe', '.WJMUdc', '.zPEcBd']);
    const link = card.querySelector('a[href]');
    return {
      title: txt(titleEl) || null,
      price_text: priceText,
      merchant: txt(merchantEl) || null,
      link: link ? link.href : null,
    };
  };

  const known = 'div.sh-dgr__gr-auto, div.sh-dgr__content, div[data-docid], div.sh-pr__product-results-grid > div';
  let cards = [...document.querySelectorAll(known)];
  let mode = 'selector';

  if (!cards.length) {
    mode = 'heuristic';
    const found = new Set();
    for (const leaf of document.querySelectorAll('#search span, #rso span, main span')) {
      if (leaf.children.length || !priceRe.test(txt(leaf))) continue;
      let el = leaf;
      for (let i = 0; i < 8 && el; i++) {
        el = el.parentElement;
        if (!el) break;
        const t = txt(el);
        if (t.length > 500) { el = null; break; }
        if (el.querySelector('a[href]') && t.split('\n').length >= 2) break;
      }
      if (el) found.add(el);
    }
    // a beágyazott duplikátumok kiszűrése
    cards = [...found].filter(c => ![...found].some(o => o !== c && o.contains(c)));
  }

  const out = [];
  for (const c of cards) {
    const r = parse(c);
    if (!r.title) {
      const lines = txt(c).split('\n').map(s => s.trim()).filter(s => s && !priceRe.test(s));
      r.title = lines[0] || null;
    }
    if (r.title && r.price_text) out.push(r);
    if (out.length >= max) break;
  }
  return { mode, items: out };
}
"""


async def extract_organic(page: Page, max_items: int) -> list[dict]:
    raw = await page.evaluate(ORGANIC_JS, max_items)
    return [{
        "result_type": "organic",
        "position": i + 1,
        "title": r["title"],
        "price": None,
        "currency": None,
        "merchant": None,
        "link": r["link"],
        "snippet": r["snippet"],
    } for i, r in enumerate(raw)]


async def extract_shopping(page: Page, max_items: int) -> tuple[list[dict], str]:
    raw = await page.evaluate(SHOPPING_JS, max_items)
    items = [{
        "result_type": "shopping",
        "position": i + 1,
        "title": r["title"][:500],
        "price": clean_price(r["price_text"]),
        "currency": "EUR",
        "merchant": r["merchant"],
        "link": r["link"],
        "snippet": None,
    } for i, r in enumerate(raw["items"])]
    return items, raw["mode"]


# ---------------------------------------------------------------- scrape

async def goto_google(page: Page, url: str):
    await page.goto(url, wait_until="domcontentloaded", timeout=45000)
    await handle_consent(page)
    await asyncio.sleep(random.uniform(1.5, 3.5))


async def scrape_product(page: Page, product: dict, config: dict) -> dict:
    settings = config["settings"]
    debug = settings.get("debug", False)
    query = product["ean"]  # EAN a legpontosabb

    result = {
        "code": product["code"],
        "ean": product["ean"],
        "name": product["name"],
        "query": query,
        "status": "error",
        "error": None,
        "organic": [],
        "shopping": [],
    }

    try:
        # Organikus
        await goto_google(page, f"https://www.google.nl/search?q={query}&hl=nl&gl=nl")
        if await is_captcha(page):
            result["status"], result["error"] = "captcha", "CAPTCHA (organic)"
            return result

        result["organic"] = await extract_organic(page, settings.get("max_organic", 8))
        if debug and not result["organic"]:
            await dump_debug(page, query, "organic")

        # Shopping fül (udm=28 az új shopping nézet, a tbm=shop elavult)
        await goto_google(page, f"https://www.google.nl/search?q={query}&udm=28&hl=nl&gl=nl")
        if await is_captcha(page):
            result["status"], result["error"] = "captcha", "CAPTCHA (shopping)"
            return result

        shopping, mode = await extract_shopping(page, settings.get("max_shopping", 12))
        result["shopping"] = shopping
        if debug and (not shopping or mode == "heuristic"):
            await dump_debug(page, query, f"shopping_{mode}")

        result["status"] = "ok" if (result["organic"] or shopping) else "no_results"

    except Exception as e:
        result["error"] = str(e)[:500]
        result["status"] = "error"

    return result


def save_results(supabase: Client, data: dict):
    now = datetime.now(timezone.utc).isoformat()
    search_row = {
        "code": data["code"],
        "ean": data["ean"],
        "name": data["name"],
        "query": data["query"],
        "status": data["status"],
        "error": data.get("error"),
        "organic_count": len(data.get("organic", [])),
        "shopping_count": len(data.get("shopping", [])),
        "scraped_at": now,
    }
    res = supabase.table("google_searches").insert(search_row).execute()
    search_id = res.data[0]["id"]

    rows = [{
        "search_id": search_id,
        "ean": data["ean"],
        "result_type": item["result_type"],
        "position": item["position"],
        "title": item["title"],
        "price": item["price"],
        "currency": item.get("currency") or "EUR",
        "merchant": item.get("merchant"),
        "link": item.get("link"),
        "snippet": item.get("snippet"),
        "scraped_at": now,
    } for item in data.get("organic", []) + data.get("shopping", [])]

    if rows:
        supabase.table("google_results").insert(rows).execute()


# ---------------------------------------------------------------- main

async def main():
    config = load_config()
    settings = config["settings"]
    batch_size = int(os.getenv("BATCH_SIZE") or settings.get("batch_size", 40))
    skip_hours = settings.get("skip_hours", 40)

    print(f"[{datetime.now().isoformat()}] Google scraper indult")

    products_file = resolve_products_file(settings["products_file"])
    products = load_products(products_file)
    print(f"Összesen {len(products)} termék az Excelben ({products_file})")

    supabase = get_supabase(config)
    already = get_already_scraped_eans(supabase, hours=skip_hours)
    print(f"Már scrapelve (utolsó {skip_hours} óra): {len(already)} db")

    todo = [p for p in products if p["ean"] not in already][:batch_size]
    print(f"Mostani batch: {len(todo)} termék")
    if not todo:
        print("Nincs több teendő termék. Kilépés.")
        return

    if config.get("proxy", {}).get("enabled"):
        u = build_proxy(config["proxy"])["username"]
        print(f"Proxy: {u[:2]}***{u[u.find('-'):] if '-' in u else ''}")

    consecutive_captcha = 0
    stats = {"ok": 0, "no_results": 0, "captcha": 0, "error": 0}

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=settings.get("headless", True),
            args=["--disable-blink-features=AutomationControlled"],
        )

        for i, product in enumerate(todo, 1):
            print(f"\n[{i}/{len(todo)}] {product['name'][:60]} (EAN: {product['ean']})")

            context = await create_context(browser, config)
            page = await context.new_page()
            await stealth_async(page)

            data = await scrape_product(page, product, config)
            await context.close()

            status = data["status"]
            stats[status] = stats.get(status, 0) + 1
            print(f"  → {status.upper()} | organic: {len(data['organic'])} | shopping: {len(data['shopping'])}")
            if data.get("error"):
                print(f"     Hiba: {data['error']}")

            try:
                save_results(supabase, data)
                print("  → Mentve Supabase-be")
            except Exception as e:
                print(f"  → Mentési hiba: {e}")

            consecutive_captcha = consecutive_captcha + 1 if status == "captcha" else 0
            if consecutive_captcha >= MAX_CONSECUTIVE_CAPTCHA:
                print(f"\n{MAX_CONSECUTIVE_CAPTCHA} CAPTCHA egymás után, a batch leáll (a proxy IP-k le vannak tiltva).")
                break

            if i < len(todo):
                delay = random.uniform(
                    settings.get("min_delay_seconds", 9),
                    settings.get("max_delay_seconds", 22),
                )
                print(f"  Várakozás {delay:.1f} mp...")
                await asyncio.sleep(delay)

        await browser.close()

    print(f"\n[{datetime.now().isoformat()}] Batch kész. {stats}")


if __name__ == "__main__":
    asyncio.run(main())
