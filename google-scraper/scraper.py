#!/usr/bin/env python3
"""
Google.nl Shopping Scraper (EAN alapján)
Patchright + valódi Chrome + Webshare NL sticky residential proxy + Supabase

CAPTCHA csökkentés:
- patchright + Chrome (xvfb-n látható módban), nem headless Chromium
- egy holland IP (sticky session) több keresést szolgál ki, a cookie-k megmaradnak
- a session (IP + cookie) Supabase-be mentve (browser_sessions, domain = google.nl)
- termékenként csak egy keresés (shopping nézet), hosszú, véletlen szünetek
- CAPTCHA esetén új IP, nem áll le azonnal
"""

import asyncio
import json
import os
import random
import re
import unicodedata
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Optional

import pandas as pd
from supabase import create_client, Client

CONFIG_PATH = Path("config.json")
DEBUG_DIR = Path("debug")
SESSION_KEY = "google.nl"

SOCS_COOKIE = "CAESHAgBEhJnd3NfMjAyMzA4MTAtMF9SQzIaAmVuIAEaBgiAo_CmBg"
CAPTCHA_MARKERS = ["unusual traffic", "ongebruikelijk verkeer", 'id="captcha-form"', "g-recaptcha"]
BLOCKED_TYPES = {"image", "media", "font"}


# ---------------------------------------------------------------- config / input

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



def _norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "").encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


STOPWORDS = {"voor", "met", "van", "en", "de", "het", "een", "the", "and", "with", "stuks", "stuk", "set"}


def is_relevant(title: str, product: dict) -> bool:
    """A találat címe illik-e a termékhez (márka vagy a név fontos szavai)."""
    t = _norm(title)
    if not t:
        return False
    brand = _norm(product.get("brand", ""))
    if brand and len(brand) >= 3 and brand in t:
        return True
    words = [w for w in _norm(product.get("name", "")).split()
             if len(w) >= 4 and not w.isdigit() and w not in STOPWORDS]
    if not words:
        return False
    hits = sum(1 for w in words if w in t)
    return hits >= 2 or hits / len(words) >= 0.5


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



# ---------------------------------------------------------------- browser session

class GoogleSession:
    """Egy böngésző, egy sticky proxy IP-vel és megmaradó cookie-kkal."""

    def __init__(self, config: dict, supabase: Client):
        self.config = config
        self.supabase = supabase
        self.pw = self.browser = self.context = None
        self.engine = None
        self.sid = None
        self.cookies = []
        self.used = 0

    # ---- Supabase-ben tárolt session
    def load(self):
        try:
            rows = self.supabase.table("browser_sessions").select("*").eq("domain", SESSION_KEY).execute().data
            if rows:
                self.sid = rows[0].get("proxy_session")
                self.cookies = rows[0].get("cookies") or []
                print(f"Mentett Google session betöltve (sid {self.sid}, {len(self.cookies)} cookie)")
        except Exception as e:
            print(f"Figyelem: a mentett session nem töltődött be: {e}")

    async def save(self):
        if not self.context:
            return
        try:
            self.cookies = await self.context.cookies()
            self.supabase.table("browser_sessions").upsert({
                "domain": SESSION_KEY,
                "proxy_session": self.sid,
                "cookies": self.cookies,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            }, on_conflict="domain").execute()
        except Exception as e:
            print(f"Figyelem: a session mentése nem sikerült: {e}")

    # ---- böngésző
    async def start(self):
        headful = bool(os.environ.get("DISPLAY"))
        attempts = []
        try:
            from patchright.async_api import async_playwright as patch_pw
            attempts.append(("patchright+chrome", patch_pw, {"channel": "chrome", "headless": not headful}))
            attempts.append(("patchright+chromium", patch_pw, {"headless": not headful}))
        except ImportError:
            pass
        for name, factory, opts in attempts:
            try:
                self.pw = await factory().start()
                self.browser = await self.pw.chromium.launch(**opts)
                self.engine = name
                print(f"Böngésző: {name}{' (látható, virtuális kijelzőn)' if headful else ' (headless)'}")
                return
            except Exception as e:
                print(f"{name} nem indult: {str(e).splitlines()[0][:150]}")
                if self.pw:
                    await self.pw.stop()
                self.pw = self.browser = None
        raise SystemExit("HIBA: nem indult el a böngésző")

    async def new_identity(self, reason: str = ""):
        """Új holland IP és tiszta cookie-k."""
        if reason:
            print(f"  ↻ új IP ({reason})")
        await self._close_context()
        self.sid = str(random.randint(100000, 999999))
        self.cookies = []
        await self._open_context()

    async def ensure(self):
        if not self.context:
            if not self.sid:
                self.sid = str(random.randint(100000, 999999))
            await self._open_context()

    async def _open_context(self):
        opts = {"locale": "nl-NL", "timezone_id": "Europe/Amsterdam", "no_viewport": True}
        proxy_cfg = self.config.get("proxy", {})
        if proxy_cfg.get("enabled"):
            p = build_proxy({**proxy_cfg, "rotate": False})
            # sticky: {user}-{country}-{sid} = ugyanaz az IP a session alatt
            p["username"] = f"{p['username']}-{self.sid}"
            opts["proxy"] = p
        self.context = await self.browser.new_context(**opts)

        async def route(r):
            if r.request.resource_type in BLOCKED_TYPES:
                await r.abort()
            else:
                await r.continue_()
        await self.context.route("**/*", route)

        base = [{"name": "SOCS", "value": SOCS_COOKIE, "domain": ".google.nl", "path": "/"}]
        try:
            await self.context.add_cookies(self.cookies or base)
        except Exception:
            await self.context.add_cookies(base)
        self.used = 0
        await self._warm_up()

    async def _warm_up(self):
        """Főoldal megnyitása, mint egy valódi látogató."""
        page = await self.context.new_page()
        try:
            await page.goto("https://www.google.nl/?hl=nl", wait_until="domcontentloaded", timeout=45000)
            await handle_consent(page)
            await page.wait_for_timeout(random.randint(1500, 3500))
            await page.mouse.move(random.randint(200, 900), random.randint(150, 500))
            await page.wait_for_timeout(random.randint(600, 1500))
        except Exception as e:
            print(f"  warm-up hiba: {str(e)[:120]}")
        finally:
            await page.close()

    async def _close_context(self):
        if self.context:
            try:
                await self.context.close()
            except Exception:
                pass
        self.context = None

    async def close(self):
        await self._close_context()
        if self.browser:
            await self.browser.close()
        if self.pw:
            await self.pw.stop()


async def handle_consent(page):
    if "consent.google" not in page.url:
        return
    for label in ["Alles afwijzen", "Reject all", "Alles accepteren", "Accept all"]:
        btn = page.get_by_role("button", name=label)
        if await btn.count():
            await btn.first.click()
            await page.wait_for_load_state("domcontentloaded")
            await asyncio.sleep(1)
            return


async def is_captcha(page) -> bool:
    if "/sorry/" in page.url:
        return True
    content = (await page.content()).lower()
    return any(m in content for m in CAPTCHA_MARKERS)


async def dump_debug(page, ean: str, tag: str):
    try:
        DEBUG_DIR.mkdir(exist_ok=True)
        base = DEBUG_DIR / f"{ean}_{tag}"
        base.with_suffix(".html").write_text(await page.content(), encoding="utf-8")
        await page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
    except Exception as e:
        print(f"     debug mentés hiba: {e}")


# ---------------------------------------------------------------- extract

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



async def search_shopping(session: GoogleSession, product: dict, settings: dict) -> dict:
    query = product["ean"]
    result = {
        "code": product["code"], "ean": product["ean"], "name": product["name"],
        "query": query, "status": "error", "error": None, "organic": [], "shopping": [],
    }
    page = await session.context.new_page()
    try:
        url = f"https://www.google.nl/search?q={query}&udm=28&hl=nl&gl=nl"
        await page.goto(url, wait_until="domcontentloaded", timeout=45000)
        await handle_consent(page)
        await page.wait_for_timeout(random.randint(1800, 3500))
        session.used += 1

        if await is_captcha(page):
            result["status"], result["error"] = "captcha", "CAPTCHA"
            return result

        try:
            await page.mouse.wheel(0, random.randint(300, 800))
            await page.wait_for_timeout(random.randint(500, 1200))
        except Exception:
            pass

        raw = await page.evaluate(SHOPPING_JS, settings.get("max_shopping", 12) * 2)
        items, dropped = [], 0
        for r in raw["items"]:
            if not is_relevant(r["title"], product):
                dropped += 1
                continue
            items.append({
                "result_type": "shopping",
                "position": len(items) + 1,
                "title": r["title"][:500],
                "price": clean_price(r["price_text"]),
                "currency": "EUR",
                "merchant": r["merchant"],
                "link": r["link"],
                "snippet": None,
            })
            if len(items) >= settings.get("max_shopping", 12):
                break
        result["shopping"] = items
        result["status"] = "ok" if items else "no_results"
        if dropped:
            result["error"] = f"{dropped} nem releváns találat kiszűrve ({raw['mode']})"
        if settings.get("debug") and (not items or raw["mode"] == "heuristic"):
            await dump_debug(page, query, f"shopping_{raw['mode']}")
    except Exception as e:
        result["error"] = str(e)[:500]
        result["status"] = "error"
    finally:
        await page.close()
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
    per_ip = settings.get("searches_per_ip", 8)
    max_captcha = settings.get("max_captcha_per_run", 4)

    print(f"[{datetime.now().isoformat()}] Google shopping scraper indult")

    products_file = resolve_products_file(settings["products_file"])
    products = load_products(products_file)
    print(f"Összesen {len(products)} termék ({products_file})")

    supabase = get_supabase(config)
    already = get_already_scraped_eans(supabase, hours=skip_hours)
    print(f"Már kész (utolsó {skip_hours} óra): {len(already)} db")

    todo = [p for p in products if p["ean"] not in already]
    random.shuffle(todo)
    todo = todo[:batch_size]
    print(f"Mostani batch: {len(todo)} termék")
    if not todo:
        return

    session = GoogleSession(config, supabase)
    session.load()
    await session.start()
    await session.ensure()

    stats = {"ok": 0, "no_results": 0, "captcha": 0, "error": 0}
    try:
        for i, product in enumerate(todo, 1):
            print(f"\n[{i}/{len(todo)}] {product['name'][:60]} (EAN: {product['ean']})")

            if session.used >= per_ip:
                await session.save()
                await session.new_identity(f"{per_ip} keresés után")

            data = await search_shopping(session, product, settings)
            if data["status"] == "captcha":
                stats["captcha"] += 1
                print(f"  → CAPTCHA ({stats['captcha']}/{max_captcha})")
                if stats["captcha"] >= max_captcha:
                    print("Túl sok CAPTCHA ebben a futásban, leállok.")
                    break
                await session.new_identity("CAPTCHA")
                await asyncio.sleep(random.uniform(20, 40))
                data = await search_shopping(session, product, settings)   # egy újrapróbálás

            stats[data["status"]] = stats.get(data["status"], 0) + 1
            print(f"  → {data['status'].upper()} | shopping: {len(data['shopping'])}"
                  + (f" | {data['error']}" if data.get("error") else ""))
            for it in data["shopping"][:3]:
                print(f"     {it['price']} € | {it['merchant']} | {it['title'][:60]}")

            if data["status"] != "captcha":
                try:
                    save_results(supabase, data)
                except Exception as e:
                    print(f"  → Mentési hiba: {e}")

            if i < len(todo):
                delay = random.uniform(settings.get("min_delay_seconds", 20), settings.get("max_delay_seconds", 45))
                await asyncio.sleep(delay)
    finally:
        await session.save()
        await session.close()

    print(f"\n[{datetime.now().isoformat()}] Batch kész. {stats}")


if __name__ == "__main__":
    asyncio.run(main())
