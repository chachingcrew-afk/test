#!/usr/bin/env python3
"""
Price scraper: products.xlsx -> Webshare proxy -> Supabase (product_prices) + CSV backup.

products.xlsx columns: name, ean, url, [price_selector], [currency]
  - price_selector is optional. If set, it is tried first; if empty or it
    finds nothing, automatic strategies run (JSON-LD, meta, WooCommerce,
    Shopify JSON, generic).

Secrets: env vars (GitHub Secrets) always win over config.json. Keep
YOUR_... placeholders in the committed config.json.
"""
import csv
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from supabase import create_client, Client

CONFIG_PATH = Path(os.getenv("CONFIG_FILE", "config.json"))
BACKUP_CSV = Path("data/prices.csv")


# ---------------------------------------------------------------- config


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def _real(v):
    """Empty values and YOUR_... / xxxxxxxx placeholders count as missing."""
    v = str(v or "").strip()
    if not v or v.upper().startswith("YOUR_") or "xxxxxxxx" in v:
        return None
    return v


def get_supabase_client(config) -> Client | None:
    url = _real(os.getenv("SUPABASE_URL")) or _real(config.get("supabase", {}).get("url"))
    key = _real(os.getenv("SUPABASE_KEY")) or _real(config.get("supabase", {}).get("key"))
    if not url or not key:
        print("FIGYELEM: nincs Supabase URL/KEY, csak CSV mentés lesz.")
        return None
    return create_client(url, key)


def get_proxies(config):
    proxy_cfg = config.get("proxy", {})
    if not proxy_cfg.get("enabled", False):
        print("FIGYELEM: proxy kikapcsolva, a runner saját IP-jéről megy a kérés.")
        return None

    user = _real(os.getenv("WEBSHARE_USER")) or _real(proxy_cfg.get("username"))
    password = _real(os.getenv("WEBSHARE_PASS")) or _real(proxy_cfg.get("password"))
    if not user or not password:
        sys.exit("HIBA: a proxy be van kapcsolva, de nincs username/password (WEBSHARE_USER / WEBSHARE_PASS secret).")

    # Webshare username parameters: {username}-{country}-rotate
    #   country: e.g. "nl" -> only Dutch IPs; "rotate" -> new IP per request
    country = str(proxy_cfg.get("country") or "").strip().lower()
    if country and f"-{country}" not in user:
        user = f"{user}-{country}"
    if proxy_cfg.get("rotate", True) and not user.endswith("-rotate"):
        user = f"{user}-rotate"
    host = proxy_cfg.get("host", "p.webshare.io")
    port = proxy_cfg.get("port", 80)

    proxy_url = f"http://{user}:{password}@{host}:{port}"
    return {"http": proxy_url, "https": proxy_url}


# ---------------------------------------------------------------- products


def load_products(path: Path) -> list[dict]:
    # dtype=str: otherwise pandas reads EAN as a number and drops leading zeros / adds '.0'
    df = pd.read_excel(path, engine="openpyxl", dtype=str)
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]

    required = {"name", "url"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Hiányzó oszlop(ok) az Excelben: {', '.join(missing)}")

    df = df.dropna(subset=["name", "url"])

    def cell(row, col):
        v = row.get(col) if col in df.columns else None
        return str(v).strip() if pd.notna(v) and str(v).strip() else None

    products = []
    for _, row in df.iterrows():
        url = cell(row, "url")
        if not url or not url.startswith("http"):
            continue
        products.append({
            "name": cell(row, "name"),
            "ean": cell(row, "ean"),
            "url": url,
            "price_selector": cell(row, "price_selector"),
            "currency": cell(row, "currency") or "EUR",
        })
    return products


# ---------------------------------------------------------------- http

TRAFFIC = {"bytes": 0}  # rough proxy usage (response bodies only)


def fetch(session, url, timeout, retries, expect_json=False):
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            headers = {"Accept": "application/json"} if expect_json else None
            resp = session.get(url, timeout=timeout, headers=headers)
            TRAFFIC["bytes"] += len(resp.content)
            if resp.status_code == 200:
                return resp
            last_err = f"HTTP {resp.status_code}"
            if resp.status_code in (404, 410):
                break
        except requests.exceptions.RequestException as e:
            last_err = f"Hálózati hiba: {type(e).__name__}"
        time.sleep(2 * attempt + random.random())
    raise RuntimeError(last_err or "ismeretlen hiba")


# ---------------------------------------------------------------- price parsing


def clean_price(text) -> float | None:
    """'€ 1.299,95' / '1,299.95' / '24,95' / '1.299' -> float.
    If the text holds several amounts (old + sale price), takes the LAST one."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    amounts = re.findall(r"\d[\d.,]*\d|\d", str(text).replace("\xa0", ""))
    if not amounts:
        return None
    s = amounts[-1]
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        head, _, tail = s.rpartition(",")
        s = f"{head.replace(',', '')}.{tail}" if len(tail) <= 2 else s.replace(",", "")
    elif s.count(".") == 1 and len(s.rpartition(".")[2]) == 3:
        s = s.replace(".", "")  # '1.299' = thousand separator
    elif s.count(".") > 1:
        head, _, tail = s.rpartition(".")
        s = f"{head.replace('.', '')}.{tail}" if len(tail) <= 2 else s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def strip_query(url):
    p = urlparse(url)
    return urlunparse(p._replace(query="", fragment="")).rstrip("/")


# ---------------------------------------------------------------- strategies
# Each returns (price, raw_text) or None.


def strat_selector(soup, selector):
    for sel in [s.strip() for s in selector.split(",") if s.strip()]:
        el = soup.select_one(sel)
        if el:
            raw = el.get("content") or el.get_text(" ", strip=True)
            price = clean_price(raw)
            if price is not None:
                return price, raw
    return None


def _iter_jsonld(soup):
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
        except (json.JSONDecodeError, TypeError):
            continue
        stack = [data]
        while stack:
            node = stack.pop()
            if isinstance(node, list):
                stack.extend(node)
            elif isinstance(node, dict):
                yield node
                stack.extend(v for v in node.values() if isinstance(v, (dict, list)))


def strat_jsonld(soup):
    for node in _iter_jsonld(soup):
        t = node.get("@type")
        if "Product" not in (t if isinstance(t, list) else [t]):
            continue
        offers = node.get("offers")
        if isinstance(offers, list):
            offers = offers[0] if offers else None
        if not isinstance(offers, dict):
            continue
        raw = offers.get("price") or offers.get("lowPrice")
        if raw is None and isinstance(offers.get("priceSpecification"), dict):
            raw = offers["priceSpecification"].get("price")
        price = clean_price(raw)
        if price is not None:
            return price, f"jsonld: {raw}"
    return None


def strat_meta(soup):
    for attr, key in [("property", "product:price:amount"), ("property", "og:price:amount"), ("itemprop", "price")]:
        tag = soup.find(attrs={attr: key})
        if tag:
            raw = tag.get("content") or tag.get_text(strip=True)
            price = clean_price(raw)
            if price is not None:
                return price, f"meta: {raw}"
    return None


def strat_woocommerce(soup):
    box = soup.select_one(".summary .price, .product .price, p.price")
    if not box:
        return None
    amount = box.select_one("ins .woocommerce-Price-amount") or box.select_one(".woocommerce-Price-amount")
    if not amount:
        return None
    price = clean_price(amount.get_text())
    return (price, f"woo: {box.get_text(' ', strip=True)}") if price is not None else None


def strat_generic(soup):
    for el in soup.select('[class*="price" i], [id*="price" i], [data-price]'):
        raw = el.get("data-price") or el.get_text(" ", strip=True)
        if raw and re.search(r"[€$£]|\bEUR\b|\d+[.,]\d{2}", raw):
            price = clean_price(raw)
            if price:
                return price, f"generic: {raw[:200]}"
    return None


def strat_shopify(session, url, timeout, retries):
    if "/products/" not in url:
        return None
    data = fetch(session, strip_query(url) + ".js", timeout, retries, expect_json=True).json()
    variants = data.get("variants") or []
    m = re.search(r"variant=(\d+)", url)
    v = next((x for x in variants if m and str(x.get("id")) == m.group(1)), None) or (variants[0] if variants else None)
    if not v:
        return None
    raw = v.get("price")
    price = raw / 100 if isinstance(raw, int) else clean_price(raw)
    return (price, f"shopify: {raw}") if price is not None else None


# ---------------------------------------------------------------- scrape


def scrape_product(session, product, timeout, retries) -> dict:
    result = {
        "name": product["name"],
        "ean": product["ean"],
        "url": product["url"],
        "price": None,
        "currency": product.get("currency", "EUR"),
        "status": "error",
        "error": None,
        "raw_text": None,
        "scraped_at": datetime.now(timezone.utc).isoformat(),
    }
    url = product["url"]
    try:
        found = None
        html_err = None
        soup = None
        # 1) Shopify JSON first: a few KB instead of a 0.5-1.5 MB HTML page
        #    (residential proxy is billed per GB). Skipped if you set a selector.
        if not product.get("price_selector"):
            try:
                found = strat_shopify(session, url, timeout, retries=1)
            except Exception:
                pass
        # 2) HTML page only when needed
        if not found:
            try:
                soup = BeautifulSoup(fetch(session, url, timeout, retries).text, "lxml")
            except RuntimeError as e:
                html_err = str(e)
        if not found and soup is not None:
            if product.get("price_selector"):
                found = strat_selector(soup, product["price_selector"])
            if not found:
                for fn in (strat_jsonld, strat_meta, strat_woocommerce, strat_generic):
                    found = fn(soup)
                    if found:
                        break

        if not found:
            result["error"] = html_err or (
                f"Nem található ár (selector: {product['price_selector']})"
                if product.get("price_selector") else "Nem található ár"
            )
            return result

        result["price"], raw = found
        result["raw_text"] = str(raw)[:500]
        result["status"] = "ok"
    except Exception as e:
        result["error"] = f"Váratlan hiba: {e}"[:300]
    return result


# ---------------------------------------------------------------- save


def save_csv(results):
    BACKUP_CSV.parent.mkdir(parents=True, exist_ok=True)
    new = not BACKUP_CSV.exists()
    with open(BACKUP_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        if new:
            w.writeheader()
        w.writerows(results)
    print(f"CSV mentve: {BACKUP_CSV} ({len(results)} sor)")


def save_to_supabase(supabase: Client, results: list[dict]):
    """Mentés a product_prices táblába, 500-as kötegekben."""
    cols = ["name", "ean", "url", "price", "currency", "status", "error", "raw_text", "scraped_at"]
    rows = [{c: r[c] for c in cols} for r in results]
    for i in range(0, len(rows), 500):
        supabase.table("product_prices").insert(rows[i:i + 500]).execute()
    print(f"Supabase-be mentve: {len(rows)} rekord")


# ---------------------------------------------------------------- main


def check_ip(session, expected_country=None):
    """Logs the outgoing proxy IP and its country."""
    try:
        info = session.get("https://ipinfo.io/json", timeout=15).json()
        ip, country = info.get("ip"), str(info.get("country", "")).lower()
        print(f"Kimenő IP: {ip} (ország: {country.upper() or '?'})")
        if expected_country and country and country != expected_country.lower():
            print(f"FIGYELEM: a proxy nem {expected_country.upper()} IP-t adott, hanem {country.upper()}-t. "
                  "Nézd meg a Webshare proxy listádban, van-e benne holland proxy.")
    except Exception:
        try:
            print(f"Kimenő IP: {session.get('https://ipv4.webshare.io/', timeout=15).text.strip()}")
        except Exception as e:
            print(f"IP ellenőrzés sikertelen: {e}")


def main():
    config = load_config()
    settings = config.get("settings", {})
    timeout = int(settings.get("timeout", 20))
    retries = int(settings.get("max_retries", 3))
    delay_min = float(settings.get("delay_min", 1.5))
    delay_max = float(settings.get("delay_max", 4.0))
    user_agent = settings.get("user_agent")

    products = load_products(Path(settings.get("products_file", "products.xlsx")))
    proxies = get_proxies(config)
    supabase = get_supabase_client(config)

    session = requests.Session()
    session.headers.update({
        "User-Agent": user_agent or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    })
    if proxies:
        session.proxies.update(proxies)
        check_ip(session, config.get("proxy", {}).get("country"))

    results = []
    print(f"[{datetime.now().isoformat()}] Scraping indult - {len(products)} termék")

    for i, product in enumerate(products, 1):
        print(f"  [{i}/{len(products)}] {product['name']} (EAN: {product['ean']}) ...", end=" ", flush=True)
        res = scrape_product(session, product, timeout, retries)
        results.append(res)
        if res["status"] == "ok":
            print(f"OK - {res['price']} {res['currency']}")
        else:
            print(f"HIBA - {res['error']}")
        time.sleep(random.uniform(delay_min, delay_max))

    if not results:
        print("Nincs termék az Excelben.")
        return

    # CSV first, so a Supabase error never loses the data
    save_csv(results)
    save_failed = False
    if supabase:
        try:
            save_to_supabase(supabase, results)
        except Exception as e:
            save_failed = True
            print(f"HIBA a Supabase mentésnél: {e}")

    ok = sum(r["status"] == "ok" for r in results)
    print(f"Kész: {ok}/{len(results)} sikeres. Becsült proxy forgalom: {TRAFFIC['bytes'] / 1_000_000:.1f} MB")
    # red workflow if everything failed (likely proxy/block) or the DB save failed
    if ok == 0 or save_failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
