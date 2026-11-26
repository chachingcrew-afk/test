#!/usr/bin/env python3
"""
Price scraper that FINDS the product on competitor sites by itself.

products.xlsx
  sheet "products":    name, ean, [brand]
                       (the Dutch pricetool export also works: Product, EAN, Merk)
                       optional "url" column: if filled, that exact page is used, no search
  sheet "competitors": domain, [enabled], [search_url], [price_selector]
                       search_url: optional template with {q}, e.g. https://shop.nl/zoeken?q={q}

For every product x competitor:
  1. Known URL in the cache (Supabase competitor_urls)?  -> just read the price.
  2. Otherwise search the competitor's own site search for the EAN, then the name.
     Accept a hit only if the EAN is on the page, or the name matches very closely.
  3. Store the found URL (or "not found") so the next run does not search again.

Proxy: Webshare rotating residential, Dutch IPs only (config.json proxy.country).
Secrets come from env vars (GitHub Secrets); config.json keeps YOUR_... placeholders.
"""
import csv
import difflib
import json
import os
import random
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote_plus, urljoin, urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from supabase import create_client, Client

CONFIG_PATH = Path(os.getenv("CONFIG_FILE", "config.json"))
BACKUP_CSV = Path("data/prices.csv")
LOCAL_CACHE = Path("data/cache.json")

NAME_MATCH_MIN = 0.8        # share of our name's words that must appear on the page
MAX_CANDIDATES = 3          # result links checked per search (bandwidth)


# ================================================================ config


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
        print("FIGYELEM: nincs Supabase URL/KEY, csak CSV mentés és helyi cache lesz.")
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
    country = str(proxy_cfg.get("country") or "").strip().lower()
    if country and f"-{country}" not in user:
        user = f"{user}-{country}"
    if proxy_cfg.get("rotate", True) and not user.endswith("-rotate"):
        user = f"{user}-rotate"
    host = proxy_cfg.get("host", "p.webshare.io")
    port = proxy_cfg.get("port", 80)

    proxy_url = f"http://{user}:{password}@{host}:{port}"
    return {"http": proxy_url, "https": proxy_url}


# ================================================================ input


def _cell(row, col):
    v = row.get(col)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    v = str(v).strip()
    return v or None


def _truthy(v, default=True):
    if v is None:
        return default
    return str(v).strip().lower() not in ("0", "no", "nee", "nem", "false", "n", "off")


def load_products(path: Path) -> list[dict]:
    # dtype=str: otherwise pandas turns EANs into numbers ('.0', lost leading zeros)
    xls = pd.ExcelFile(path, engine="openpyxl")
    sheet = "products" if "products" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet, dtype=str)
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    df = df.rename(columns={"product": "name", "merk": "brand"})  # pricetool export
    if "name" not in df.columns:
        raise ValueError(f"Hiányzik a 'name' oszlop a(z) '{sheet}' munkalapon")

    products = []
    for _, row in df.iterrows():
        name = _cell(row, "name")
        if not name:
            continue
        ean = _cell(row, "ean")
        if ean:
            ean = re.sub(r"\D", "", ean.split(".")[0]) or None
        products.append({
            "name": name,
            "ean": ean,
            "brand": _cell(row, "brand"),
            "url": _cell(row, "url"),
            "price_selector": _cell(row, "price_selector"),
            "currency": _cell(row, "currency") or "EUR",
        })
    return products


def norm_domain(d):
    d = str(d or "").strip().lower()
    d = re.sub(r"^https?://", "", d).split("/")[0]
    return d[4:] if d.startswith("www.") else d


def load_competitors(path: Path) -> list[dict]:
    xls = pd.ExcelFile(path, engine="openpyxl")
    if "competitors" not in xls.sheet_names:
        return []
    df = pd.read_excel(xls, "competitors", dtype=str)
    df.columns = [str(c).strip().lower().replace(" ", "_") for c in df.columns]
    out = []
    for _, row in df.iterrows():
        domain = norm_domain(_cell(row, "domain"))
        if not domain or not _truthy(_cell(row, "enabled")):
            continue
        out.append({
            "domain": domain,
            "search_url": _cell(row, "search_url"),
            "price_selector": _cell(row, "price_selector"),
        })
    return out


# ================================================================ http

TRAFFIC = {"bytes": 0}  # rough proxy usage (response bodies only)


class FetchError(RuntimeError):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


def fetch(session, url, timeout, retries, expect_json=False):
    last_err, last_status = None, None
    for attempt in range(1, retries + 1):
        try:
            headers = {"Accept": "application/json"} if expect_json else None
            resp = session.get(url, timeout=timeout, headers=headers)
            TRAFFIC["bytes"] += len(resp.content)
            if resp.status_code == 200:
                return resp
            last_err, last_status = f"HTTP {resp.status_code}", resp.status_code
            if resp.status_code in (404, 410):
                break
        except requests.exceptions.RequestException as e:
            last_err = f"Hálózati hiba: {type(e).__name__}"
        time.sleep(2 * attempt + random.random())
    raise FetchError(last_err or "ismeretlen hiba", last_status)


# ================================================================ text matching

STOPWORDS = {"de", "het", "een", "en", "met", "voor", "van", "x", "st", "stk", "stuks", "stuk",
             "pcs", "pc", "set", "the", "and", "with", "of", "in", "op"}


def tokens(s):
    s = str(s or "").lower().replace("&", " en ")
    s = re.sub(r"(\d)([a-z])", r"\1 \2", s)
    s = re.sub(r"([a-z])(\d)", r"\1 \2", s)
    return [t for t in re.findall(r"[a-z0-9]+", s) if t not in STOPWORDS]


def name_score(ours, theirs, brand=None):
    """Share of our name's words found in their title. Numbers (sizes, counts,
    ages) must all match, otherwise 0: '30 jaar' must never match '40 jaar'."""
    a = set(tokens(ours))
    b = set(tokens(theirs))
    if brand:
        a -= set(tokens(brand))  # brand is often missing from one of the titles
    if not a or not b:
        return 0.0
    nums_a = {t for t in a if t.isdigit()}
    nums_b = {t for t in b if t.isdigit()}
    if not nums_a <= nums_b:
        return 0.0
    hit = 0.0
    for t in a:
        if t in b:
            hit += 1
        elif len(t) > 4 and difflib.get_close_matches(t, b, n=1, cutoff=0.85):
            hit += 0.8  # plural / spelling variants
    return round(hit / len(a), 3)


def ean_variants(ean):
    e = ean.lstrip("0")
    return {ean, e, e.zfill(13), e.zfill(14)}


def ean_on_page(html, ean):
    if not ean:
        return False
    found = set(re.findall(r"(?<!\d)\d{8,14}(?!\d)", html))
    return bool(found & ean_variants(ean))


def search_query_for_name(product):
    name = re.sub(r"\(.*?\)", " ", product["name"])
    return " ".join(name.split())


# ================================================================ price parsing


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


# ================================================================ price strategies
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


def _jsonld_products(soup):
    for node in _iter_jsonld(soup):
        t = node.get("@type")
        if "Product" in (t if isinstance(t, list) else [t]):
            yield node


def strat_jsonld(soup):
    for node in _jsonld_products(soup):
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


def price_from_soup(soup, selector=None):
    if selector:
        found = strat_selector(soup, selector)
        if found:
            return found
    for fn in (strat_jsonld, strat_meta, strat_woocommerce, strat_generic):
        found = fn(soup)
        if found:
            return found
    return None


def shopify_json(session, url, timeout):
    """Shopify /products/<handle>.js: a few KB, has barcode (EAN) and price."""
    if "/products/" not in url:
        return None
    data = fetch(session, strip_query(url) + ".js", timeout, 1, expect_json=True).json()
    return data if isinstance(data, dict) and data.get("variants") else None


def price_from_shopify(data, ean=None, url=""):
    variants = data.get("variants") or []
    v = None
    if ean:
        v = next((x for x in variants if str(x.get("barcode") or "") in ean_variants(ean)), None)
    m = re.search(r"variant=(\d+)", url)
    if not v and m:
        v = next((x for x in variants if str(x.get("id")) == m.group(1)), None)
    v = v or (variants[0] if variants else None)
    if not v:
        return None
    raw = v.get("price")
    price = raw / 100 if isinstance(raw, int) else clean_price(raw)
    return (price, f"shopify: {raw}") if price is not None else None


def page_title(soup):
    for node in _jsonld_products(soup):
        if node.get("name"):
            return str(node["name"])
    og = soup.find("meta", attrs={"property": "og:title"})
    if og and og.get("content"):
        return og["content"]
    h1 = soup.find("h1")
    if h1:
        return h1.get_text(" ", strip=True)
    return soup.title.get_text(" ", strip=True) if soup.title else ""


# ================================================================ site search

SEARCH_TEMPLATES = [
    ("shopify", "https://{d}/search/suggest.json?q={q}&resources[type]=product&resources[limit]=5"),
    ("woo_api", "https://{d}/wp-json/wc/store/v1/products?search={q}&per_page=5"),
    ("html", "https://{d}/search?q={q}"),
    ("html", "https://{d}/zoeken?q={q}"),
    ("html", "https://{d}/?s={q}&post_type=product"),
    ("html", "https://{d}/catalogsearch/result/?q={q}"),
    ("html", "https://{d}/search/{q}/"),
    ("html", "https://{d}/zoeken/{q}/"),
]

SKIP_PATH = ("/search", "/zoeken", "/cart", "/winkelwagen", "/account", "/login", "/inloggen",
             "/checkout", "/afrekenen", "/contact", "/blog", "/wishlist", "/verlanglijst",
             "/klantenservice", "/service", "/over-ons", "/about", "/pages/", "/policies/")


class SearchUnsupported(Exception):
    pass


def _same_site(url, domain):
    host = urlparse(url).netloc.lower()
    return host == domain or host.endswith("." + domain)


def candidates_from_html(html, base_url, domain, product, by_ean):
    soup = BeautifulSoup(html, "lxml")
    for t in soup.select("header, nav, footer"):
        t.decompose()
    seen = {}
    for i, a in enumerate(soup.find_all("a", href=True)):
        href = urljoin(base_url, a["href"]).split("#")[0]
        if not _same_site(href, domain):
            continue
        path = urlparse(href).path.lower()
        if path in ("", "/") or any(s in path for s in SKIP_PATH):
            continue
        in_box = any(
            "product" in " ".join(p.get("class", []) if hasattr(p, "get") else []).lower()
            for p in list(a.parents)[:5]
        )
        looks_product = any(s in path for s in ("/product", "/products/", "/p/", "/artikel")) or path.endswith(".html")
        if not (in_box or looks_product):
            continue
        img = a.find("img")
        text = a.get_text(" ", strip=True) or a.get("title") or (img.get("alt") if img else "") or ""
        score = name_score(product["name"], text, product.get("brand")) if text else 0
        key = strip_query(href) if "/products/" in href else href
        if key not in seen or score > seen[key][1]:
            seen[key] = (i, score)
    items = list(seen.items())
    if by_ean:
        items.sort(key=lambda kv: kv[1][0])              # EAN search: keep page order
    else:
        items = [kv for kv in items if kv[1][1] >= 0.5]   # name search: only similar titles
        items.sort(key=lambda kv: -kv[1][1])
    return [u for u, _ in items[:MAX_CANDIDATES]]


def run_search(session, method, template, domain, query, product, timeout, by_ean):
    """Returns a list of candidate dicts: {url, price?, ean_match?}.
    Raises SearchUnsupported if this search method does not exist on the site."""
    url = template.format(d=domain, q=quote_plus(query))
    try:
        resp = fetch(session, url, timeout, 1, expect_json=method != "html")
    except FetchError as e:
        if e.status in (404, 405, 410, 400, 401):
            raise SearchUnsupported(str(e))
        raise

    if method == "shopify":
        try:
            prods = resp.json()["resources"]["results"]["products"]
        except Exception:
            raise SearchUnsupported("not shopify")
        return [{"url": urljoin(f"https://{domain}/", p["url"]).split("?")[0], "title": p.get("title", "")}
                for p in prods[:MAX_CANDIDATES]]

    if method == "woo_api":
        try:
            prods = resp.json()
            assert isinstance(prods, list)
        except Exception:
            raise SearchUnsupported("not woocommerce")
        out = []
        for p in prods[:MAX_CANDIDATES]:
            c = {"url": p.get("permalink"), "title": p.get("name", "")}
            pr = p.get("prices") or {}
            if pr.get("price") not in (None, ""):
                c["price"] = int(pr["price"]) / (10 ** int(pr.get("currency_minor_unit", 2)))
            if by_ean and str(p.get("sku") or "") in ean_variants(query):
                c["ean_match"] = True
            out.append(c)
        return [c for c in out if c["url"]]

    # html search page: it only "exists" if it answered with a page that echoes the query
    if "text/html" not in resp.headers.get("Content-Type", "text/html"):
        raise SearchUnsupported("not html")
    if query.lower() not in resp.text.lower() and quote_plus(query).lower() not in resp.url.lower():
        raise SearchUnsupported("search page does not echo query")
    return [{"url": u} for u in candidates_from_html(resp.text, resp.url, domain, product, by_ean)]


# ================================================================ verification


def verify_candidate(session, cand, product, competitor, timeout, retries):
    """Returns dict(match_type, score, price, raw) or None."""
    ean, name, brand = product.get("ean"), product["name"], product.get("brand")
    url = cand["url"]

    # cheap path: Shopify JSON (barcode + title + price, a few KB)
    try:
        data = shopify_json(session, url, timeout)
    except Exception:
        data = None
    if data:
        barcodes = {str(v.get("barcode") or "") for v in data["variants"]}
        found = price_from_shopify(data, ean, url)
        if ean and barcodes & ean_variants(ean):
            return {"match_type": "ean", "score": 1.0, "found": found}
        s = name_score(name, data.get("title", ""), brand)
        if s >= NAME_MATCH_MIN:
            return {"match_type": "name", "score": s, "found": found}
        if any(barcodes - {""}):
            return None  # has barcodes, none is ours: definitely another product

    # WooCommerce Store API: SKU == EAN and price already known
    if cand.get("ean_match") and cand.get("price") is not None:
        return {"match_type": "ean", "score": 1.0, "found": (cand["price"], "woo_api")}

    html = fetch(session, url, timeout, retries).text
    soup = BeautifulSoup(html, "lxml")
    found = price_from_soup(soup, competitor.get("price_selector"))
    if ean_on_page(html, ean):
        return {"match_type": "ean", "score": 1.0, "found": found}
    s = name_score(name, page_title(soup), brand)
    if s >= NAME_MATCH_MIN:
        return {"match_type": "name", "score": s, "found": found}
    return None


# ================================================================ cache (Supabase or local file)


class Cache:
    """competitor_urls: (product_key, domain) -> url / not found
       competitor_sites: domain -> working search method"""

    def __init__(self, supabase):
        self.sb = supabase
        self.urls, self.sites = {}, {}
        self.dirty_urls, self.dirty_sites = {}, {}
        if supabase:
            for row in self._select_all("competitor_urls"):
                self.urls[(row["product_key"], row["domain"])] = row
            for row in self._select_all("competitor_sites"):
                self.sites[row["domain"]] = row
        elif LOCAL_CACHE.exists():
            data = json.loads(LOCAL_CACHE.read_text(encoding="utf-8"))
            self.urls = {tuple(k.split("|", 1)): v for k, v in data.get("urls", {}).items()}
            self.sites = data.get("sites", {})
        print(f"Cache: {len(self.urls)} ismert URL, {len(self.sites)} ismert kereső")

    def _select_all(self, table):
        rows, start = [], 0
        while True:
            batch = self.sb.table(table).select("*").range(start, start + 999).execute().data
            rows.extend(batch)
            if len(batch) < 1000:
                return rows
            start += 1000

    def get_url(self, key, domain):
        return self.urls.get((key, domain))

    def set_url(self, key, domain, name, url, match_type, score):
        now = datetime.now(timezone.utc).isoformat()
        row = {"product_key": key, "domain": domain, "product_name": name, "url": url,
               "match_type": match_type, "match_score": score, "checked_at": now}
        self.urls[(key, domain)] = row
        self.dirty_urls[(key, domain)] = row

    def get_site(self, domain):
        return self.sites.get(domain)

    def set_site(self, domain, method, template):
        row = {"domain": domain, "search_method": method, "search_template": template,
               "updated_at": datetime.now(timezone.utc).isoformat()}
        self.sites[domain] = row
        self.dirty_sites[domain] = row

    def flush(self):
        if self.sb:
            if self.dirty_urls:
                self.sb.table("competitor_urls").upsert(
                    list(self.dirty_urls.values()), on_conflict="product_key,domain").execute()
            if self.dirty_sites:
                self.sb.table("competitor_sites").upsert(
                    list(self.dirty_sites.values()), on_conflict="domain").execute()
        else:
            LOCAL_CACHE.parent.mkdir(parents=True, exist_ok=True)
            LOCAL_CACHE.write_text(json.dumps({
                "urls": {f"{k[0]}|{k[1]}": v for k, v in self.urls.items()},
                "sites": self.sites}, ensure_ascii=False, indent=1), encoding="utf-8")
        print(f"Cache mentve: {len(self.dirty_urls)} URL, {len(self.dirty_sites)} kereső frissítve")
        self.dirty_urls, self.dirty_sites = {}, {}


# ================================================================ find + scrape


def find_on_site(session, product, competitor, cache, timeout, retries):
    """Search the competitor site.
    Returns (kind, url, verification_or_reason), kind: found | not_sold | blocked | no_search"""
    domain = competitor["domain"]
    queries = []
    if product.get("ean"):
        queries.append((product["ean"], True))
    queries.append((search_query_for_name(product), False))

    site = cache.get_site(domain)
    if competitor.get("search_url"):
        methods = [("html", competitor["search_url"])]
    elif site and site.get("search_template"):
        methods = [(site["search_method"], site["search_template"])]
    else:
        methods = SEARCH_TEMPLATES

    last_reason = "nem található"
    for method, template in methods:
        try:
            for query, by_ean in queries:
                cands = run_search(session, method, template, domain, query, product, timeout, by_ean)
                if not (site and site.get("search_template") == template):
                    cache.set_site(domain, method, template)  # this search works on this site
                    site = cache.get_site(domain)
                for cand in cands:
                    try:
                        v = verify_candidate(session, cand, product, competitor, timeout, retries)
                    except FetchError:
                        continue
                    if v:
                        return "found", cand["url"], v
            return "not_sold", None, "nem árulja (a kereső működik, de nincs egyező termék)"
        except SearchUnsupported:
            continue
        except FetchError as e:
            # blocked (403/429) or network trouble: try again next run, cache nothing
            return "blocked", None, f"keresés hiba: {e}"
    if not competitor.get("search_url"):
        cache.set_site(domain, "none", None)  # don't probe all URLs again for a while
    return "no_search", None, ("nem találtam működő keresőt ezen az oldalon, "
                               "add meg a search_url oszlopot a competitors munkalapon")


def price_for_known_url(session, url, product, competitor, timeout, retries):
    try:
        data = shopify_json(session, url, timeout)
        if data:
            found = price_from_shopify(data, product.get("ean"), url)
            if found:
                return found
    except FetchError as e:
        if e.status in (404, 410):
            raise
    except Exception:
        pass
    html = fetch(session, url, timeout, retries).text
    return price_from_soup(BeautifulSoup(html, "lxml"),
                           product.get("price_selector") or competitor.get("price_selector"))


def new_result(product, domain, url=None):
    return {
        "name": product["name"], "ean": product["ean"], "url": url, "domain": domain,
        "match_type": None, "price": None, "currency": product.get("currency", "EUR"),
        "status": "error", "error": None, "raw_text": None,
        "scraped_at": datetime.now(timezone.utc).isoformat(),
    }


def site_has_no_search(cache, competitor, research_days):
    site = cache.get_site(competitor["domain"])
    if competitor.get("search_url") or not site or site.get("search_method") != "none":
        return False
    when = datetime.fromisoformat(str(site["updated_at"]).replace("Z", "+00:00"))
    return datetime.now(timezone.utc) - when < timedelta(days=research_days)


def process(session, product, competitor, cache, settings):
    timeout, retries = settings["timeout"], settings["retries"]
    domain = competitor["domain"]
    key = product.get("ean") or product["name"]
    res = new_result(product, domain)

    cached = cache.get_url(key, domain)
    if cached and not cached.get("url"):
        checked = datetime.fromisoformat(str(cached["checked_at"]).replace("Z", "+00:00"))
        if datetime.now(timezone.utc) - checked < timedelta(days=settings["research_days"]):
            return None  # searched recently, not sold there: skip quietly
        cached = None

    if not cached and site_has_no_search(cache, competitor, settings["research_days"]):
        return None  # no usable site search known; fill search_url in the Excel

    if cached:
        res["url"], res["match_type"] = cached["url"], cached.get("match_type")
        try:
            found = price_for_known_url(session, cached["url"], product, competitor, timeout, retries)
        except FetchError as e:
            if e.status in (404, 410):
                cached = None  # page is gone: search again below
            else:
                res["error"] = str(e)
                return res
        else:
            if found:
                res["price"], res["raw_text"], res["status"] = found[0], str(found[1])[:500], "ok"
            else:
                res["error"] = "Nem található ár az ismert oldalon"
            return res

    kind, url, v = find_on_site(session, product, competitor, cache, timeout, retries)
    if kind == "not_sold":
        res["status"], res["error"] = "not_found", v
        cache.set_url(key, domain, product["name"], None, None, None)
        return res
    if kind in ("blocked", "no_search"):
        res["error"] = v
        return res

    cache.set_url(key, domain, product["name"], url, v["match_type"], v["score"])
    res["url"], res["match_type"] = url, v["match_type"]
    if v["found"]:
        res["price"], res["raw_text"], res["status"] = v["found"][0], str(v["found"][1])[:500], "ok"
    else:
        res["error"] = "Megtaláltam a terméket, de árat nem"
    return res


def process_direct(session, product, settings):
    """Products with a fixed url column: no search."""
    domain = norm_domain(product["url"])
    res = new_result(product, domain, product["url"])
    res["match_type"] = "manual"
    try:
        found = price_for_known_url(session, product["url"], product, {}, settings["timeout"], settings["retries"])
        if found:
            res["price"], res["raw_text"], res["status"] = found[0], str(found[1])[:500], "ok"
        else:
            res["error"] = "Nem található ár"
    except FetchError as e:
        res["error"] = str(e)
    return res


# ================================================================ save


RESULT_COLS = ["name", "ean", "url", "domain", "match_type", "price", "currency",
               "status", "error", "raw_text", "scraped_at"]


def save_csv(results):
    BACKUP_CSV.parent.mkdir(parents=True, exist_ok=True)
    new = not BACKUP_CSV.exists()
    with open(BACKUP_CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=RESULT_COLS)
        if new:
            w.writeheader()
        w.writerows(results)
    print(f"CSV mentve: {BACKUP_CSV} ({len(results)} sor)")


def save_to_supabase(supabase: Client, results: list[dict]):
    rows = []
    for r in results:
        row = {c: r[c] for c in RESULT_COLS}
        row["url"] = row["url"] or f"https://{r['domain']}/"  # url is NOT NULL in the table
        rows.append(row)
    for i in range(0, len(rows), 500):
        supabase.table("product_prices").insert(rows[i:i + 500]).execute()
    print(f"Supabase-be mentve: {len(rows)} rekord")


# ================================================================ main


def check_ip(session, expected_country=None):
    try:
        info = session.get("https://ipinfo.io/json", timeout=15).json()
        ip, country = info.get("ip"), str(info.get("country", "")).lower()
        print(f"Kimenő IP: {ip} (ország: {country.upper() or '?'})")
        if expected_country and country and country != expected_country.lower():
            print(f"FIGYELEM: a proxy nem {expected_country.upper()} IP-t adott, hanem {country.upper()}-t.")
    except Exception as e:
        print(f"IP ellenőrzés sikertelen: {e}")


def main():
    config = load_config()
    s = config.get("settings", {})
    settings = {
        "timeout": int(s.get("timeout", 20)),
        "retries": int(s.get("max_retries", 2)),
        "research_days": int(s.get("research_not_found_after_days", 14)),
    }
    delay_min, delay_max = float(s.get("delay_min", 1.5)), float(s.get("delay_max", 4.0))
    max_mb = float(s.get("max_mb_per_run", 0) or 0)

    products_file = Path(s.get("products_file", "products.xlsx"))
    products = load_products(products_file)
    competitors = load_competitors(products_file)
    proxies = get_proxies(config)
    supabase = get_supabase_client(config)
    cache = Cache(supabase)

    session = requests.Session()
    session.headers.update({
        "User-Agent": s.get("user_agent") or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8",
    })
    if proxies:
        session.proxies.update(proxies)
        check_ip(session, config.get("proxy", {}).get("country"))

    jobs = []
    for p in products:
        if p.get("url"):
            jobs.append((p, None))
        else:
            jobs.extend((p, c) for c in competitors)
    print(f"[{datetime.now().isoformat()}] Indul: {len(products)} termék, {len(competitors)} versenytárs, {len(jobs)} feladat")
    if not competitors and not any(p.get("url") for p in products):
        sys.exit("HIBA: nincs 'competitors' munkalap (vagy minden versenytárs ki van kapcsolva).")

    results, skipped = [], 0
    blocked = {}  # domain -> consecutive block count (403/429)
    for i, (product, comp) in enumerate(jobs, 1):
        if max_mb and TRAFFIC["bytes"] / 1_000_000 >= max_mb:
            print(f"Forgalmi limit elérve ({max_mb} MB), a futás itt megáll. A többi a következő futásban jön.")
            break
        dom = comp["domain"] if comp else norm_domain(product["url"])
        if blocked.get(dom, 0) >= 3:
            skipped += 1
            continue
        label = f"  [{i}/{len(jobs)}] {product['name']} @ {dom}"
        try:
            res = process_direct(session, product, settings) if comp is None else process(session, product, comp, cache, settings)
        except Exception as e:
            res = new_result(product, dom)
            res["error"] = f"Váratlan hiba: {e}"[:300]
        if res is None:
            skipped += 1
            continue
        results.append(res)
        err = str(res.get("error") or "")
        blocked[dom] = blocked.get(dom, 0) + 1 if ("403" in err or "429" in err) else 0
        if blocked[dom] == 3:
            print(f"  ! {dom} 3x egymás után blokkolt, ebben a futásban kihagyom")
        if res["status"] == "ok":
            print(f"{label} OK - {res['price']} {res['currency']} ({res['match_type']}) {res['url']}")
        elif res["status"] == "not_found":
            print(f"{label} nincs - {res['error']}")
        else:
            print(f"{label} HIBA - {res['error']}")
        time.sleep(random.uniform(delay_min, delay_max))

    try:
        cache.flush()
    except Exception as e:
        print(f"HIBA a cache mentésnél: {e}")

    if not results:
        print(f"Nincs új eredmény ({skipped} kihagyva, mert nemrég kerestük és nem árulják).")
        return

    save_csv(results)
    save_failed = False
    if supabase:
        try:
            save_to_supabase(supabase, results)
        except Exception as e:
            save_failed = True
            print(f"HIBA a Supabase mentésnél: {e}")

    ok = sum(r["status"] == "ok" for r in results)
    nf = sum(r["status"] == "not_found" for r in results)
    print(f"Kész: {ok} ár, {nf} nem árulja, {len(results) - ok - nf} hiba, {skipped} kihagyva. "
          f"Becsült proxy forgalom: {TRAFFIC['bytes'] / 1_000_000:.1f} MB")
    if (ok == 0 and nf == 0) or save_failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
