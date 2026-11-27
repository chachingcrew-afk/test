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
from curl_cffi import requests as curl_requests
from curl_cffi.requests.exceptions import RequestException as CurlRequestException

# curl_cffi sends the same TLS/HTTP2 fingerprint and headers as a real Chrome,
# so bot protection that blocks Python's "requests" library lets it through.
IMPERSONATE = "chrome"
NET_ERRORS = (requests.exceptions.RequestException, CurlRequestException)
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
    # the dashboard often shows the URL with /rest/v1/ at the end; the client adds that itself
    url = re.sub(r"/(rest|auth|storage)/v1/?$", "", url.strip()).rstrip("/")
    # format checks (values themselves are never printed)
    problems = []
    if not re.fullmatch(r"https://[a-z0-9]{20}\.supabase\.co/?", url):
        problems.append("SUPABASE_URL formátuma nem 'https://<projekt-azonosító>.supabase.co' "
                        "(ne a dashboard linkje legyen, és ne legyen a végén /rest/v1)")
    if key.startswith("sb_publishable_"):
        problems.append("SUPABASE_KEY a publishable kulcs, a Secret keys alatti sb_secret_... kulcs kell")
    elif not (key.startswith("sb_secret_") or key.startswith("eyJ")):
        problems.append("SUPABASE_KEY nem sb_secret_... (vagy legacy eyJ...) kezdetű")
    if re.search(r"\s", url + key):
        problems.append("a Supabase secretek egyikében szóköz vagy sortörés van")
    if problems:
        sys.exit("HIBA a Supabase beállításban:\n  - " + "\n  - ".join(problems))
    return create_client(url.rstrip("/"), key)


PROXY_PARTS = {}  # filled by get_proxies, used by the browser fallback


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
    # The dashboard often shows the username already ending in "-rotate" (or "-nl-rotate"):
    # strip those first, then add them back in the right order.
    country = str(proxy_cfg.get("country") or "").strip().lower()
    base = re.sub(r"-rotate$", "", user)
    if country:
        base = re.sub(rf"-{re.escape(country)}$", "", base, flags=re.I)
    user = base
    if country:
        user = f"{user}-{country}"
    if proxy_cfg.get("rotate", True):
        user = f"{user}-rotate"
    masked = (base[:2] + "*" * max(len(base) - 2, 0)) + user[len(base):]
    print(f"Webshare username formátum: {masked}")
    host = proxy_cfg.get("host", "p.webshare.io")
    port = proxy_cfg.get("port", 80)
    PROXY_PARTS.update(host=host, port=port, base=base, country=country, password=password)

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
            "our_price": clean_price(_cell(row, "our_price") or _cell(row, "prijs_(€)")),
            "our_sale_price": clean_price(_cell(row, "our_sale_price") or _cell(row, "actieprijs_(€)")),
        })
    return products


def sync_products(supabase, products):
    """Keep the products table (name, brand, our price) in step with the Excel, for the dashboard."""
    rows = {}
    for p in products:
        if p.get("ean"):
            rows[p["ean"]] = {"ean": p["ean"], "name": p["name"], "brand": p.get("brand"),
                              "our_price": p.get("our_price"),
                              "our_sale_price": p.get("our_sale_price"),
                              "updated_at": datetime.now(timezone.utc).isoformat()}
    if rows:
        supabase.table("products").upsert(list(rows.values()), on_conflict="ean").execute()


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

TRAFFIC = {"bytes": 0, "direct": 0}  # bytes = residential proxy usage, direct = free

# Per shop: how we reach it. "direct" (no proxy, free) -> "proxy" (Dutch residential,
# when the shop blocks direct) -> "browser" (Playwright over the proxy, when even that
# is blocked). Remembered in Supabase competitor_sites.access_mode.
ACCESS = {}
SESSIONS = {}  # "direct" and "proxy" curl_cffi sessions
LEVELS = ["direct", "proxy", "browser"]


def access_mode(domain):
    return ACCESS.get(domain, SESSIONS.get("default_mode", "direct"))


def escalate_domain(domain):
    """Move a shop one level up after it blocked us. Returns the new mode or None."""
    mode = access_mode(domain)
    nxt = None
    for cand in LEVELS[LEVELS.index(mode) + 1:]:
        if cand == "proxy" and "proxy" not in SESSIONS:
            continue
        if cand == "browser" and not BROWSER.enabled:
            continue
        nxt = cand
        break
    if not nxt:
        return None
    ACCESS[domain] = nxt
    label = {"proxy": "holland residential proxyn", "browser": "álcázott böngészővel"}[nxt]
    print(f"  > {domain} blokkol, mostantól {label} próbálom")
    cache = CACHE_REF.get("cache")
    if cache is not None:
        cache.set_access(domain, nxt)
    return nxt


# ================================================================ browser fallback (Playwright)

CHALLENGE_MARKERS = ("cf-chl", "challenge-platform", "just a moment...", "captcha-delivery.com",
                     "_incapsula_resource", "px-captcha", "checking your browser", "ddos-guard")
# third-party scripts that bot checks need; every other third-party request is blocked
CHALLENGE_HOSTS = ("cloudflare", "datadome", "captcha-delivery", "perimeterx", "px-cdn",
                   "px-cloud", "akamai", "incapsula", "imperva", "hcaptcha", "recaptcha", "gstatic")
BLOCKED_TYPES = {"image", "media", "font", "stylesheet", "manifest", "other"}


def looks_like_challenge(text):
    head = (text or "")[:20000].lower()
    return any(m in head for m in CHALLENGE_MARKERS)


def _domain_of(url):
    return norm_domain(urlparse(url).netloc)


class BrowserResponse:
    """Same interface the scraper uses from curl_cffi responses."""

    def __init__(self, status_code, text, url, content_type):
        self.status_code, self.text, self.url = status_code, text, url
        self.content = text.encode("utf-8", "ignore")
        self.headers = {"Content-Type": content_type}

    def json(self):
        t = self.text.strip()
        if t.startswith("<"):  # Chrome shows JSON wrapped in <pre>
            pre = BeautifulSoup(t, "lxml").find("pre")
            t = pre.get_text() if pre else BeautifulSoup(t, "lxml").get_text()
        return json.loads(t)


class BrowserFallback:
    """Real Chrome, started only for shops that block plain requests.

    Stealth: Patchright (a patched Playwright that hides automation traces) driving the
    installed Google Chrome, visible on a virtual display when one exists (xvfb-run in the
    workflow). Falls back to plain Playwright Chromium if Patchright or Chrome is missing.
    One context per shop with a sticky proxy IP; its cookies and IP session are saved in
    Supabase (browser_sessions), so a bot check passed once stays passed next run.
    Images, fonts, CSS and third-party trackers are blocked; bot-check providers are allowed."""

    def __init__(self, enabled=True):
        self.enabled = enabled
        self.pw = self.browser = None
        self.engine = None
        self.contexts = {}
        self.sessions = {}        # domain -> {"sid", "cookies"} loaded from / saved to Supabase
        self.dirty = set()
        self.domains = set()      # domains that need the browser in this run
        self.failed_start = False
        self.supabase = None

    def needs(self, domain):
        return self.enabled and domain in self.domains

    def mark(self, domain, cache=None):
        if self.enabled and domain not in self.domains:
            self.domains.add(domain)
            print(f"  > {domain} blokkol, mostantól böngészővel próbálom")
            if cache is not None:
                cache.mark_browser(domain)

    # ---- saved sessions
    def load_sessions(self, supabase):
        self.supabase = supabase
        if not supabase:
            return
        try:
            rows = supabase.table("browser_sessions").select("*").execute().data
            self.sessions = {r["domain"]: {"sid": r.get("proxy_session"), "cookies": r.get("cookies") or []} for r in rows}
        except Exception as e:
            print(f"FIGYELEM: a mentett böngésző-sütik nem töltődtek be: {e}")

    def save_sessions(self):
        if not self.supabase or not self.dirty:
            return
        rows = [{"domain": d, "proxy_session": self.sessions[d].get("sid"),
                 "cookies": self.sessions[d].get("cookies") or [],
                 "updated_at": datetime.now(timezone.utc).isoformat()} for d in self.dirty if d in self.sessions]
        try:
            self.supabase.table("browser_sessions").upsert(rows, on_conflict="domain").execute()
            print(f"Böngésző-sütik mentve: {', '.join(sorted(self.dirty))}")
        except Exception as e:
            print(f"FIGYELEM: a böngésző-sütik mentése nem sikerült: {e}")

    # ---- browser
    def _start(self):
        if self.browser or self.failed_start:
            return self.browser is not None
        headful = bool(os.environ.get("DISPLAY"))
        attempts = []
        try:
            from patchright.sync_api import sync_playwright as patch_pw
            attempts.append(("patchright+chrome", patch_pw, {"channel": "chrome", "headless": not headful}))
            attempts.append(("patchright+chromium", patch_pw, {"headless": not headful}))
        except ImportError:
            pass
        try:
            from playwright.sync_api import sync_playwright as plain_pw
            attempts.append(("playwright+chromium", plain_pw,
                             {"headless": True, "args": ["--disable-blink-features=AutomationControlled"]}))
        except ImportError:
            pass
        for name, factory, opts in attempts:
            try:
                self.pw = factory().start()
                self.browser = self.pw.chromium.launch(**opts)
                self.engine = name
                print(f"  > böngésző: {name}{' (látható, virtuális kijelzőn)' if headful and 'patchright' in name else ''}")
                return True
            except Exception as e:
                print(f"  > {name} nem indult ({str(e).splitlines()[0][:120]})")
                try:
                    if self.pw:
                        self.pw.stop()
                except Exception:
                    pass
                self.pw = self.browser = None
        self.failed_start = True
        print("FIGYELEM: a böngésző nem indult el, a blokkoló oldalakat kihagyom")
        return False

    def _context(self, domain):
        if domain in self.contexts:
            return self.contexts[domain]
        saved = self.sessions.setdefault(domain, {})
        stealth = self.engine and self.engine.startswith("patchright")
        # Patchright works best with Chrome's own settings: no custom user agent or viewport
        opts = {"locale": "nl-NL", "timezone_id": "Europe/Amsterdam"}
        opts["no_viewport" if stealth else "viewport"] = True if stealth else {"width": 1366, "height": 850}
        if PROXY_PARTS:
            # sticky session: {username}-{country}-{session id} keeps one IP; the same id is
            # reused next run so saved cookies match the IP they were issued to
            sid = saved.get("sid") or str(random.randint(100000, 999999))
            saved["sid"] = sid
            user = "-".join(x for x in (PROXY_PARTS["base"], PROXY_PARTS.get("country"), sid) if x)
            opts["proxy"] = {"server": f"http://{PROXY_PARTS['host']}:{PROXY_PARTS['port']}",
                             "username": user, "password": PROXY_PARTS["password"]}
        ctx = self.browser.new_context(**opts)
        if not stealth:
            ctx.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        if saved.get("cookies"):
            try:
                ctx.add_cookies(saved["cookies"])
            except Exception:
                pass

        def route(r):
            req = r.request
            host = urlparse(req.url).netloc.lower()
            same = _domain_of(req.url) == domain or host.endswith("." + domain)
            if req.resource_type in BLOCKED_TYPES and not any(h in host for h in CHALLENGE_HOSTS):
                return r.abort()
            if not same and not any(h in host for h in CHALLENGE_HOSTS):
                return r.abort()
            return r.continue_()
        ctx.route("**/*", route)

        def count(resp):
            try:
                TRAFFIC["bytes"] += int(resp.headers.get("content-length") or 0)
            except Exception:
                pass
        ctx.on("response", count)
        self.contexts[domain] = ctx
        self._warm_up(ctx, domain)
        return ctx

    def _warm_up(self, ctx, domain):
        """Open the home page once, like a visitor, before asking for anything else."""
        page = ctx.new_page()
        try:
            page.goto(f"https://{domain}/", wait_until="domcontentloaded", timeout=45000)
            self._wait_challenge(page)
            page.wait_for_timeout(random.randint(1500, 3500))
            try:
                page.mouse.move(random.randint(100, 600), random.randint(100, 500))
                page.mouse.wheel(0, random.randint(300, 900))
            except Exception:
                pass
            page.wait_for_timeout(random.randint(800, 1800))
        except Exception:
            pass
        finally:
            page.close()

    def _wait_challenge(self, page, limit=30):
        html = page.content()
        waited = 0
        while looks_like_challenge(html) and waited < limit:
            page.wait_for_timeout(2500)
            waited += 2.5
            try:
                page.wait_for_load_state("domcontentloaded", timeout=5000)
            except Exception:
                pass
            html = page.content()
        return html, waited

    def get(self, url, timeout):
        if not self._start():
            raise FetchError("böngésző nem elérhető", 403)
        domain = _domain_of(url)
        _polite_wait(domain)
        try:
            ctx = self._context(domain)
            page = ctx.new_page()
        except Exception as e:
            raise FetchError(f"böngésző hiba: {type(e).__name__}", 403)
        try:
            try:
                resp = page.goto(url, wait_until="domcontentloaded", timeout=timeout * 1000 + 25000)
            except Exception as e:
                raise FetchError(f"böngésző hiba: {str(e).splitlines()[0][:120]}", 403)
            status = resp.status if resp else 0
            ctype = ""
            try:
                ctype = (resp.headers.get("content-type") or "") if resp else ""
            except Exception:
                pass
            # XML (sitemaps) and JSON: take the raw body, not the page Chrome draws around it
            if status == 200 and ("xml" in ctype or "json" in ctype):
                try:
                    body = resp.body()
                    self._save_cookies(ctx, domain)
                    return BrowserResponse(200, body.decode("utf-8", "ignore"), page.url, ctype)
                except Exception:
                    pass
            html, waited = self._wait_challenge(page)
            if domain in SEARCH_BROWSER_DOMAINS and is_search_url(url):
                # the search results are drawn by JavaScript after the page loads: wait for them
                try:
                    page.wait_for_selector(".product-item-info, a.product-item-link, .product-item", timeout=12000)
                except Exception:
                    pass  # no results shown: keep what the page has (a "no results" page)
                html = page.content()
            if DEBUG["on"]:
                debug_save(url, status, html, f"browser:{self.engine}")
            if looks_like_challenge(html):
                raise FetchError("bot-ellenőrzés a böngészőben sem ment át", 403)
            if waited:
                status = 200  # challenge solved, the final page is the real one
            if status != 200:
                raise FetchError(f"HTTP {status} (böngésző)", status)
            self._save_cookies(ctx, domain)
            return BrowserResponse(200, html, page.url, ctype or "text/html")
        finally:
            page.close()

    def _save_cookies(self, ctx, domain):
        try:
            self.sessions.setdefault(domain, {})["cookies"] = ctx.cookies()
            self.dirty.add(domain)
        except Exception:
            pass

    def close(self):
        self.save_sessions()
        try:
            for c in self.contexts.values():
                c.close()
            if self.browser:
                self.browser.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass


BROWSER = BrowserFallback()
CACHE_REF = {}  # lets fetch() remember browser-only domains in the cache


class FetchError(RuntimeError):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


DEBUG = {"on": os.getenv("DEBUG_CAPTURE", "").strip().lower() in ("1", "true", "yes"), "count": {}}


def debug_save(url, status, body, via):
    """DEBUG_CAPTURE=1: keep what the shop sent (first 400 KB) in Supabase debug_pages."""
    if not DEBUG["on"]:
        return
    domain = _domain_of(url)
    n = DEBUG["count"].get(domain, 0)
    if n >= 30:
        return
    DEBUG["count"][domain] = n + 1
    cache = CACHE_REF.get("cache")
    sb = cache.sb if cache is not None else None
    if not sb:
        return
    try:
        sb.table("debug_pages").insert({"domain": domain, "url": url[:2000], "status": status, "via": via,
                                        "body": (body or "")[:400_000]}).execute()
    except Exception as e:
        print(f"debug mentés hiba: {e}")


LAST_HIT = {}       # domain -> time of the last request
CRAWL_DELAY = {}    # domain -> seconds, from robots.txt


def _polite_wait(domain):
    """Respect the shop's robots.txt Crawl-delay (max 10 s) between our requests."""
    delay = min(CRAWL_DELAY.get(domain, 0), 10)
    if delay and domain in LAST_HIT:
        wait = delay - (time.monotonic() - LAST_HIT[domain])
        if wait > 0:
            time.sleep(wait)
    LAST_HIT[domain] = time.monotonic()


def escalate_ok(flag):
    return flag


def fetch(session, url, timeout, retries, expect_json=False, escalate=True):
    """session is ignored when SESSIONS is set up (main); kept for tests.
    escalate=False: a block (403/429) is reported as an error instead of moving the whole
    shop to the proxy or browser. Used for search pages, which shops often protect alone."""
    domain = _domain_of(url)
    if access_mode(domain) == "browser" or BROWSER.needs(domain):
        return BROWSER.get(url, timeout)
    if domain in SEARCH_BROWSER_DOMAINS and BROWSER.enabled and is_search_url(url):
        return BROWSER.get(url, timeout)
    last_err, last_status = None, None
    for attempt in range(1, retries + 1):
        _polite_wait(domain)
        mode = access_mode(domain)
        if mode == "browser":
            return BROWSER.get(url, timeout)
        sess = SESSIONS.get(mode) or session
        try:
            headers = {"Accept": "application/json"} if expect_json else None
            resp = sess.get(url, timeout=timeout, headers=headers)
            TRAFFIC["bytes" if mode == "proxy" else "direct"] += len(resp.content)
            if DEBUG["on"]:
                debug_save(url, resp.status_code, resp.text, mode)
            blocked = resp.status_code in (403, 429, 503) and (
                resp.status_code != 503 or looks_like_challenge(resp.text))
            if resp.status_code == 200 and not expect_json and looks_like_challenge(resp.text):
                blocked = True
            if blocked and not escalate_ok(escalate):
                raise FetchError(f"HTTP {resp.status_code} (keresés blokkolva)", resp.status_code if resp.status_code != 200 else 403)
            if blocked:
                if SESSIONS and escalate_domain(domain):
                    return fetch(session, url, timeout, retries, expect_json, escalate)
                if not SESSIONS and BROWSER.enabled and not BROWSER.needs(domain):
                    BROWSER.mark(domain, CACHE_REF.get("cache"))
                    return BROWSER.get(url, timeout)
            if resp.status_code == 200:
                return resp
            last_err, last_status = f"HTTP {resp.status_code}", resp.status_code
            if resp.status_code in (404, 410):
                break
        except NET_ERRORS as e:
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


def page_eans(soup):
    """EANs the page itself declares (schema.org gtin fields, itemprop gtin)."""
    out = set()
    for node in _jsonld_products(soup):
        stack = [node]
        while stack:
            n = stack.pop()
            if isinstance(n, list):
                stack.extend(n)
            elif isinstance(n, dict):
                for k, v in n.items():
                    if k.lower() in ("gtin", "gtin8", "gtin12", "gtin13", "gtin14", "ean") and v:
                        out.add(re.sub(r"\D", "", str(v)))
                    elif isinstance(v, (dict, list)):
                        stack.append(v)
    for tag in soup.select('[itemprop^="gtin"]'):
        v = tag.get("content") or tag.get_text(strip=True)
        if v:
            out.add(re.sub(r"\D", "", v))
    return {e for e in out if 8 <= len(e) <= 14}


def ean_on_page(html, ean):
    """EAN shown as part of the product itself. Ignores places where our own
    search term can be echoed back (search box, 'recent searches', tracking
    scripts), because those made a wrong page look like a match."""
    if not ean:
        return False
    # variants: EAN in the page's product JSON, e.g. "ean":"5400858181343" or "barcode": ...
    in_json = set(re.findall(r'"(?:ean|ean13|gtin|gtin8|gtin12|gtin13|gtin14|barcode)"\s*:\s*"?(\d{8,14})', html, re.I))
    if in_json & ean_variants(ean):
        return True
    soup = BeautifulSoup(html, "lxml")
    declared = page_eans(soup)
    if declared:
        return bool(declared & ean_variants(ean))
    for t in soup.find_all(["script", "style", "noscript", "input", "form", "textarea", "select", "header", "nav"]):
        if t.name == "script" and (t.get("type") or "").lower() == "application/ld+json":
            continue
        t.decompose()
    for t in soup.find_all(True):
        if t.attrs is None:
            continue
        marker = " ".join(t.get("class", []) if isinstance(t.get("class"), list) else [str(t.get("class") or "")])
        marker = (marker + " " + str(t.get("id") or "")).lower()
        if any(w in marker for w in ("search", "zoek", "recent", "history", "autocomplete", "suggest")):
            t.decompose()
    text = soup.get_text(" ")
    found = set(re.findall(r"(?<!\d)\d{8,14}(?!\d)", text))
    return bool(found & ean_variants(ean))


def search_queries_for_name(product):
    """Full name first, then a short version (brand + first 2 words).
    Many shop searches need ALL words to match, so a long name like
    'Liewood lunchbox Arthur Peach sea shell mix' finds nothing when the shop
    calls it 'Arthur Lunchbox - Peach / Sea shell'. The match is still checked
    strictly afterwards, so a shorter search query does not lower accuracy."""
    name = " ".join(re.sub(r"\(.*?\)", " ", product["name"]).split())
    brand = (product.get("brand") or "").strip()
    words = [w for w in re.split(r"[\s\-/,]+", name) if w]
    if brand:
        bl = brand.lower().split()
        words = [w for w in words if w.lower() not in bl]
    words = [w for w in words if w.lower().strip(".") not in STOPWORDS]
    short = " ".join(([brand] if brand else []) + words[:2]).strip()
    out = []
    for q in (name, short):
        if q and q.lower() not in [o.lower() for o in out]:
            out.append(q)
    return out


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

# shop-specific search pages, tried first on that shop only
EXTRA_SEARCH = {
    "planethappy.nl": [("html", "https://{d}/zoeken.html?s={q}")],
}

# shops whose search results are drawn by JavaScript: only their search pages go through
# the browser (the product pages keep their normal route)
SEARCH_BROWSER_DOMAINS = {"babylux.nl"}

SKIP_PATH = ("/search", "/zoeken", "/catalogsearch", "/zoek", "/suche", "/recherche", "/cart", "/winkelwagen", "/account", "/login", "/inloggen",
             "/checkout", "/afrekenen", "/contact", "/blog", "/wishlist", "/verlanglijst",
             "/klantenservice", "/service", "/over-ons", "/about", "/pages/", "/policies/")


class SearchUnsupported(Exception):
    pass


def _same_site(url, domain):
    host = urlparse(url).netloc.lower()
    return host == domain or host.endswith("." + domain)


def is_search_url(url):
    """A search or result listing, never a product page."""
    p = urlparse(url)
    path, q = p.path.lower(), p.query.lower()
    if any(s in path for s in ("/search", "/zoeken", "/catalogsearch", "/zoek", "/suche", "/recherche")):
        return True
    return bool(re.search(r"(^|&)(q|s|query|search|keyword|zoekterm|term)=", q))


def candidates_from_html(html, base_url, domain, product, by_ean):
    soup = BeautifulSoup(html, "lxml")
    for t in soup.select("header, nav, footer"):
        t.decompose()
    seen = {}
    for i, a in enumerate(soup.find_all("a", href=True)):
        href = urljoin(base_url, a["href"]).split("#")[0]
        if not _same_site(href, domain) or is_search_url(href):
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


def run_search(session, method, template, domain, query, product, timeout, by_ean, escalate=False):
    """Returns a list of candidate dicts: {url, price?, ean_match?}.
    Raises SearchUnsupported if this search method does not exist on the site."""
    url = template.format(d=domain, q=quote_plus(query))
    try:
        resp = fetch(session, url, timeout, 1, expect_json=method != "html", escalate=escalate)
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
    cands = candidates_from_html(resp.text, resp.url, domain, product, by_ean)
    if not echoes_query(resp.text, query):
        # some result pages don't print the query; for a name search, links that look like
        # our product (name score >= 0.5, already filtered) still show the search worked
        if by_ean or not cands:
            raise SearchUnsupported("search page does not show the query")
    return [{"url": u} for u in cands]


def echoes_query(html, query):
    """A real result page says what was searched in its visible text ("Zoekresultaten voor ...",
    "0 results for ..."). The query only sitting in a canonical link, the search box or a
    tracking script is not enough: a home page that ignores ?s= would pass that test."""
    soup = BeautifulSoup(html, "lxml")
    for t in soup.find_all(["script", "style", "noscript", "head", "input", "textarea", "select", "template"]):
        t.decompose()
    text = " ".join(soup.get_text(" ").split()).lower()
    q = " ".join(str(query).split()).lower()
    return q in text


# ================================================================ verification


def verify_candidate(session, cand, product, competitor, timeout, retries):
    """Returns dict(match_type, score, found, url) or None."""
    v = _verify_candidate(session, cand, product, competitor, timeout, retries)
    if v:
        v["url"] = cand["url"]
    return v


def _verify_candidate(session, cand, product, competitor, timeout, retries):
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
    declared = page_eans(soup)
    if ean and declared and not (declared & ean_variants(ean)):
        return None  # page states another EAN: different product, whatever the name says
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

    def url_owner(self, domain, url):
        """Which product (EAN) is this page already linked to on this domain."""
        u = strip_query(url)
        for (key, d), row in self.urls.items():
            if d == domain and row.get("url") and strip_query(row["url"]) == u:
                return key
        return None

    def get_site(self, domain):
        return self.sites.get(domain)

    def set_site(self, domain, method, template, verified=None):
        old = self.sites.get(domain) or {}
        row = {"domain": domain, "search_method": method, "search_template": template,
               "needs_browser": bool(old.get("needs_browser")),
               "access_mode": old.get("access_mode"),
               "search_verified": bool(old.get("search_verified")) if verified is None else verified,
               "updated_at": datetime.now(timezone.utc).isoformat()}
        self.sites[domain] = row
        self.dirty_sites[domain] = row

    def mark_browser(self, domain):
        old = dict(self.sites.get(domain) or {"domain": domain, "search_method": None, "search_template": None})
        old.update(needs_browser=True, updated_at=datetime.now(timezone.utc).isoformat())
        self.sites[domain] = old
        self.dirty_sites[domain] = old

    def browser_domains(self):
        return {d for d, r in self.sites.items() if r.get("needs_browser")}

    def set_access(self, domain, mode):
        old = dict(self.sites.get(domain) or {"domain": domain, "search_method": None, "search_template": None})
        old.update(access_mode=mode, needs_browser=(mode == "browser"),
                   updated_at=datetime.now(timezone.utc).isoformat())
        self.sites[domain] = old
        self.dirty_sites[domain] = old

    def access_modes(self):
        out = {}
        for d, r in self.sites.items():
            if r.get("access_mode"):
                out[d] = r["access_mode"]
            elif r.get("needs_browser"):
                out[d] = "browser"
        return out

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


FORM_SEARCH = {}    # domain -> discovered search template or None (once per run)


def discover_search_form(session, domain, timeout):
    """Read the search form on the shop's home page: its action URL and the text field name.
    Returns a ("html", template) pair, or None."""
    if domain in FORM_SEARCH:
        return FORM_SEARCH[domain]
    FORM_SEARCH[domain] = None
    try:
        resp = fetch(session, f"https://{domain}/", timeout, 1)
    except Exception:
        return None
    soup = BeautifulSoup(resp.text, "lxml")
    base = resp.url or f"https://{domain}/"
    best = None
    for form in soup.find_all("form"):
        method = (form.get("method") or "get").lower()
        if method != "get":
            continue
        fields = form.find_all("input")
        text = [i for i in fields if (i.get("type") or "text").lower() in ("search", "text") and i.get("name")]
        if not text:
            continue
        marker = " ".join([form.get("action") or "", form.get("id") or "", " ".join(form.get("class") or []),
                           form.get("role") or ""] + [f"{i.get('name')} {i.get('placeholder') or ''}" for i in text]).lower()
        score = sum(w in marker for w in ("search", "zoek", "query", "keyword", "suche", "recherche"))
        if text[0].get("type", "").lower() == "search":
            score += 2
        if best is None or score > best[0]:
            best = (score, form, text[0])
    if not best or best[0] == 0:
        return None
    _, form, field = best
    action = urljoin(base, form.get("action") or base)
    if not _same_site(action, domain):
        return None
    p = urlparse(action)
    hidden = [(i.get("name"), i.get("value") or "") for i in form.find_all("input")
              if (i.get("type") or "").lower() == "hidden" and i.get("name")]
    params = "&".join([f"{quote_plus(n)}={quote_plus(v)}" for n, v in hidden] + [f"{quote_plus(field['name'])}={{q}}"])
    path = p.path or "/"
    host = p.netloc.lower()
    host_tpl = "{d}" if host == domain else host  # keep "www." etc. as the shop uses it
    template = f"https://{host_tpl}{path}?{params}"
    FORM_SEARCH[domain] = ("html", template)
    print(f"  > {domain} kereső az oldal űrlapjából: {template}")
    return FORM_SEARCH[domain]


SITEMAPS = {}       # domain -> list of product-ish URLs (loaded once per run)
PROBED = set()      # domains whose search URLs were all tried in this run


def _locs(xml_text):
    return [u.strip() for u in re.findall(r"<loc>\s*([^<\s]+)\s*</loc>", xml_text)]


SITEMAP_PAID_MB = 50.0  # max sitemap MB per shop over proxy/browser (config: sitemap_max_mb_paid)
SITEMAP_BYTES = {}
SITEMAP_CUT = set()  # domains whose sitemap download hit the MB cap (not saved: incomplete)


def _fetch_text(session, url, timeout):
    domain = _domain_of(url)
    paid = access_mode(domain) != "direct" or BROWSER.needs(domain)
    if paid and SITEMAP_BYTES.get(domain, 0) / 1_000_000 >= SITEMAP_PAID_MB:
        SITEMAP_CUT.add(domain)
        raise FetchError("sitemap forgalmi limit", 413)
    resp = fetch(session, url, timeout, 1)
    body = resp.content
    SITEMAP_BYTES[domain] = SITEMAP_BYTES.get(domain, 0) + len(body)
    if body[:2] == b"\x1f\x8b":  # .xml.gz
        import gzip
        body = gzip.decompress(body)
        return body.decode("utf-8", "ignore")
    return resp.text


SITEMAP_CACHE_DAYS = 7


def _sb():
    cache = CACHE_REF.get("cache")
    return cache.sb if cache is not None else None


def _load_saved_sitemap(domain):
    sb = _sb()
    if not sb:
        return None
    try:
        rows = sb.table("site_sitemaps").select("*").eq("domain", domain).execute().data
    except Exception:
        return None
    if not rows:
        return None
    row = rows[0]
    when = datetime.fromisoformat(str(row["fetched_at"]).replace("Z", "+00:00"))
    if datetime.now(timezone.utc) - when > timedelta(days=SITEMAP_CACHE_DAYS):
        return None
    urls = [u for u in (row.get("urls") or "").split("\n") if u]
    return urls or None


def _save_sitemap(domain, urls):
    sb = _sb()
    if not sb or not urls:
        return
    try:
        sb.table("site_sitemaps").upsert({"domain": domain, "urls": "\n".join(urls), "url_count": len(urls),
                                          "fetched_at": datetime.now(timezone.utc).isoformat()},
                                         on_conflict="domain").execute()
    except Exception as e:
        print(f"FIGYELEM: a(z) {domain} sitemap mentése nem sikerült: {e}")


def sitemap_urls(session, domain, timeout):
    """All page URLs from the shop's sitemaps (product sitemaps first). Cached per run, and in
    Supabase for SITEMAP_CACHE_DAYS when the shop is reached over the paid proxy/browser."""
    if domain in SITEMAPS:
        return SITEMAPS[domain]
    paid = access_mode(domain) != "direct" or BROWSER.needs(domain)
    if paid:
        saved = _load_saved_sitemap(domain)
        if saved:
            SITEMAPS[domain] = saved
            print(f"  > {domain} sitemap a mentésből: {len(saved)} oldal")
            return saved
    roots = []
    try:
        robots = _fetch_text(session, f"https://{domain}/robots.txt", timeout)
        roots += re.findall(r"(?im)^\s*sitemap:\s*(\S+)", robots)
        m = re.search(r"(?is)user-agent:\s*\*.*?crawl-delay:\s*([\d.]+)", robots)
        if m:
            CRAWL_DELAY[domain] = float(m.group(1))
    except Exception:
        pass
    roots += [f"https://{domain}/sitemap.xml", f"https://{domain}/sitemap_index.xml",
              f"https://www.{domain}/sitemap.xml", f"https://{domain}/media/sitemap.xml",
              f"https://{domain}/pub/media/sitemap.xml", f"https://{domain}/media/sitemap/sitemap.xml",
              f"https://{domain}/sitemap/sitemap.xml", f"https://{domain}/sitemaps/sitemap.xml",
              f"https://{domain}/sitemap_nl.xml", f"https://{domain}/media/sitemap_nl.xml"]
    urls, seen_maps = [], set()
    for root in roots:
        if urls:
            break
        try:
            text = _fetch_text(session, root, timeout)
        except Exception:
            continue
        seen_maps.add(root)
        locs = _locs(text)
        if "<sitemapindex" in text:
            children = [u for u in locs if u not in seen_maps]
            # product sitemaps first; skip images, posts, pages, categories
            children.sort(key=lambda u: (0 if "product" in u.lower() else 1))
            children = [u for u in children if not re.search(r"(image|post|page|categor|blog|tag|author)", u, re.I)
                        or "product" in u.lower()][:12]
            for child in children:
                try:
                    urls += _locs(_fetch_text(session, child, timeout))
                except Exception:
                    continue
                if len(urls) > 60000:
                    break
        else:
            urls += locs
    urls = [u for u in urls if _same_site(u, domain) and not u.lower().endswith((".jpg", ".png", ".webp", ".xml"))]
    SITEMAPS[domain] = urls
    if paid and urls and domain not in SITEMAP_CUT:
        _save_sitemap(domain, urls)  # complete download only; a cut-off list would hide products
    if domain in SITEMAP_CUT:
        print(f"  ! {domain} sitemap csak részben jött le ({SITEMAP_PAID_MB:.0f} MB limit), nem mentem el")
    print(f"  > {domain} sitemap: {len(urls)} oldal")
    return urls


def sitemap_candidates(session, product, domain, timeout, limit=5):
    """Best-matching product pages from the sitemap, by the words in their URL."""
    urls = sitemap_urls(session, domain, timeout)
    if not urls:
        return []
    ean = product.get("ean")
    if ean:
        by_ean = [u for u in urls if any(v in u for v in ean_variants(ean) if len(v) >= 12)]
        if by_ean:  # some shops put the EAN in the product URL
            return [{"url": u, "via": "sitemap-ean"} for u in by_ean[:limit]]
    ours = set(tokens(product["name"])) - set(tokens(product.get("brand") or ""))
    scored = []
    for u in urls:
        slug = urlparse(u).path.rsplit("/", 1)[-1] or urlparse(u).path
        slug_tokens = set(tokens(slug.replace("-", " ").replace("_", " ").replace(".html", "")))
        if not slug_tokens or not ours:
            continue
        hit = 0.0
        for t in ours:
            if t in slug_tokens:
                hit += 1
            elif len(t) > 4 and difflib.get_close_matches(t, slug_tokens, n=1, cutoff=0.8):
                hit += 0.8
        score = hit / len(ours)
        if product.get("brand") and set(tokens(product["brand"])) & slug_tokens:
            score += 0.15
        if score >= 0.34:
            scored.append((score, u))
    scored.sort(key=lambda x: -x[0])
    return [{"url": u, "via": "sitemap"} for _, u in scored[:limit]]


def _check_candidates(session, cands, product, competitor, cache, timeout, retries):
    domain = competitor["domain"]
    for cand in cands:
        if not cand.get("url") or is_search_url(cand["url"]):
            continue
        owner = cache.url_owner(domain, cand["url"])
        if owner and owner != (product.get("ean") or product["name"]):
            continue  # this page already belongs to another product
        try:
            v = verify_candidate(session, cand, product, competitor, timeout, retries)
        except FetchError:
            continue
        if v:
            return v
    return None


def find_on_site(session, product, competitor, cache, timeout, retries):
    """Search the competitor site, then its sitemap.
    Returns (kind, url, verification_or_reason), kind: found | not_sold | blocked | no_search"""
    domain = competitor["domain"]
    queries = []
    if product.get("ean"):
        queries.append((product["ean"], True))
    queries.extend((q, False) for q in search_queries_for_name(product))

    site = cache.get_site(domain) or {}
    if competitor.get("search_url"):
        methods = [("html", competitor["search_url"])]
    elif site.get("search_template") and (site.get("search_verified") or domain in PROBED):
        methods = [(site["search_method"], site["search_template"])]
    else:
        # not proven yet: try every known search URL once per run, the cached one first
        methods = list(EXTRA_SEARCH.get(domain, [])) + list(SEARCH_TEMPLATES)
        if site.get("search_template"):
            methods.sort(key=lambda m: 0 if m[1] == site["search_template"] else 1)
        PROBED.add(domain)

    search_ok = False  # some search answered
    search_blocked = False
    blocked_template = None
    for method, template in methods:
        try:
            got_any = False
            for query, by_ean in queries:
                cands = run_search(session, method, template, domain, query, product, timeout, by_ean)
                search_ok = True
                if cands:
                    got_any = True
                    if not site.get("search_verified") or site.get("search_template") != template:
                        cache.set_site(domain, method, template, verified=True)
                        site = cache.get_site(domain)
                v = _check_candidates(session, cands, product, competitor, cache, timeout, retries)
                if v:
                    return "found", _found_url(cands, v), v
            if got_any:
                break  # this search returns results on this shop: trust its "nothing matched"
            if not site.get("search_template"):
                cache.set_site(domain, method, template, verified=False)  # answers, but unproven
                site = cache.get_site(domain)
        except SearchUnsupported:
            continue
        except FetchError as e:
            if e.status in (404, 405, 410):
                continue
            if e.status in (403, 429):
                search_blocked = True
                blocked_template = (method, template)
                break  # the shop protects its search: use the sitemap below
            # network trouble: try again next run, cache nothing
            return "blocked", None, f"keresés hiba: {e}"

    # the shop's own search form, read from its home page
    if not competitor.get("search_url") and not (site.get("search_verified")):
        form = discover_search_form(session, domain, timeout)
        if form and form[1] not in [m[1] for m in methods]:
            method, template = form
            try:
                got_any = False
                for query, by_ean in queries:
                    cands = run_search(session, method, template, domain, query, product, timeout, by_ean, escalate=True)
                    search_ok = True
                    if cands:
                        got_any = True
                        cache.set_site(domain, method, template, verified=True)
                        site = cache.get_site(domain)
                    v = _check_candidates(session, cands, product, competitor, cache, timeout, retries)
                    if v:
                        return "found", _found_url(cands, v), v
                if got_any:
                    return "not_sold", None, "nem árulja (a shop saját keresője alapján)"
            except (SearchUnsupported, FetchError):
                pass

    # sitemap fallback: shops whose search runs in JavaScript, or lists variants under one product.
    # Not over the browser: a sitemap there would cost a lot of residential traffic.
    try:
        # sitemaps can be tens of MB: over the paid proxy/browser only up to SITEMAP_PAID_MB
        cands = sitemap_candidates(session, product, domain, timeout)
    except FetchError as e:
        cands = []
        if e.status in (403, 429):
            return "blocked", None, f"sitemap hiba: {e}"
    if cands:
        v = _check_candidates(session, cands, product, competitor, cache, timeout, retries)
        if v:
            return "found", _found_url(cands, v), v
    if search_ok or cands or SITEMAPS.get(domain):
        return "not_sold", None, "nem árulja (keresés és sitemap alapján sincs egyező termék)"
    if search_blocked and blocked_template:
        # last resort: the search itself through the proxy / browser
        method, template = blocked_template
        try:
            for query, by_ean in queries:
                cands = run_search(session, method, template, domain, query, product, timeout, by_ean, escalate=True)
                v = _check_candidates(session, cands, product, competitor, cache, timeout, retries)
                if v:
                    cache.set_site(domain, method, template, verified=True)
                    return "found", _found_url(cands, v), v
            return "not_sold", None, "nem árulja (böngészős keresés alapján)"
        except (SearchUnsupported, FetchError) as e:
            return "blocked", None, f"a kereső és a sitemap is blokkol ({e})"
    if not competitor.get("search_url") and not site.get("search_verified"):
        cache.set_site(domain, "none", None)
    return "no_search", None, ("nem találtam működő keresőt és sitemapot ezen az oldalon, "
                               "add meg a search_url oszlopot a competitors munkalapon")


def _found_url(cands, v):
    return v.get("url") or cands[0]["url"]


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
        # a targeted run (ONLY_DOMAINS / ONLY_EANS) always rechecks "not sold"
        if not settings.get("recheck") and datetime.now(timezone.utc) - checked < timedelta(days=settings["research_days"]):
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


class Saver:
    """Saves results and cache in batches, so a killed run loses at most one batch."""

    def __init__(self, supabase, cache, every=10):
        self.sb, self.cache, self.every = supabase, cache, every
        self.pending, self.failed, self.saved = [], False, 0

    def add(self, res):
        self.pending.append(res)
        if len(self.pending) >= self.every:
            self.flush()

    def flush(self):
        try:
            self.cache.flush()
        except Exception as e:
            print(f"HIBA a cache mentésnél: {e}")
        if not self.pending:
            return
        rows, self.pending = self.pending, []
        save_csv(rows)
        if self.sb:
            try:
                save_to_supabase(self.sb, rows)
            except Exception as e:
                self.failed = True
                print(f"HIBA a Supabase mentésnél: {e}")
        self.saved += len(rows)


def check_ip(session, expected_country=None):
    try:
        info = session.get("https://ipinfo.io/json", timeout=15).json()
        ip, country = info.get("ip"), str(info.get("country", "")).lower()
        print(f"Kimenő IP: {ip} (ország: {country.upper() or '?'})")
        if expected_country and country and country != expected_country.lower():
            print(f"FIGYELEM: a proxy nem {expected_country.upper()} IP-t adott, hanem {country.upper()}-t.")
    except NET_ERRORS as e:
        if "407" in str(e) or "Proxy Authentication" in str(e):
            sys.exit("HIBA: a Webshare elutasította a belépést (407 Proxy Authentication Required).\n"
                     "  - WEBSHARE_USER / WEBSHARE_PASS a Rotating Residential csomag Proxy List oldaláról kell,\n"
                     "    Rotating Proxy Endpoint + Username/Password módban\n"
                     "  - ha több csomagod van, a residential csomag adatai kellenek (nem a régi ingyenesé)\n"
                     "  - nézd meg, nincs-e a secret elején/végén szóköz")
        print(f"IP ellenőrzés sikertelen: {e}")
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
    global SITEMAP_PAID_MB
    SITEMAP_PAID_MB = float(s.get("sitemap_max_mb_paid", SITEMAP_PAID_MB))
    max_minutes = float(s.get("max_minutes_per_run", 100) or 0)
    started = time.monotonic()

    products_file = Path(s.get("products_file", "products.xlsx"))
    products = load_products(products_file)
    competitors = load_competitors(products_file)
    only_domains = {norm_domain(d) for d in os.getenv("ONLY_DOMAINS", "").replace(";", ",").split(",") if d.strip()}
    if only_domains:
        competitors = [c for c in competitors if c["domain"] in only_domains]
        print(f"Csak ezek a shopok: {', '.join(sorted(only_domains))}")
    only_eans = {e.strip() for e in os.getenv("ONLY_EANS", "").replace(";", ",").split(",") if e.strip()}
    settings["recheck"] = bool(only_domains or only_eans)
    if only_eans:
        products = [p for p in products if p.get("ean") in only_eans]
        print(f"Csak ezek a termékek (EAN): {', '.join(sorted(only_eans))}")
    if DEBUG["on"]:
        print("Diagnosztika BE: a shopok válaszai a Supabase debug_pages táblába kerülnek")
    proxies = get_proxies(config)
    supabase = get_supabase_client(config)
    BROWSER.enabled = bool(s.get("browser_fallback", True))
    if supabase:
        try:
            sync_products(supabase, products)
        except Exception as e:
            print(f"FIGYELEM: a termékek feltöltése nem sikerült: {e}")
    try:
        cache = Cache(supabase)
        CACHE_REF["cache"] = cache
        BROWSER.load_sessions(supabase)
        for d in cache.browser_domains():
            BROWSER.domains.add(d)
        if BROWSER.enabled and BROWSER.domains:
            print(f"Böngészővel nyitott oldalak (korábban blokkoltak): {', '.join(sorted(BROWSER.domains))}")
    except Exception as e:
        sys.exit(f"HIBA: nem sikerült kapcsolódni a Supabase-hez ({type(e).__name__}: {e}).\n"
                 "  Ellenőrizd a SUPABASE_URL és SUPABASE_KEY secretet (price-tracker projekt).")

    # Chrome impersonation sets a matching User-Agent and headers itself;
    # a custom User-Agent would contradict the fingerprint, so it is not set.
    def new_session(px=None):
        sess = curl_requests.Session(impersonate=IMPERSONATE)
        sess.headers.update({"Accept-Language": "nl-NL,nl;q=0.9,en;q=0.8"})
        if px:
            sess.proxies.update(px)
        return sess

    # proxy.mode: "fallback" (default) = direct first, residential only for shops that block;
    # "always" = everything over the residential proxy
    proxy_mode = str(config.get("proxy", {}).get("mode", "fallback")).lower()
    session = new_session()
    SESSIONS["direct"] = session
    if proxies:
        SESSIONS["proxy"] = new_session(proxies)
        check_ip(SESSIONS["proxy"], config.get("proxy", {}).get("country"))
        if proxy_mode == "always":
            SESSIONS["default_mode"] = "proxy"
    print(f"Proxy mód: {'mindig proxy' if SESSIONS.get('default_mode') == 'proxy' else 'közvetlen, proxy csak blokkolásnál'}")
    ACCESS.update(cache.access_modes())
    if ACCESS:
        print("Korábban blokkoló shopok: " + ", ".join(f"{d} ({m})" for d, m in sorted(ACCESS.items()) if m != "direct"))

    jobs = []
    for p in products:
        if p.get("url"):
            jobs.append((p, None))
        else:
            jobs.extend((p, c) for c in competitors)
    print(f"[{datetime.now().isoformat()}] Indul: {len(products)} termék, {len(competitors)} versenytárs, {len(jobs)} feladat")
    if not competitors and not any(p.get("url") for p in products):
        sys.exit("HIBA: nincs 'competitors' munkalap (vagy minden versenytárs ki van kapcsolva).")

    saver = Saver(supabase, cache)
    results, skipped = [], 0
    blocked = {}  # domain -> consecutive block count (403/429)
    for i, (product, comp) in enumerate(jobs, 1):
        if max_mb and TRAFFIC["bytes"] / 1_000_000 >= max_mb:
            print(f"Forgalmi limit elérve ({max_mb} MB), a futás itt megáll. A többi a következő futásban jön.")
            break
        if max_minutes and (time.monotonic() - started) / 60 >= max_minutes:
            print(f"Időlimit elérve ({max_minutes:.0f} perc), a futás itt megáll. A többi a következő futásban jön.")
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
        saver.add(res)
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

    saver.flush()
    save_failed = saver.failed
    BROWSER.close()

    if not results:
        print(f"Nincs új eredmény ({skipped} kihagyva, mert nemrég kerestük és nem árulják).")
        return

    ok = sum(r["status"] == "ok" for r in results)
    nf = sum(r["status"] == "not_found" for r in results)
    print(f"Kész: {ok} ár, {nf} nem árulja, {len(results) - ok - nf} hiba, {skipped} kihagyva. "
          f"Becsült proxy forgalom: {TRAFFIC['bytes'] / 1_000_000:.1f} MB "
          f"(közvetlen, ingyenes: {TRAFFIC['direct'] / 1_000_000:.1f} MB)")
    if (ok == 0 and nf == 0) or save_failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
