"""The fallback chain. Each extractor tries one way to get the price and returns
a PriceResult or None. The chain runs them cheapest first and stops at the first
valid price.

Order:
 1. Platform APIs guessed from the URL (Shopify /products/..., WooCommerce /product/...)
 2. Static HTML fetch, then platform APIs detected from the HTML
 3. HTML extractors: custom CSS selector, JSON-LD, meta tags, microdata,
    Magento / BigCommerce markup, Next.js data
 4. Headless browser render (JavaScript pages), then the HTML extractors again
"""
from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field
from typing import Callable
from urllib.parse import parse_qs, quote, urlparse, urlunparse

from bs4 import BeautifulSoup

from .fetcher import BrowserRenderer, FetchError, Fetcher
from .price import PriceResult, detect_currency, normalize_currency, parse_price

log = logging.getLogger(__name__)


@dataclass
class Context:
    url: str
    product: dict
    fetcher: Fetcher
    html: str | None = None
    soup: BeautifulSoup | None = None
    used_proxy: bool = False
    rendered: bool = False
    blocked: bool = False
    not_found: bool = False
    attempts: list[str] = field(default_factory=list)

    @property
    def proxy_mode(self) -> str:
        return self.product.get("proxy_mode", "auto")

    def set_html(self, html: str) -> None:
        self.html = html
        self.soup = BeautifulSoup(html, "html.parser")

    def get_json(self, url: str):
        resp = self.fetcher.get(url, proxy_mode=self.proxy_mode, accept="application/json")
        self.used_proxy = self.used_proxy or resp.used_proxy
        return resp.json()


def _strip_url(url: str) -> str:
    p = urlparse(url)
    return urlunparse((p.scheme, p.netloc, p.path.rstrip("/"), "", "", ""))


def _origin(url: str) -> str:
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"


def _availability(value) -> bool | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    text = str(value).lower()
    if any(k in text for k in ("outofstock", "out of stock", "soldout", "discontinued")):
        return False
    if any(k in text for k in ("instock", "in stock", "limitedavailability", "preorder", "onlineonly")):
        return True
    return None


# =========================================================================
# Platform APIs
# =========================================================================

def shopify_api(ctx: Context) -> PriceResult | None:
    """Every Shopify store serves /products/<handle>.js (prices in cents)."""
    data = ctx.get_json(_strip_url(ctx.url) + ".js")
    variants = data.get("variants") or []
    if not variants:
        return None
    wanted = parse_qs(urlparse(ctx.url).query).get("variant", [None])[0]
    variant = next((v for v in variants if str(v.get("id")) == str(wanted)), None) or \
        next((v for v in variants if v.get("available")), variants[0])
    price = parse_price(variant.get("price"))
    if price is None:
        return None
    currency = ctx.product.get("currency")
    if not currency:
        try:
            currency = ctx.get_json(_origin(ctx.url) + "/cart.js").get("currency")
        except (FetchError, ValueError):
            currency = None
    return PriceResult(price=price / 100, currency=normalize_currency(currency),
                       in_stock=variant.get("available"), title=data.get("title"))


def woocommerce_api(ctx: Context) -> PriceResult | None:
    """WooCommerce Store API (public, no key): /wp-json/wc/store/v1/products?slug=..."""
    slug = [s for s in urlparse(ctx.url).path.split("/") if s][-1]
    items = ctx.get_json(f"{_origin(ctx.url)}/wp-json/wc/store/v1/products?slug={quote(slug)}")
    if not items:
        return None
    item = items[0]
    prices = item.get("prices") or {}
    raw = parse_price(prices.get("price"))
    if raw is None:
        return None
    minor = int(prices.get("currency_minor_unit", 2))
    return PriceResult(price=raw / (10 ** minor), currency=normalize_currency(prices.get("currency_code")),
                       in_stock=item.get("is_in_stock"), title=item.get("name"))


def squarespace_api(ctx: Context) -> PriceResult | None:
    """Squarespace product pages return their data with ?format=json."""
    data = ctx.get_json(_strip_url(ctx.url) + "?format=json")
    item = data.get("item") or {}
    variants = (item.get("structuredContent") or {}).get("variants") or item.get("variants") or []
    if not variants:
        return None
    v = variants[0]
    money = v.get("salePriceMoney") if v.get("onSale") else v.get("priceMoney")
    if money and money.get("value"):
        price, currency = parse_price(money["value"]), money.get("currency")
    else:
        cents = parse_price(v.get("salePrice") if v.get("onSale") else v.get("price"))
        price, currency = (cents / 100 if cents is not None else None), None
    if price is None:
        return None
    stock = None if v.get("unlimited") else (v.get("qtyInStock", 0) > 0 if "qtyInStock" in v else None)
    return PriceResult(price=price, currency=normalize_currency(currency), in_stock=stock,
                       title=item.get("title"))


@dataclass
class ApiStrategy:
    name: str
    platform: str
    fn: Callable[[Context], PriceResult | None]
    url_match: Callable[[str], bool] = lambda url: False


API_STRATEGIES = [
    ApiStrategy("shopify_api", "shopify", shopify_api, lambda u: "/products/" in urlparse(u).path),
    ApiStrategy("woocommerce_api", "woocommerce", woocommerce_api, lambda u: "/product/" in urlparse(u).path),
    ApiStrategy("squarespace_api", "squarespace", squarespace_api),
]

PLATFORM_SIGNATURES = {
    "shopify": ("cdn.shopify.com", "shopify.theme", "myshopify.com"),
    "woocommerce": ("woocommerce", "wc-block", "/wp-content/plugins/woocommerce"),
    "squarespace": ("static1.squarespace.com", "squarespace-cdn.com"),
    "magento": ("mage/cookies", "magento_", "data-mage-init"),
    "bigcommerce": ("bigcommerce.com", "bcdata"),
    "wix": ("wixstatic.com", "x-wix-"),
    "prestashop": ("prestashop",),
    "shopware": ("shopware",),
}


def detect_platform(html: str) -> str | None:
    head = html[:300000].lower()
    for platform, signs in PLATFORM_SIGNATURES.items():
        if any(s in head for s in signs):
            return platform
    return None


# =========================================================================
# HTML extractors
# =========================================================================

def css_selector(ctx: Context) -> PriceResult | None:
    selector = ctx.product.get("selector")
    if not selector:
        return None
    el = ctx.soup.select_one(selector)
    if el is None:
        return None
    text = el.get("content") or el.get("data-price") or el.get_text(" ", strip=True)
    price = parse_price(text)
    if price is None:
        return None
    return PriceResult(price=price, currency=detect_currency(el.get_text(" ", strip=True)))


def _iter_jsonld(soup: BeautifulSoup):
    for tag in soup.find_all("script", attrs={"type": re.compile("ld\\+json", re.I)}):
        raw = tag.string or tag.get_text() or ""
        try:
            yield json.loads(raw, strict=False)
        except ValueError:
            # common breakage: trailing commas, html comments
            cleaned = re.sub(r",\s*([}\]])", r"\1", raw.replace("<!--", "").replace("-->", ""))
            try:
                yield json.loads(cleaned, strict=False)
            except ValueError:
                continue


def _walk(node):
    if isinstance(node, dict):
        yield node
        for v in node.values():
            yield from _walk(v)
    elif isinstance(node, list):
        for v in node:
            yield from _walk(v)


def _types(node: dict) -> set[str]:
    t = node.get("@type")
    return {str(x).lower() for x in (t if isinstance(t, list) else [t]) if x}


def _offer_price(offer: dict):
    for key in ("price", "lowPrice"):
        if offer.get(key) not in (None, ""):
            return parse_price(offer[key]), offer.get("priceCurrency")
    spec = offer.get("priceSpecification")
    specs = spec if isinstance(spec, list) else [spec] if spec else []
    for s in specs:
        if isinstance(s, dict) and s.get("price") not in (None, ""):
            return parse_price(s["price"]), s.get("priceCurrency")
    return None, None


def jsonld(ctx: Context) -> PriceResult | None:
    url_path = urlparse(ctx.url).path.rstrip("/")
    for doc in _iter_jsonld(ctx.soup):
        for node in _walk(doc):
            if not _types(node) & {"product", "productgroup", "individualproduct"}:
                continue
            offers = node.get("offers")
            if offers is None and node.get("hasVariant"):
                offers = [v.get("offers") for v in node["hasVariant"] if isinstance(v, dict)]
            offer_list = [o for o in (offers if isinstance(offers, list) else [offers]) if isinstance(o, dict)]
            # flatten AggregateOffer.offers
            flat = []
            for o in offer_list:
                flat.append(o)
                inner = o.get("offers")
                if isinstance(inner, list):
                    flat.extend(i for i in inner if isinstance(i, dict))
            # prefer the offer whose url matches the page (variant pages)
            flat.sort(key=lambda o: 0 if str(o.get("url", "")).rstrip("/").endswith(url_path) and url_path else 1)
            for offer in flat:
                price, currency = _offer_price(offer)
                if price is not None:
                    return PriceResult(price=price, currency=normalize_currency(currency),
                                       in_stock=_availability(offer.get("availability")),
                                       title=node.get("name"))
    return None


def meta_tags(ctx: Context) -> PriceResult | None:
    def meta(*names):
        for n in names:
            el = ctx.soup.find("meta", attrs={"property": n}) or ctx.soup.find("meta", attrs={"name": n})
            if el and el.get("content"):
                return el["content"]
        return None

    price = parse_price(meta("product:price:amount", "og:price:amount", "twitter:data1"))
    if price is None:
        return None
    return PriceResult(price=price,
                       currency=normalize_currency(meta("product:price:currency", "og:price:currency")),
                       in_stock=_availability(meta("product:availability", "og:availability")),
                       title=meta("og:title"))


def microdata(ctx: Context) -> PriceResult | None:
    el = ctx.soup.find(attrs={"itemprop": "price"})
    if el is None:
        return None
    price = parse_price(el.get("content") or el.get_text(" ", strip=True))
    if price is None:
        return None
    cur = ctx.soup.find(attrs={"itemprop": "priceCurrency"})
    avail = ctx.soup.find(attrs={"itemprop": "availability"})
    return PriceResult(
        price=price,
        currency=normalize_currency(cur.get("content") or cur.get_text(strip=True)) if cur else
        detect_currency(el.get_text(" ", strip=True)),
        in_stock=_availability((avail.get("href") or avail.get("content")) if avail else None),
    )


def magento_markup(ctx: Context) -> PriceResult | None:
    el = ctx.soup.find(attrs={"data-price-type": "finalPrice", "data-price-amount": True}) or \
        ctx.soup.find(attrs={"data-price-amount": True})
    if el is None:
        return None
    price = parse_price(el["data-price-amount"])
    return PriceResult(price=price, currency=detect_currency(el.get_text(" ", strip=True))) if price else None


def bigcommerce_bcdata(ctx: Context) -> PriceResult | None:
    m = re.search(r"var\s+BCData\s*=\s*(\{.*?\});", ctx.html, re.S)
    if not m:
        return None
    try:
        data = json.loads(m.group(1))
    except ValueError:
        return None
    attrs = data.get("product_attributes") or {}
    price_block = attrs.get("price") or {}
    for key in ("with_tax", "without_tax"):
        p = price_block.get(key)
        if isinstance(p, dict) and p.get("value") is not None:
            return PriceResult(price=parse_price(p["value"]), currency=normalize_currency(p.get("currency")),
                               in_stock=attrs.get("instock"))
    return None


PRICE_KEYS = ("price", "salePrice", "finalPrice", "currentPrice", "amount")
CURRENCY_KEYS = ("currency", "currencyCode", "priceCurrency")


def nextjs_data(ctx: Context) -> PriceResult | None:
    """Last resort on Next.js sites: first object that has both a price and a currency."""
    tag = ctx.soup.find("script", id="__NEXT_DATA__")
    if not tag or not tag.string:
        return None
    try:
        data = json.loads(tag.string)
    except ValueError:
        return None
    for node in _walk(data):
        cur = next((node[k] for k in CURRENCY_KEYS if isinstance(node.get(k), str)), None)
        if not cur:
            continue
        for k in PRICE_KEYS:
            val = node.get(k)
            if isinstance(val, dict):
                val = val.get("value") or val.get("amount")
            price = parse_price(val) if isinstance(val, (int, float, str)) else None
            if price:
                return PriceResult(price=price, currency=normalize_currency(cur))
    return None


HTML_STRATEGIES: list[tuple[str, Callable[[Context], PriceResult | None]]] = [
    ("css_selector", css_selector),
    ("jsonld", jsonld),
    ("meta_tags", meta_tags),
    ("microdata", microdata),
    ("magento_markup", magento_markup),
    ("bigcommerce_bcdata", bigcommerce_bcdata),
    ("nextjs_data", nextjs_data),
]


# =========================================================================
# The chain
# =========================================================================

def _run(ctx: Context, name: str, fn) -> PriceResult | None:
    try:
        result = fn(ctx)
    except FetchError as err:
        ctx.blocked = ctx.blocked or err.blocked
        ctx.attempts.append(f"{name}: {err}")
        return None
    except Exception as err:  # one broken extractor must never kill the chain
        ctx.attempts.append(f"{name}: error {type(err).__name__}: {err}")
        log.debug("extractor %s crashed", name, exc_info=True)
        return None
    if result is None or not result.is_valid():
        ctx.attempts.append(f"{name}: no price")
        return None
    result.method = name + (" (rendered)" if ctx.rendered and name in dict(HTML_STRATEGIES) else "")
    ctx.attempts.append(f"{name}: ok {result.price}")
    return result


def _run_html_strategies(ctx: Context) -> PriceResult | None:
    for name, fn in HTML_STRATEGIES:
        if name == "css_selector" and not ctx.product.get("selector"):
            continue
        result = _run(ctx, name, fn)
        if result:
            return result
    return None


def extract_price(ctx: Context, renderer: BrowserRenderer | None) -> PriceResult | None:
    forced = ctx.product.get("platform")
    tried: set[str] = set()

    # 1) platform APIs we can guess from the URL (or the config says the platform)
    for s in API_STRATEGIES:
        if (forced and s.platform == forced) or (not forced and s.url_match(ctx.url)):
            tried.add(s.name)
            result = _run(ctx, s.name, s.fn)
            if result:
                return result

    # 2) static HTML
    try:
        resp = ctx.fetcher.get(ctx.url, proxy_mode=ctx.proxy_mode, accept="text/html,application/xhtml+xml")
        ctx.used_proxy = ctx.used_proxy or resp.used_proxy
        ctx.set_html(resp.text)
    except FetchError as err:
        ctx.blocked = ctx.blocked or err.blocked
        ctx.attempts.append(f"html_fetch: {err}")
        if err.status in (404, 410):
            # the product page is gone: a browser will not bring it back
            ctx.not_found = True
            return None

    if ctx.html:
        platform = detect_platform(ctx.html)
        if platform:
            ctx.attempts.append(f"platform: {platform}")
        for s in API_STRATEGIES:
            if s.platform == platform and s.name not in tried:
                tried.add(s.name)
                result = _run(ctx, s.name, s.fn)
                if result:
                    return result
        result = _run_html_strategies(ctx)
        if result:
            return result

    # 3) headless browser for JavaScript rendered pages
    if renderer and ctx.product.get("render", True):
        use_proxy = ctx.fetcher.should_proxy(ctx.url, ctx.proxy_mode) or (ctx.blocked and ctx.proxy_mode != "never"
                                                                          and bool(ctx.fetcher.proxy_url))
        try:
            ctx.set_html(renderer.render(ctx.url, use_proxy=use_proxy))
            ctx.rendered = True
            ctx.used_proxy = ctx.used_proxy or use_proxy
        except FetchError as err:
            ctx.blocked = ctx.blocked or err.blocked
            ctx.attempts.append(f"browser: {err}")
            # blocked without proxy: one more go through the proxy
            if err.blocked and not use_proxy and ctx.fetcher.proxy_url and ctx.proxy_mode != "never":
                try:
                    ctx.set_html(renderer.render(ctx.url, use_proxy=True))
                    ctx.rendered = ctx.used_proxy = True
                    ctx.fetcher.mark_blocked(ctx.url)
                except FetchError as err2:
                    ctx.attempts.append(f"browser+proxy: {err2}")
        except Exception as err:
            ctx.attempts.append(f"browser: error {type(err).__name__}: {err}")
        if ctx.rendered:
            return _run_html_strategies(ctx)
    return None
