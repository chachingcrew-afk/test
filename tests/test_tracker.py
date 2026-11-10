"""Offline tests: no network, a fake fetcher serves canned responses."""
import json
import sys
from decimal import Decimal
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tracker.extractors import Context, detect_platform, extract_price  # noqa: E402
from tracker.fetcher import FetchError, FetchResponse, looks_blocked  # noqa: E402
from tracker.main import active_slots  # noqa: E402
from tracker.price import parse_price  # noqa: E402
from tracker.storage import CsvStorage, PriceRecord, from_csv_row  # noqa: E402


# ------------------------------------------------------------------ helpers
class FakeFetcher:
    """routes: url -> text, or url -> FetchError to simulate a block."""

    def __init__(self, routes):
        self.routes = routes
        self.proxy_url = "http://user:pass@proxy:8000"
        self.calls = []

    def get(self, url, *, proxy_mode="auto", accept=None):
        self.calls.append(url)
        value = self.routes.get(url)
        if isinstance(value, FetchError):
            raise value
        if value is None:
            raise FetchError("HTTP 404", status=404)
        return FetchResponse(url=url, status=200, text=value if isinstance(value, str) else json.dumps(value),
                             used_proxy=False)

    def should_proxy(self, url, proxy_mode):
        return False

    def mark_blocked(self, url):
        pass


class FakeRenderer:
    def __init__(self, html):
        self.html = html
        self.used = False

    def render(self, url, *, use_proxy):
        self.used = True
        return self.html


def run(url, routes, product=None, renderer=None):
    ctx = Context(url=url, product={"id": "x", "url": url, **(product or {})}, fetcher=FakeFetcher(routes))
    return extract_price(ctx, renderer), ctx


# -------------------------------------------------------------- price parse
@pytest.mark.parametrize("text,expected", [
    ("€ 1.299,99", "1299.99"),
    ("$1,299.99", "1299.99"),
    ("12,50 €", "12.50"),
    ("1.299", "1299"),
    ("1,299", "1299"),
    ("19.95", "19.95"),
    ("Price: 1 299,00 Ft", "1299.00"),
    ("CHF 2'450.00", "2450.00"),
    (2999, "2999"),
    ("free", None),
])
def test_parse_price(text, expected):
    assert parse_price(text) == (Decimal(expected) if expected else None)


# ------------------------------------------------------------- extractors
def test_shopify_api_from_url_picks_variant_and_currency():
    url = "https://shop.example/products/thing?variant=2"
    routes = {
        "https://shop.example/products/thing.js": {
            "title": "Thing",
            "variants": [{"id": 1, "price": 1999, "available": True},
                         {"id": 2, "price": 2499, "available": False}],
        },
        "https://shop.example/cart.js": {"currency": "USD"},
    }
    result, ctx = run(url, routes)
    assert result.price == Decimal("24.99") and result.currency == "USD"
    assert result.in_stock is False and result.method == "shopify_api"
    assert "https://shop.example/products/thing" not in ctx.fetcher.calls  # never needed the HTML


def test_woocommerce_store_api():
    url = "https://wp.example/product/blue-mug/"
    routes = {"https://wp.example/wp-json/wc/store/v1/products?slug=blue-mug": [{
        "name": "Blue mug", "is_in_stock": True,
        "prices": {"price": "1250", "currency_code": "EUR", "currency_minor_unit": 2}}]}
    result, _ = run(url, routes)
    assert result.price == Decimal("12.50") and result.currency == "EUR" and result.method == "woocommerce_api"


def test_jsonld_when_api_blocked():
    url = "https://shop.example/products/thing"
    html = """<html><script type="application/ld+json">
      {"@context":"https://schema.org","@graph":[{"@type":"Product","name":"Thing",
       "offers":{"@type":"Offer","price":"49.00","priceCurrency":"EUR",
                 "availability":"https://schema.org/InStock"}}]}</script></html>"""
    routes = {"https://shop.example/products/thing.js": FetchError("blocked (HTTP 403)", blocked=True),
              url: html}
    result, ctx = run(url, routes)
    assert result.price == Decimal("49.00") and result.in_stock is True and result.method == "jsonld"
    assert ctx.blocked  # remembered for the status column


def test_jsonld_aggregate_offer_and_trailing_comma():
    html = """<script type="application/ld+json">{"@type":"Product","offers":
      {"@type":"AggregateOffer","lowPrice":"9,99","priceCurrency":"EUR",},}</script>"""
    result, _ = run("https://x.example/p/1", {"https://x.example/p/1": html})
    assert result.price == Decimal("9.99")


def test_meta_tags():
    html = """<meta property="product:price:amount" content="1299.00">
              <meta property="product:price:currency" content="HUF">"""
    result, _ = run("https://x.example/p", {"https://x.example/p": html})
    assert result.price == Decimal("1299.00") and result.currency == "HUF" and result.method == "meta_tags"


def test_microdata():
    html = """<div itemscope itemtype="https://schema.org/Product">
      <span itemprop="price" content="15.90">15,90 €</span>
      <meta itemprop="priceCurrency" content="EUR">
      <link itemprop="availability" href="https://schema.org/OutOfStock"></div>"""
    result, _ = run("https://x.example/p", {"https://x.example/p": html})
    assert result.price == Decimal("15.90") and result.in_stock is False


def test_custom_selector_wins():
    html = """<meta property="product:price:amount" content="99.00">
              <div class="price"><span class="now">€ 79,95</span></div>"""
    result, _ = run("https://x.example/p", {"https://x.example/p": html}, {"selector": ".price .now"})
    assert result.price == Decimal("79.95") and result.method == "css_selector"


def test_magento_and_bigcommerce():
    magento = '<span data-price-amount="34.5" data-price-type="finalPrice">€34,50</span>'
    r1, _ = run("https://m.example/p", {"https://m.example/p": magento})
    assert r1.price == Decimal("34.5") and r1.method == "magento_markup"
    bc = """<script>var BCData = {"product_attributes":{"instock":true,
      "price":{"without_tax":{"value":20,"currency":"USD"}}}};</script>"""
    r2, _ = run("https://b.example/p", {"https://b.example/p": bc})
    assert r2.price == Decimal("20") and r2.currency == "USD"


def test_platform_detected_from_html_triggers_api():
    url = "https://store.example/shop/thing"   # no /products/ in the path
    html = '<html><link href="https://cdn.shopify.com/s/x.css"></html>'
    routes = {url: html,
              "https://store.example/shop/thing.js": {"variants": [{"id": 1, "price": 500, "available": True}]},
              "https://store.example/cart.js": {"currency": "EUR"}}
    result, _ = run(url, routes)
    assert result.price == Decimal("5") and result.method == "shopify_api"


def test_browser_fallback_for_js_pages():
    url = "https://spa.example/p"
    renderer = FakeRenderer('<meta property="og:price:amount" content="10.00">')
    result, ctx = run(url, {url: "<html><div id='root'></div></html>"}, renderer=renderer)
    assert renderer.used and ctx.rendered and result.method == "meta_tags (rendered)"


def test_everything_fails_gives_none_with_log():
    url = "https://x.example/p"
    result, ctx = run(url, {url: "<html>no price here</html>"})
    assert result is None and any("jsonld" in a for a in ctx.attempts)


def test_missing_product_stops_without_browser():
    url = "https://x.example/products/gone"
    renderer = FakeRenderer("<html></html>")
    result, ctx = run(url, {}, renderer=renderer)
    assert result is None and ctx.not_found and not renderer.used


def test_block_detection():
    assert looks_blocked(403, "")
    assert looks_blocked(200, "<title>Just a moment...</title>")
    assert not looks_blocked(200, "<html>normal product page</html>")


def test_detect_platform():
    assert detect_platform("x /wp-content/plugins/woocommerce/ y") == "woocommerce"
    assert detect_platform("<html></html>") is None


# --------------------------------------------------------------- schedule
def test_active_slots():
    assert active_slots(3, 1) == {0}
    assert active_slots(3, 2) == {0, 2}
    assert active_slots(3, 3) == {0, 1, 2}
    assert active_slots(3, 9) == {0, 1, 2}


# ---------------------------------------------------------------- storage
def test_csv_roundtrip(tmp_path):
    path = tmp_path / "prices.csv"
    rec = PriceRecord(checked_at="2026-11-09T20:00:00+00:00", run_id="r", product_id="p", product_name="P",
                      url="https://x", status="ok", price=Decimal("12.50"), currency="EUR", in_stock=True,
                      method="jsonld")
    CsvStorage(path).write([rec])
    CsvStorage(path).write([rec])
    import csv
    rows = list(csv.DictReader(path.open()))
    assert len(rows) == 2
    back = from_csv_row(rows[0])
    assert back.price == Decimal("12.50") and back.in_stock is True and back.used_proxy is False


# ------------------------------------------------------- shopify fallbacks
def test_shopify_meta_when_js_endpoint_404s():
    url = "https://s.example/en/products/lamp"
    html = ('<link rel="canonical" href="https://s.example/products/lamp-2">'
            '<script>Shopify.currency = {"active":"EUR","rate":"1.0"};'
            'var meta = {"product":{"variants":[{"id":7,"price":3495}]}};</script>'
            '<link href="https://cdn.shopify.com/x.css">')
    # neither .js variant exists, so the analytics meta in the HTML is the source
    result, _ = run(url, {url: html})
    assert result.price == Decimal("34.95") and result.currency == "EUR" and result.method == "shopify_meta"


def test_shopify_api_uses_canonical_url():
    url = "https://s.example/collections/all/products/old-handle"
    html = '<link rel="canonical" href="https://s.example/products/new-handle"><link href="https://cdn.shopify.com/a.css">'
    routes = {url: html,
              "https://s.example/products/new-handle.js": {"variants": [{"id": 1, "price": 1000, "available": True}]},
              "https://s.example/cart.js": {"currency": "EUR"}}
    result, _ = run(url, routes)
    assert result.price == Decimal("10") and result.method == "shopify_api"


# --------------------------------------------------------------- shop rules
def test_shop_rules_apply_to_subdomains_and_product_wins(tmp_path):
    from tracker.main import load_config
    (tmp_path / "products.yaml").write_text(
        "products:\n"
        "  - {id: a, url: 'https://nl.shop.example/p/1'}\n"
        "  - {id: b, url: 'https://www.shop.example/p/2', selector: '.own'}\n"
        "  - {id: c, url: 'https://other.example/p/3'}\n")
    (tmp_path / "shops.yaml").write_text("shops:\n  shop.example: {selector: '.price', currency: SEK}\n")
    _, products = load_config(tmp_path / "products.yaml", tmp_path / "none", tmp_path / "shops.yaml")
    by_id = {p["id"]: p for p in products}
    assert by_id["a"]["selector"] == ".price" and by_id["a"]["currency"] == "SEK"
    assert by_id["b"]["selector"] == ".own" and by_id["b"]["currency"] == "SEK"
    assert "selector" not in by_id["c"]
