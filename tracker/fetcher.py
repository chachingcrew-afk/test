"""HTTP fetching with browser TLS impersonation, block detection and proxy fallback.

proxy_mode per request:
- "auto"   : try direct first, retry through the proxy only if the site blocks us
- "always" : always go through the proxy
- "never"  : never use the proxy
In auto mode a domain that blocked us once uses the proxy for the rest of the run.
"""
from __future__ import annotations

import json
import logging
import random
import threading
import time
from dataclasses import dataclass
from urllib.parse import urlparse

from curl_cffi import requests as cffi_requests

log = logging.getLogger(__name__)

BLOCK_STATUSES = {401, 403, 407, 429, 503}
BLOCK_MARKERS = (
    "cf-chl-",
    "challenge-platform",
    "attention required! | cloudflare",
    "<title>just a moment...</title>",
    "captcha-delivery.com",
    "px-captcha",
    "_incapsula_resource",
    "<title>access denied</title>",
    "perimeterx",
    "are you a robot",
)


class FetchError(Exception):
    def __init__(self, message: str, *, blocked: bool = False, status: int | None = None):
        super().__init__(message)
        self.blocked = blocked
        self.status = status


@dataclass
class FetchResponse:
    url: str
    status: int
    text: str
    used_proxy: bool

    def json(self):
        return json.loads(self.text)


def looks_blocked(status: int, text: str) -> bool:
    if status in BLOCK_STATUSES:
        return True
    # challenge pages are small; a real product page that merely mentions "captcha" is not a block
    head = text[:20000].lower()
    return len(text) < 200_000 and any(marker in head for marker in BLOCK_MARKERS)


class Fetcher:
    def __init__(self, proxy_url: str | None, timeout: int = 30, retries: int = 2,
                 impersonate: str = "chrome"):
        self.proxy_url = proxy_url or None
        self.timeout = timeout
        self.retries = retries
        self.impersonate = impersonate
        self._local = threading.local()   # one curl session per worker thread
        self._lock = threading.Lock()
        self._proxy_domains: set[str] = set()
        self.bytes_via_proxy = 0

    # ------------------------------------------------------------------ public
    def get(self, url: str, *, proxy_mode: str = "auto", accept: str | None = None) -> FetchResponse:
        domain = urlparse(url).netloc
        use_proxy_first = proxy_mode == "always" or (proxy_mode == "auto" and domain in self._proxy_domains)

        if use_proxy_first:
            return self._get_with_retries(url, use_proxy=True, accept=accept)

        try:
            return self._get_with_retries(url, use_proxy=False, accept=accept)
        except FetchError as err:
            if proxy_mode == "auto" and err.blocked and self.proxy_url:
                log.info("blocked on %s (%s), retrying through proxy", domain, err)
                self._proxy_domains.add(domain)
                return self._get_with_retries(url, use_proxy=True, accept=accept)
            raise

    def get_json(self, url: str, *, proxy_mode: str = "auto"):
        resp = self.get(url, proxy_mode=proxy_mode, accept="application/json")
        try:
            return resp.json()
        except ValueError as err:
            raise FetchError(f"not JSON: {url}") from err

    def should_proxy(self, url: str, proxy_mode: str) -> bool:
        """Whether the headless browser should use the proxy for this URL."""
        if not self.proxy_url or proxy_mode == "never":
            return False
        return proxy_mode == "always" or urlparse(url).netloc in self._proxy_domains

    def mark_blocked(self, url: str) -> None:
        self._proxy_domains.add(urlparse(url).netloc)

    @property
    def session(self):
        if not hasattr(self._local, "session"):
            self._local.session = cffi_requests.Session(impersonate=self.impersonate)
        return self._local.session

    # ----------------------------------------------------------------- private
    def _get_with_retries(self, url: str, *, use_proxy: bool, accept: str | None) -> FetchResponse:
        if use_proxy and not self.proxy_url:
            raise FetchError("proxy requested but PROXY_URL is not set", blocked=True)

        last_err: FetchError | None = None
        for attempt in range(self.retries + 1):
            try:
                return self._get_once(url, use_proxy=use_proxy, accept=accept)
            except FetchError as err:
                last_err = err
                # a block will not go away by hammering the same IP; a rotating proxy might help
                if err.blocked and not use_proxy:
                    raise
                if err.status and 400 <= err.status < 500 and not err.blocked:
                    raise
                if attempt < self.retries:
                    time.sleep(2 ** attempt + random.random())
        raise last_err  # type: ignore[misc]

    def _get_once(self, url: str, *, use_proxy: bool, accept: str | None) -> FetchResponse:
        headers = {"Accept-Language": "en-US,en;q=0.9"}
        if accept:
            headers["Accept"] = accept
        proxies = {"http": self.proxy_url, "https": self.proxy_url} if use_proxy else None
        try:
            resp = self.session.get(url, headers=headers, proxies=proxies, timeout=self.timeout,
                                    allow_redirects=True)
        except Exception as err:  # network level errors from libcurl
            raise FetchError(f"network error: {err}") from err

        text = resp.text or ""
        if use_proxy:
            with self._lock:
                self.bytes_via_proxy += len(resp.content or b"")
        if looks_blocked(resp.status_code, text):
            raise FetchError(f"blocked (HTTP {resp.status_code})", blocked=True, status=resp.status_code)
        if resp.status_code >= 400:
            raise FetchError(f"HTTP {resp.status_code}", status=resp.status_code)
        return FetchResponse(url=str(resp.url), status=resp.status_code, text=text, used_proxy=use_proxy)


# ---------------------------------------------------------------------- browser
def _playwright_proxy(proxy_url: str) -> dict:
    parsed = urlparse(proxy_url)
    proxy = {"server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}"}
    if parsed.username:
        proxy["username"] = parsed.username
    if parsed.password:
        proxy["password"] = parsed.password
    return proxy


class BrowserRenderer:
    """Headless Chromium for JavaScript-heavy pages. Started lazily, only if needed."""

    def __init__(self, proxy_url: str | None, timeout_ms: int = 45000):
        self.proxy_url = proxy_url
        self.timeout_ms = timeout_ms
        self._pw = None
        self._browser = None

    def _ensure(self):
        if self._browser is None:
            try:
                from playwright.sync_api import sync_playwright
            except ImportError as err:
                raise FetchError("playwright is not installed (pip install -r requirements-browser.txt)") from err
            self._pw = sync_playwright().start()
            self._browser = self._pw.chromium.launch(headless=True)

    def render(self, url: str, *, use_proxy: bool) -> str:
        self._ensure()
        context_kwargs = {
            "locale": "en-US",
            "viewport": {"width": 1366, "height": 900},
        }
        if use_proxy and self.proxy_url:
            context_kwargs["proxy"] = _playwright_proxy(self.proxy_url)
        context = self._browser.new_context(**context_kwargs)
        try:
            page = context.new_page()
            # images, fonts and media are the bulk of the traffic and useless for a price
            page.route("**/*", lambda route: route.abort()
                       if route.request.resource_type in {"image", "media", "font"} else route.continue_())
            resp = page.goto(url, wait_until="domcontentloaded", timeout=self.timeout_ms)
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass  # some pages never go idle; what we have is usually enough
            html = page.content()
            status = resp.status if resp else 0
            if looks_blocked(status, html):
                raise FetchError(f"blocked in browser (HTTP {status})", blocked=True, status=status)
            return html
        finally:
            context.close()

    def close(self):
        if self._browser:
            self._browser.close()
        if self._pw:
            self._pw.stop()
