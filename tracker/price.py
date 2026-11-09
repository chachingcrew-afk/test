"""Price parsing helpers and the result type every extractor returns."""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

CURRENCY_SYMBOLS = {
    "€": "EUR",
    "$": "USD",
    "£": "GBP",
    "¥": "JPY",
    "zł": "PLN",
    "kr": "SEK",
    "Ft": "HUF",
    "CHF": "CHF",
}
ISO_CODE_RE = re.compile(r"\b(USD|EUR|GBP|HUF|PLN|CZK|SEK|NOK|DKK|CHF|CAD|AUD|JPY|RON)\b", re.I)
NUMBER_RE = re.compile(r"\d[\d.,'\s  ]*")

MAX_SANE_PRICE = Decimal("10000000")


@dataclass
class PriceResult:
    price: Decimal
    currency: str | None = None
    in_stock: bool | None = None
    title: str | None = None
    method: str = ""

    def is_valid(self) -> bool:
        return self.price is not None and Decimal("0") < self.price < MAX_SANE_PRICE


def parse_price(value) -> Decimal | None:
    """Turn '1.299,99 €', '$1,299.99', '12,50', 1299 or '1299.0' into a Decimal.

    Rules for separators:
    - both ',' and '.' present: the one that comes last is the decimal separator
    - only one kind present: if it is followed by exactly 3 digits (and appears
      once or repeatedly in thousand groups) it is a thousands separator,
      otherwise it is the decimal separator
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        try:
            return Decimal(str(value))
        except InvalidOperation:
            return None

    text = str(value)
    match = NUMBER_RE.search(text)
    if not match:
        return None
    raw = re.sub(r"[\s  ']", "", match.group(0)).rstrip(".,")
    if not raw:
        return None

    has_comma, has_dot = "," in raw, "." in raw
    if has_comma and has_dot:
        decimal_sep = "," if raw.rfind(",") > raw.rfind(".") else "."
        thousands_sep = "." if decimal_sep == "," else ","
        raw = raw.replace(thousands_sep, "").replace(decimal_sep, ".")
    elif has_comma or has_dot:
        sep = "," if has_comma else "."
        parts = raw.split(sep)
        is_thousands = len(parts) > 1 and all(len(p) == 3 for p in parts[1:]) and 1 <= len(parts[0]) <= 3
        if is_thousands:
            raw = raw.replace(sep, "")
        elif len(parts) == 2:
            raw = raw.replace(sep, ".")
        else:
            # e.g. "1.2.3": not a price we understand
            return None

    try:
        return Decimal(raw)
    except InvalidOperation:
        return None


def detect_currency(text: str | None) -> str | None:
    if not text:
        return None
    code = ISO_CODE_RE.search(text)
    if code:
        return code.group(1).upper()
    for symbol, iso in CURRENCY_SYMBOLS.items():
        if symbol in text:
            return iso
    return None


def normalize_currency(value) -> str | None:
    if not value:
        return None
    value = str(value).strip()
    if len(value) == 3 and value.isalpha():
        return value.upper()
    return detect_currency(value)
