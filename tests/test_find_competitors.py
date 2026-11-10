import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from find_competitors import ean_on_page, ean_variants  # noqa: E402


def test_variants_cover_ean13_ean12_and_leading_zero():
    assert ean_variants("840355802573") == {"840355802573", "0840355802573"}
    assert "5707304142522" in ean_variants("5707304142522")
    assert "8713291773349" in ean_variants("08713291773349")


def test_ean_found_in_jsonld_gtin():
    html = '<script type="application/ld+json">{"@type":"Product","gtin13":"8712051163192"}</script>'
    assert ean_on_page(html, "8712051163192")


def test_ean_found_with_leading_zero_on_page():
    assert ean_on_page("<p>EAN 0840355802573</p>", "840355802573")


def test_partial_number_is_not_a_match():
    assert not ean_on_page("<p>1871205116319234</p>", "8712051163192")
    assert not ean_on_page("<p>no code here</p>", "8712051163192")
