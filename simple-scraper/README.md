# Simple Price Scraper (GitHub Actions + Webshare + Supabase)

A repó gyökerében lévő "Track prices" mellett fut, attól független. Workflow: `.github/workflows/scrape.yml` ("Price Scraper (simple)").

## Mit csinál

Nem kell versenytárs URL: a scraper minden terméket megkeres minden versenytárs oldalán.

1. Ha már tudja, hol van a termék az adott shopban (Supabase `competitor_urls`), egyből azt az oldalt nyitja meg.
2. Ha nem, a shop saját keresőjével rákeres az **EAN**-re, ha az nem hoz egyezést, a **névre**.
   Kereső típusok, amiket magától kipróbál: Shopify, WooCommerce, és a gyakori `/search?q=`, `/zoeken?q=`, `?s=`, Magento, Lightspeed keresők.
3. Csak akkor fogadja el a találatot, ha az **EAN rajta van az oldalon**, vagy a **név nagyon pontosan egyezik** (a számoknak, pl. 30 jaar, 280ml, 6st, mindig egyezniük kell). A `match_type` oszlop mutatja, melyik alapján (`ean` / `name`).
4. A megtalált URL-t elmenti, a "nem árulja" eredményt is (14 napig nem keresi újra).

## products.xlsx

- **products** munkalap: `name`, `ean`, `brand` (a `url` oszlop opcionális: ha kitöltöd, nem keres, hanem azt az oldalt használja). A pricetool export (Product, EAN, Merk) is közvetlenül működik.
- **competitors** munkalap: `domain`, `enabled` (yes/no), opcionálisan `search_url` (pl. `https://shop.nl/zoeken?q={q}`) és `price_selector`.

## Proxy és forgalom

Webshare rotating residential, csak holland IP (`config.json` > `proxy.country = "nl"`). A log elején látszik a kimenő IP országa, a végén a becsült forgalom.
- `max_mb_per_run`: ha egy futás eléri, megáll (a többi a következő futásban jön).
- Az első futás drágább (keresés), utána a cache miatt termékenként és shoponként általában 1 kérés.
- Ha egy shop 3x egymás után blokkol (403/429), abban a futásban kihagyja.

## Secrets

`WEBSHARE_USER`, `WEBSHARE_PASS`, `SUPABASE_URL`, `SUPABASE_KEY` (GitHub > Settings > Secrets and variables > Actions). A `config.json`-ban maradjanak a `YOUR_...` placeholderek.

## Supabase táblák (`schema.sql`)

- `product_prices`: minden futás minden eredménye
- `latest_prices` (view): URL-enként a legfrissebb sikeres ár
- `competitor_urls`: hol találta meg a termékeket
- `competitor_sites`: melyik shopnál melyik kereső működik
