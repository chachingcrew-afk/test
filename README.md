# Price Tracker

Termékárakat ment le megadott weboldalakról, GitHub Actions-ön futva, naponta 1-3 alkalommal. Ha egy oldal blokkol, automatikusan proxyn keresztül próbálja újra.

## Hogyan találja meg az árat

Minden terméknél sorban próbálkozik, és az első sikeres módszernél megáll:

1. **Platform API az URL alapján**: Shopify (`/products/...` → `.js` végpont), WooCommerce (`/product/...` → Store API). Ezek a legolcsóbbak és legpontosabbak, sokszor akkor is működnek, ha a HTML oldalt blokkolják.
2. **HTML letöltése**, majd platform felismerés a HTML-ből (Shopify, WooCommerce, Squarespace, Magento, BigCommerce, Wix, PrestaShop, Shopware), és ha van hozzá API, azzal próbálkozik.
3. **HTML kinyerők**: saját CSS selector (ha megadtad), JSON-LD (schema.org Product), meta tagek (`product:price:amount`), microdata (`itemprop="price"`), Magento és BigCommerce jelölések, Next.js adatok.
4. **Headless böngésző** (Playwright) JavaScriptes oldalakhoz, utána újra a HTML kinyerők.

Az eredmény CSV-ben a `method` oszlop mutatja, melyik módszer talált árat. Sikertelen próbálkozásnál az `error` oszlopban ott van az összes kísérlet naplója, így látod, hol akadt el.

## Proxy logika

- A lekérések Chrome böngészőnek álcázott TLS ujjlenyomattal mennek (`curl_cffi`), ez önmagában sok blokkot megkerül.
- `proxy_mode: auto` (alapértelmezett): először közvetlenül próbálja, és csak akkor vált proxyra, ha az oldal blokkol (403, 429, Cloudflare/captcha oldal). Egy domain, ami egyszer blokkolt, a futás többi részében már egyből proxyt kap. Így csak arra fizetsz, amire tényleg kell.
- `always` / `never` termékenként is beállítható.
- A böngészős renderelés letiltja a képeket, fontokat és videókat, hogy minimális legyen a proxy forgalom.

## Beállítás (egyszer)

1. Hozz létre egy **privát** GitHub repót, és töltsd fel ezt a mappát.
2. **Settings → Secrets and variables → Actions → Secrets**: add hozzá a `PROXY_URL`-t (`http://USERNAME:PASSWORD@HOST:PORT`, a proxy szolgáltató dashboardjáról).
3. Írd be a termékeidet a `config/products.yaml`-ba (a példák `enabled: false`-szal vannak kikapcsolva).
4. **Actions** fül → *Track prices* → **Run workflow**: első kézi futtatás, minden terméket ellenőriz.
5. Az eredmény a `data/prices.csv`-ben, az adott futás összefoglalója pedig az Actions futás oldalán táblázatban.

## Futási gyakoriság

A workflow-ban 3 fix időpont van (UTC 05:41, 13:41, 21:41). Hogy ebből mennyi dolgozik, azt a configban állítod:

```yaml
settings:
  runs_per_day: 1   # 1 = csak reggel, 2 = reggel + este, 3 = mindhárom
```

Termékenként felülírható (`runs_per_day: 3` egy fontos terméknél). Az időpontokat a `.github/workflows/track-prices.yml`-ben módosíthatod, a kód onnan olvassa ki őket. A GitHub ütemezett futásai csúcsidőben 10-30 percet késhetnek, ez normális.

## Átállás Supabase-re (amikor készen állsz)

A kód már fel van készítve, csak be kell kötni:

1. Supabase SQL editorban futtasd le a `supabase/schema.sql`-t (tábla, index, `latest_prices` nézet árváltozás százalékkal).
2. GitHub Secrets: `SUPABASE_URL`, `SUPABASE_SERVICE_KEY` (service role key).
3. GitHub **Variables**: `STORAGE_BACKENDS` = `csv,supabase` (átmenetileg mindkettőbe ír), később `supabase`.
4. A meglévő CSV előzmények áttöltése (többször is futtatható, duplikátumot nem csinál):
   ```
   python scripts/import_csv_to_supabase.py
   ```

## Helyi futtatás

```bash
pip install -r requirements.txt
pip install -r requirements-browser.txt && python -m playwright install chromium   # csak ha kell böngésző
cp .env.example .env    # írd be a PROXY_URL-t

python -m tracker.main run --only my-product-id --dry-run   # egy termék, mentés nélkül
python -m tracker.main run                                  # minden termék
python -m pytest -q                                         # tesztek
```

## Ha egy oldalon nem találja az árat

1. Nézd meg a CSV `error` oszlopát: `blocked` státusz esetén proxy kérdés, `failed` esetén a kinyerés nem sikerült.
2. Adj meg egy `selector`-t a termékhez (böngészőben jobb klikk az áron → Inspect → egyedi class).
3. Ha tudod a platformot, add meg: `platform: shopify`.

## Fontos tudnivalók

- Az oldalak felhasználási feltételei tilthatják az automatikus lekérdezést; ezt oldalanként érdemes megnézni.
- Privát repó ajánlott: a GitHub Actions ingyenes kerete privát repónál havi 2000 perc, ez napi 3 futásnál is bőven elég.
- Publikus repóban az ütemezett workflow 60 nap repó-inaktivitás után leáll; mivel a script minden futáskor commitol, ez nálunk nem fordul elő.
