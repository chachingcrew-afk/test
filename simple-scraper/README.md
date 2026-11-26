# Simple Price Scraper (GitHub Actions + Webshare + Supabase)

Ez a repó gyökerében lévő "Track prices" mellett fut, attól független. Workflow: `.github/workflows/scrape.yml` ("Price Scraper (simple)"). Itt minden fájl ebben a `simple-scraper/` mappában van.

Naponta lefut a GitHub szerverein, minden kérés a Webshare rotating proxyn megy ki (új IP kérésenként). Az eredmény a Supabase `product_prices` táblájába kerül, és CSV mentés is készül a futás oldalán letölthető artifactként.

## Beállítás

1. **Supabase**: SQL Editor-ban futtasd le a `schema.sql`-t.
2. **GitHub Secrets** (repo > Settings > Secrets and variables > Actions):
   - `WEBSHARE_USER`, `WEBSHARE_PASS` (Webshare > Proxy List > Rotating Proxy Endpoint; a `-rotate` végződést a script hozzáteszi)
   - `SUPABASE_URL`, `SUPABASE_KEY` (Project Settings > API, a service_role kulcs)
3. A `config.json`-ban maradjanak a `YOUR_...` placeholderek. A Secrets mindig felülírja őket, a valódi kulcs soha ne kerüljön a repóba.
4. **products.xlsx** oszlopai: `name`, `ean`, `url`, opcionálisan `price_selector` és `currency`.
5. **Teszt**: Actions > Price Scraper > Run workflow.

## Holland IP

A `config.json`-ban `"country": "nl"` van beállítva, így a script a Webshare usernevet `felhasználónév-nl-rotate` formában küldi: minden kérés más holland IP-ről megy ki. A log elején látszik a kimenő IP és az országa. Ha nem NL jön, ellenőrizd, hogy a Webshare proxy listádban van-e holland proxy. Más országhoz írd át a kódot (pl. `"de"`), vagy üresen hagyva bármelyik ország jöhet.

## Ár keresés sorrendje

1. a saját `price_selector` (ha meg van adva; vesszővel több is lehet)
2. JSON-LD, meta tagek, WooCommerce `.price`
3. Shopify `/products/<handle>.js`
4. általános ár-keresés

A `raw_text` oszlopban látod, honnan és milyen szövegből jött az ár.
