# Price Dashboard — TenXYou vs Myntra & Ajio

A Flask-based price comparison dashboard that scrapes live prices for TenXYou products across Myntra and Ajio, then displays them alongside the brand's own prices for competitive analysis.

---

## What it does

- Scrapes product listings from Myntra and Ajio search pages using Playwright
- Scrapes TenXYou's own site via sitemap + Playwright
- Matches each SKU from the internal product catalogue to the closest marketplace listing using URL-slug fuzzy matching (Jaccard similarity)
- Displays TenXYou price, Myntra price, Ajio price, GAP % (vs lowest competitor), and match confidence per SKU in a web dashboard
- Exports results to CSV

---

## Setup

```bash
pip install flask playwright openpyxl requests playwright-stealth
playwright install chromium
```

Python 3.12+ recommended.

---

## Daily workflow

Run the three search scrapers to refresh the price cache, then trigger the dashboard scrape:

```bash
# 1. Refresh TenXYou prices (uses headless browser via sitemap)
python scrapers/tenxyou_search.py

# 2. Refresh Myntra prices (opens a visible Chrome window — let it scroll)
python scrapers/myntra_search.py

# 3. Refresh Ajio prices (opens a visible Chrome window — let it scroll)
python scrapers/ajio_search.py

# 4. Start the dashboard
python app.py
# → open http://localhost:5000 and click "Scrape Now"
```

Each scraper saves its results to `data/` as JSON. The dashboard reads those files when Scrape Now is clicked — no live scraping happens during the scrape button press.

---

## File structure

```
price-dashboard/
├── app.py                        # Flask app — routes, matching logic, cache loading
├── scrapers/
│   ├── myntra_search.py          # Scrapes Myntra search page → data/myntra_search_results.json
│   ├── ajio_search.py            # Scrapes Ajio search page  → data/ajio_search_results.json
│   ├── tenxyou_search.py         # Scrapes TenXYou sitemap   → data/tenxyou_search_results.json
│   ├── myntra.py                 # Single-product Myntra scraper (used for per-URL scraping)
│   └── tenxyou.py                # Single-product TenXYou scraper (used by tenxyou_search.py)
├── templates/
│   └── index.html                # Dashboard UI
├── data/
│   ├── sku_mapping.xlsx          # Master product list (SKU, Product Name, URLs)
│   ├── product-catalog.csv       # Reference catalogue with verified Myntra search names
│   ├── myntra-matched-results.csv  # Historical matched URLs from Myntra
│   └── ajio-matched-results.csv    # Historical matched URLs from Ajio
└── scripts/                      # One-off data scripts (URL matching, sheet population)
```

---

## Adding new products

1. Open `data/sku_mapping.xlsx`
2. Add a new row with the SKU code and Product Name
3. Re-run the three scrapers (step 1–3 above) to include the new product in matching
4. Click Scrape Now in the dashboard

The fuzzy matcher will attempt to find the closest listing on each marketplace based on the product name.

---

## Known limitations

- **Fuzzy matching is approximate** — match quality depends on how closely the marketplace listing name overlaps with the internal product name. Check `myntra_match_score` and `ajio_match_score` in `data/results.json` to see confidence per SKU (0–1 scale).
- **Ajio bot detection** — Ajio's search page blocks headless browsers. The scraper runs with `headless=False` (visible window). If it returns 0 products, try running it again or introducing a longer scroll wait in `scrapers/ajio_search.py`.
- **Myntra HTTP/2 errors** — Myntra may reject connections; the scraper uses `--disable-http2` to mitigate this. If it fails, try rerunning.
- **Prices update only when scrapers are re-run** — the dashboard does not scrape live on every button click. Run the scrapers first, then click Scrape Now.
- **Color variant deduplication** — multiple color variants of the same product are collapsed to a single representative price (mode across variants, falling back to median). The displayed price is one variant, not a range.
