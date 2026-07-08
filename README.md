# TenXYou Price Dashboard

Internal price monitoring dashboard comparing TenXYou D2C prices against Myntra, Ajio, and Amazon.

---

## Daily Workflow

1. Start the app:
   ```bash
   cd "c:\Users\vacha\Desktop\I'\price-dashboard"
   python app.py
   ```
2. Open browser: [http://127.0.0.1:5000](http://127.0.0.1:5000)
3. Click **Scrape Now** — takes ~7 minutes, scrapes all platforms automatically

---

## Refreshing Ajio & Amazon prices (weekly)

Ajio and Amazon prices are cached — refresh them occasionally:

```bash
python scrapers/ajio_search.py
python scrapers/amazon_search.py
```

Then click **Scrape Now** again.

---

## Adding new products

1. Open `data/url_mapping.xlsx`
2. Add a new row: TenXYou Display Name, SKU, TenXYou URL, Myntra URL, Ajio URL, Amazon URL
3. Leave blank any platform where the product isn't listed
4. Click **Scrape Now** — new product will appear automatically

---

## File Structure

- `app.py` — Flask app, main scraping logic
- `templates/index.html` — Dashboard UI
- `scrapers/` — Individual scrapers for each platform
- `data/url_mapping.xlsx` — Master URL mapping (edit this to add/remove products)
- `data/sku_mapping.xlsx` — Fallback SKU reference
- `data/*_search_results.json` — Cached search results for Ajio/Amazon

---

## Known Limitations

- Ajio individual product pages are blocked by bot detection — prices come from cached search results
- Amazon prices cached from search results — refresh weekly for accuracy
- Socks, insoles, spikes, signed merchandise not listed on any marketplace (shows N/A)
- Tall/Regular variants share marketplace listings — same price shown for both

---

## Troubleshooting

- **Port 5000 already in use**: Run `netstat -ano | findstr :5000` then `taskkill /F /PID <pid>`
- **Prices not updating**: Make sure to restart app after any code changes
- **Ajio/Amazon showing stale prices**: Run the refresh commands above
