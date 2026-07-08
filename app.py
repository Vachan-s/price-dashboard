import asyncio
import csv
import json
import os
import re
import subprocess
import sys
from datetime import datetime, timezone
from io import StringIO

import openpyxl
from flask import Flask, jsonify, make_response, render_template
from playwright.async_api import async_playwright

from scrapers.amazon_search import scrape_amazon_price_with_page
from scrapers.myntra import scrape_myntra_price_with_page
from scrapers.tenxyou import scrape_tenxyou_price_with_page

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if sys.platform == "win32":
    subprocess.call(
        'for /f "tokens=5" %a in (\'netstat -ano ^| findstr :5000 ^| findstr LISTENING\') do taskkill /F /PID %a',
        shell=True,
    )

app = Flask(__name__)

XLSX_PATH        = "data/sku_mapping.xlsx"
URL_MAPPING_PATH = "data/url_mapping.xlsx"
RESULTS_PATH     = "data/results.json"

AMAZON_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

FUZZY_MATCH_THRESHOLD = 0.6

# ── Price helpers ────────────────────────────────────────────────────────────


def _parse_price(price_str) -> float | None:
    """Extract numeric value from a price string. Returns None if 0 or unparseable."""
    if isinstance(price_str, (int, float)):
        return float(price_str) if price_str else None
    # Match a proper number (with optional thousands commas and decimal fraction)
    # rather than stripping stray punctuation, since labels like "Rs." would
    # otherwise leave a leading decimal point behind (e.g. "Rs. 1199" → ".1199").
    m = re.search(r"\d[\d,]*\.\d+|\d[\d,]*", str(price_str or ""))
    if not m:
        return None
    try:
        v = float(m.group(0).replace(",", ""))
        return v if v else None
    except ValueError:
        return None


def _normalize_price(price_str) -> int | None:
    """Convert any price representation to a plain integer (INR). Returns None if 0 or unparseable."""
    v = _parse_price(price_str)
    return int(v) if v else None


def _norm_url(url: str) -> str:
    """Normalize a URL for equality comparison: strip query string and trailing slash."""
    return (url or "").split("?")[0].rstrip("/")


# ── Supplementary fuzzy-matching helpers (used only for products that aren't
# in url_mapping.xlsx yet — see _apply_supplementary_fuzzy_matching) ─────────


def _stem(word: str) -> str:
    return word[:-1] if word.endswith("s") and len(word) > 3 else word


def normalize(text: str) -> set:
    """Normalize a product name into a stemmed word set."""
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return {_stem(w) for w in text.split()}


def _fuzzy_score(a: str, b: str) -> float:
    wa, wb = normalize(a), normalize(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / len(wa | wb)


# ── URL mapping ──────────────────────────────────────────────────────────────


def _split_urls(cell) -> list[str]:
    """Split a (possibly comma-separated) cell into a list of clean, https:// URLs."""
    if not cell:
        return []
    text = str(cell).strip()
    if not text:
        return []
    urls = []
    for part in re.split(r"\s*,\s*", text):
        u = part.strip()
        if not u:
            continue
        if not re.match(r"^https?://", u, re.I):
            u = "https://" + u
        urls.append(u)
    return urls


def _load_url_mapping() -> dict:
    """Load data/url_mapping.xlsx into {SKU: {display_name, tenxyou, amazon, ajio, myntra}}.
    A SKU may span multiple rows (e.g. several TenXYou style variants sharing one
    SKU) — URLs from every matching row are pooled together into that SKU's lists."""
    mapping: dict = {}
    if not os.path.exists(URL_MAPPING_PATH):
        return mapping
    try:
        wb = openpyxl.load_workbook(URL_MAPPING_PATH, data_only=True)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        col = {h: i for i, h in enumerate(headers)}
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = row[col["SKU"]]
            if not sku:
                continue
            sku = str(sku).strip()
            if not sku:
                continue
            entry = mapping.setdefault(sku, {
                "display_name": "",
                "tenxyou": [], "amazon": [], "ajio": [], "myntra": [],
            })
            display_name = row[col["TenXYou Display Name"]]
            if display_name and not entry["display_name"]:
                entry["display_name"] = str(display_name).strip()
            entry["tenxyou"].extend(_split_urls(row[col["TenXYou URL"]]))
            entry["amazon"].extend(_split_urls(row[col["AMAZON"]]))
            entry["ajio"].extend(_split_urls(row[col["AJIO"]]))
            entry["myntra"].extend(_split_urls(row[col["MYNTRA"]]))
    except Exception as e:
        print(f"[URL MAPPING ERROR] {URL_MAPPING_PATH}: {e}", flush=True)
        return {}
    return mapping


def _load_sku_mapping_fallback() -> dict:
    """Load data/sku_mapping.xlsx as a fallback source for SKUs absent from
    url_mapping.xlsx — its own Myntra/Ajio/TenXYou URL columns (there's no
    Amazon column on this older sheet) become that SKU's scrape URLs."""
    fallback: dict = {}
    if not os.path.exists(XLSX_PATH):
        return fallback
    try:
        wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        col = {h: i for i, h in enumerate(headers)}
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = row[col["SKU"]]
            if not sku:
                continue
            sku = str(sku).strip()
            fallback[sku] = {
                "display_name": str(row[col["Product Name"]] or "").strip(),
                "tenxyou": _split_urls(row[col["TenXYou URL"]]),
                "amazon":  [],
                "ajio":    _split_urls(row[col["Ajio URL"]]),
                "myntra":  _split_urls(row[col["Myntra URL"]]),
            }
    except Exception as e:
        print(f"[SKU MAPPING FALLBACK ERROR] {XLSX_PATH}: {e}", flush=True)
        return {}
    return fallback


# ── Core async scrape ──────────────────────────────────────────────────────────

MYNTRA_CACHE_PATH  = "data/myntra_search_results.json"
AJIO_CACHE_PATH    = "data/ajio_search_results.json"
AMAZON_CACHE_PATH  = "data/amazon_search_results.json"


def _load_cache(path: str) -> list:
    """Load products from a JSON cache file. Returns [] if missing or empty."""
    try:
        if not os.path.exists(path):
            return []
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) and data else []
    except Exception as e:
        print(f"[CACHE ERROR] {path}: {e}", flush=True)
        return []


async def _scrape_lowest(opener, urls: list, scrape_fn, sku: str = "", debug_label: str = "") -> tuple:
    """Scrape every URL in `urls` (via a fresh page from `opener`, a Browser or
    BrowserContext) with scrape_fn(page, url), and return (lowest_price, its_url).
    Returns (None, None) if urls is empty or none scraped a price successfully.
    If debug_label is set, logs SKU/URL/error for every URL that fails to yield
    a price (used for Myntra debugging — see _run_scrape)."""
    if not urls:
        return None, None
    best_price, best_url = None, None
    for url in urls:
        page = await opener.new_page()
        error_msg = None
        try:
            result = await scrape_fn(page, url)
            if result.get("status") == "success":
                price = _normalize_price(result.get("price"))
                if price is None:
                    error_msg = f"no parsable price (raw price: {result.get('price')!r})"
            else:
                price = None
                error_msg = result.get("status")
        except Exception as e:
            print(f"[SCRAPE ERROR] {url}: {e}", flush=True)
            price = None
            error_msg = str(e)
        finally:
            await page.close()

        if price is None and debug_label:
            print(f"[{debug_label} DEBUG] SKU={sku}  URL={url}\n    -> {error_msg}", flush=True)

        if price is not None and (best_price is None or price < best_price):
            best_price, best_url = price, url
    return best_price, best_url


def _lookup_ajio_lowest(urls: list, ajio_by_url: dict) -> tuple:
    """Ajio product pages are blocked by Akamai, so look prices up from the
    cached search-results JSON instead of scraping the page directly."""
    if not urls:
        return None, None
    best_price, best_url = None, None
    for url in urls:
        product = ajio_by_url.get(_norm_url(url))
        if not product:
            continue
        price = _normalize_price(product.get("price"))
        if price is not None and (best_price is None or price < best_price):
            best_price, best_url = price, url
    return best_price, best_url


def _apply_supplementary_fuzzy_matching(results: list, raw_myntra: list, raw_ajio: list,
                                         raw_amazon: list, mapped_urls: dict) -> None:
    """For each platform, look at scraped search-results-cache products whose URL
    doesn't already appear in url_mapping.xlsx, and try to fuzzy-match them
    (word-overlap, >= FUZZY_MATCH_THRESHOLD confidence) against SKUs that still
    have no price for that platform. Mutates `results` in place. This is how new
    marketplace listings that aren't in the URL mapping sheet yet still show up."""
    platform_specs = [
        ("myntra", raw_myntra, "myntra_price", "matched_myntra_url", "myntra_fuzzy_match"),
        ("ajio",   raw_ajio,   "ajio_price",   "matched_ajio_url",   "ajio_fuzzy_match"),
        ("amazon", raw_amazon, "amazon_price", "matched_amazon_url", "amazon_fuzzy_match"),
    ]

    for platform, raw_products, price_key, url_key, flag_key in platform_specs:
        mapped = mapped_urls.get(platform, set())
        unmatched = [p for p in raw_products if _norm_url(p.get("url", "")) not in mapped]
        if not unmatched:
            continue

        candidates = [r for r in results if r.get(price_key) is None]
        for r in candidates:
            best_score, best_product = 0.0, None
            for p in unmatched:
                score = _fuzzy_score(r["name"], p.get("name", ""))
                if score > best_score:
                    best_score, best_product = score, p
            if best_product and best_score >= FUZZY_MATCH_THRESHOLD:
                price = _normalize_price(best_product.get("price"))
                if price is not None:
                    r[price_key] = price
                    r[url_key] = best_product["url"]
                    r[flag_key] = True


def _compute_status(r: dict) -> str:
    """complete — every platform with a mapped URL returned a price.
    partial — at least one price came back (URL-scraped or fuzzy-matched), but
    not every mapped platform succeeded.
    error — no prices at all.
    Platforms with no URL mapped (N/A) don't count against the SKU either way."""
    platforms = [
        (r["tenxyou_has_url"], r["tenxyou_price"]),
        (r["myntra_has_url"],  r["myntra_price"]),
        (r["ajio_has_url"],    r["ajio_price"]),
        (r["amazon_has_url"],  r["amazon_price"]),
    ]
    mapped = [price for has_url, price in platforms if has_url]
    got    = [price for price in mapped if price is not None]
    any_price = any(price is not None for _, price in platforms)

    if mapped and len(got) == len(mapped):
        return "complete"
    if any_price:
        return "partial"
    return "error"


async def _run_scrape(rows: list) -> list:
    """rows: [(sku, name), ...] from sku_mapping.xlsx — used as a display-name
    fallback and to pick up any SKUs that aren't in url_mapping.xlsx at all."""
    print("_RUN_SCRAPE CALLED", flush=True)

    url_mapping  = _load_url_mapping()
    sku_fallback = _load_sku_mapping_fallback()
    sheet_names  = {sku: name for sku, name in rows}

    raw_myntra = _load_cache(MYNTRA_CACHE_PATH)
    raw_ajio   = _load_cache(AJIO_CACHE_PATH)
    raw_amazon = _load_cache(AMAZON_CACHE_PATH)

    ajio_by_url = {_norm_url(p.get("url", "")): p for p in raw_ajio}

    print(f"[SCRAPE] URL mapping: {len(url_mapping)} SKUs loaded from {URL_MAPPING_PATH}", flush=True)
    print(f"[SCRAPE] Ajio cache:  {len(raw_ajio)} products loaded from {AJIO_CACHE_PATH}", flush=True)

    # url_mapping SKUs drive the loop; any sku_mapping-only SKU is appended as a fallback
    ordered_skus = list(url_mapping.keys())
    for sku in sheet_names:
        if sku not in url_mapping:
            ordered_skus.append(sku)

    results = []

    async with async_playwright() as p:
        tenxyou_browser = await p.chromium.launch(headless=True)
        # headless=False: Myntra hangs/times out on headless Chrome navigations
        # (confirmed — headed mode succeeds in <1s, headless hangs to timeout).
        myntra_browser = await p.chromium.launch(headless=False, args=["--disable-http2"])
        amazon_browser = await p.chromium.launch(
            headless=False, args=["--disable-blink-features=AutomationControlled"]
        )
        amazon_context = await amazon_browser.new_context(
            user_agent=AMAZON_USER_AGENT, viewport={"width": 1920, "height": 1080}
        )

        try:
            for sku in ordered_skus:
                entry = url_mapping.get(sku) or sku_fallback.get(sku) or {
                    "display_name": "", "tenxyou": [], "amazon": [], "ajio": [], "myntra": [],
                }
                name = entry["display_name"] or sheet_names.get(sku, sku)

                tenxyou_urls = entry["tenxyou"]
                myntra_urls  = entry["myntra"]
                ajio_urls    = entry["ajio"]
                amazon_urls  = entry["amazon"]

                tenxyou_price, matched_tenxyou_url = await _scrape_lowest(
                    tenxyou_browser, tenxyou_urls, scrape_tenxyou_price_with_page
                )
                myntra_price, matched_myntra_url = await _scrape_lowest(
                    myntra_browser, myntra_urls, scrape_myntra_price_with_page,
                    sku=sku, debug_label="MYNTRA",
                )
                amazon_price, matched_amazon_url = await _scrape_lowest(
                    amazon_context, amazon_urls, scrape_amazon_price_with_page
                )
                ajio_price, matched_ajio_url = _lookup_ajio_lowest(ajio_urls, ajio_by_url)

                print(
                    f"  {sku:<12} T={tenxyou_price} M={myntra_price} A={ajio_price} AZ={amazon_price}",
                    flush=True,
                )

                results.append({
                    "sku":                 sku,
                    "name":                name,
                    "tenxyou_price":       tenxyou_price,
                    "myntra_price":        myntra_price,
                    "ajio_price":          ajio_price,
                    "amazon_price":        amazon_price,
                    "matched_tenxyou_url": matched_tenxyou_url,
                    "matched_myntra_url":  matched_myntra_url,
                    "matched_ajio_url":    matched_ajio_url,
                    "matched_amazon_url":  matched_amazon_url,
                    "tenxyou_has_url":     bool(tenxyou_urls),
                    "myntra_has_url":      bool(myntra_urls),
                    "ajio_has_url":        bool(ajio_urls),
                    "amazon_has_url":      bool(amazon_urls),
                    "myntra_fuzzy_match":  False,
                    "ajio_fuzzy_match":    False,
                    "amazon_fuzzy_match":  False,
                })
        finally:
            await tenxyou_browser.close()
            await myntra_browser.close()
            await amazon_browser.close()

    mapped_urls = {
        "myntra": {_norm_url(u) for e in url_mapping.values() for u in e["myntra"]},
        "ajio":   {_norm_url(u) for e in url_mapping.values() for u in e["ajio"]},
        "amazon": {_norm_url(u) for e in url_mapping.values() for u in e["amazon"]},
    }
    _apply_supplementary_fuzzy_matching(results, raw_myntra, raw_ajio, raw_amazon, mapped_urls)

    for r in results:
        comp_prices = [p for p in [r["myntra_price"], r["ajio_price"], r["amazon_price"]] if p is not None]
        r["lowest_comp_price"] = min(comp_prices) if comp_prices else None
        r["status"] = _compute_status(r)

    return results


# ── Flask routes ───────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/scrape", methods=["POST"])
def scrape():
    print("SCRAPE ROUTE CALLED", flush=True)
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    sku_col  = headers.index("SKU")
    name_col = headers.index("Product Name")

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        sku = row[sku_col]
        if not sku:
            continue
        rows.append((sku, (row[name_col] or "").strip()))

    results = asyncio.run(_run_scrape(rows))

    output = {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "count": len(results),
        "results": results,
    }
    with open(RESULTS_PATH, "w") as f:
        json.dump(output, f, indent=2)

    return jsonify(output)


@app.route("/refresh-ajio", methods=["POST"])
def refresh_ajio():
    products = _load_cache(AJIO_CACHE_PATH)
    return jsonify({
        "count": len(products),
        "note": "Run `python scrapers/ajio_search.py` from terminal to refresh Ajio data.",
    })


@app.route("/results")
def results():
    if not os.path.exists(RESULTS_PATH):
        return jsonify({"timestamp": None, "count": 0, "results": []})
    with open(RESULTS_PATH) as f:
        return jsonify(json.load(f))


@app.route("/export")
def export():
    if not os.path.exists(RESULTS_PATH):
        return make_response("No results to export", 404)
    with open(RESULTS_PATH) as f:
        data = json.load(f)

    si = StringIO()
    writer = csv.DictWriter(
        si,
        fieldnames=["sku", "name", "tenxyou_price", "myntra_price", "ajio_price", "amazon_price",
                    "lowest_comp_price",
                    "matched_tenxyou_url", "matched_myntra_url", "matched_ajio_url", "matched_amazon_url",
                    "tenxyou_has_url", "myntra_has_url", "ajio_has_url", "amazon_has_url",
                    "myntra_fuzzy_match", "ajio_fuzzy_match", "amazon_fuzzy_match",
                    "status"],
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(data["results"])

    response = make_response(si.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=price_results.csv"
    response.headers["Content-Type"] = "text/csv"
    return response


if __name__ == "__main__":
    app.run(debug=True, use_reloader=False)
