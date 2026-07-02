import asyncio
import csv
import json
import os
import re
import subprocess
import sys
from collections import Counter
from datetime import datetime, timezone
from io import StringIO

import openpyxl
from flask import Flask, jsonify, make_response, render_template

os.chdir(os.path.dirname(os.path.abspath(__file__)))

if sys.platform == "win32":
    subprocess.call(
        'for /f "tokens=5" %a in (\'netstat -ano ^| findstr :5000 ^| findstr LISTENING\') do taskkill /F /PID %a',
        shell=True,
    )

app = Flask(__name__)

XLSX_PATH = "data/sku_mapping.xlsx"
RESULTS_PATH = "data/results.json"
# ── Matching helpers ───────────────────────────────────────────────────────────

_BRAND_WORDS = {"ten", "x", "you"}


def _stem(word: str) -> str:
    return word[:-1] if word.endswith("s") and len(word) > 3 else word


def normalize(text: str) -> set:
    """Normalize a SKU product name into a stemmed word set."""
    text = (text or "").lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    return {_stem(w) for w in text.split()}


def _slug_words(url: str) -> set:
    """Extract meaningful words from the product-name portion of a URL slug."""
    path = re.sub(r"https?://[^/]+", "", url.split("?")[0].rstrip("/"))
    words = set()
    for part in re.split(r"[/\-]", path):
        p = part.lower()
        if len(p) >= 3 and not p.isdigit() and p not in _BRAND_WORDS:
            words.add(_stem(p))
    return words


def _slug_base(url: str) -> str:
    """Return a canonical key for grouping color/size variants of the same product."""
    url = url.split("?")[0]
    # Ajio: .../p/469817206_green → "469817206"
    m = re.search(r"/p/(\d+)", url)
    if m:
        return m.group(1)
    # Amazon: .../dp/B0G53JQL2L/ref=sr_1_5 → "B0G53JQL2L"
    m = re.search(r"/dp/([A-Z0-9]{10})", url)
    if m:
        return m.group(1)
    # Myntra: .../product-slug/39062448/buy → "product-slug"
    parts = [p for p in url.split("/") if p]
    for i, part in enumerate(parts):
        if part.isdigit() and i > 0:
            return parts[i - 1]
    return url


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


def dedup_products(products: list) -> list:
    """Group color/size variants by URL slug base; keep one entry per product.
    Price = mode across variants; if no clear mode, take the median."""
    from collections import defaultdict
    groups: dict = defaultdict(list)
    for p in products:
        groups[_slug_base(p.get("url", ""))].append(p)

    result = []
    for group in groups.values():
        parsed = [(_parse_price(p["price"]), p) for p in group]
        parsed = [(n, p) for n, p in parsed if n is not None]
        if not parsed:
            result.append(group[0])
            continue
        counts = Counter(n for n, _ in parsed)
        max_count = max(counts.values())
        modes = [n for n, c in counts.items() if c == max_count]
        if max_count > 1 and len(modes) == 1:
            rep_num = modes[0]
        else:
            rep_num = sorted(n for n, _ in parsed)[len(parsed) // 2]
        rep = next((p for n, p in parsed if n == rep_num), group[0])
        result.append(rep)
    return result


def find_best_match(product_name: str, candidates: list) -> tuple:
    """Match SKU name (normalized) against URL slug words of each candidate.
    Returns (best_candidate, score). Returns (None, 0.0) if all scores are 0."""
    target = normalize(product_name)
    best_score, best = 0.0, None
    for c in candidates:
        cand = _slug_words(c.get("url", "")) or normalize(c.get("name", ""))
        if target and cand:
            score = len(target & cand) / len(target | cand)
            if score > best_score:
                best_score, best = score, c
    return (best, round(best_score, 4)) if best_score > 0.0 else (None, 0.0)


def find_tenxyou_anchor(product_name: str, tenxyou_products: list) -> dict | None:
    """Find the closest TenXYou product by plain name word-overlap (no threshold —
    every candidate is already a TenXYou product, so even a weak match is useful
    as an anchor query for Myntra/Ajio matching)."""
    target = normalize(product_name)
    if not target or not tenxyou_products:
        return None
    best_score, best = -1.0, None
    for p in tenxyou_products:
        cand = normalize(p.get("name", ""))
        union = target | cand
        score = len(target & cand) / len(union) if union else 0.0
        if score > best_score:
            best_score, best = score, p
    return best


_CATEGORY_KEYWORDS = {
    "XU": ("shoe", "sneaker", "slides", "flip", "slipper", "clog", "sandal"),
    "XA": ("cap", "sock", "insole", "spike"),
}


def filter_by_category(sku: str, products: list) -> list:
    """Restrict candidates to those matching the SKU prefix's category. Unrecognized
    prefixes are left unfiltered."""
    prefix = (sku or "").strip().upper()[:2]

    def name_of(p):
        return (p.get("name") or "").lower()

    if prefix == "XM":
        return [p for p in products if "men" in name_of(p) and "women" not in name_of(p)]
    if prefix == "XW":
        return [p for p in products if "women" in name_of(p) or "woman" in name_of(p)]
    if prefix in _CATEGORY_KEYWORDS:
        keywords = _CATEGORY_KEYWORDS[prefix]
        return [p for p in products if any(k in name_of(p) for k in keywords)]
    return products


# ── Core async scrape ──────────────────────────────────────────────────────────

MYNTRA_CACHE_PATH   = "data/myntra_search_results.json"
AJIO_CACHE_PATH     = "data/ajio_search_results.json"
AMAZON_CACHE_PATH   = "data/amazon_search_results.json"
TENXYOU_CACHE_PATH  = "data/tenxyou_search_results.json"


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


async def _run_scrape(rows: list) -> list:
    print("_RUN_SCRAPE CALLED", flush=True)
    raw_myntra  = _load_cache(MYNTRA_CACHE_PATH)
    raw_ajio    = _load_cache(AJIO_CACHE_PATH)
    raw_amazon  = _load_cache(AMAZON_CACHE_PATH)
    raw_tenxyou = _load_cache(TENXYOU_CACHE_PATH)

    myntra_products  = dedup_products(raw_myntra)
    ajio_products    = dedup_products(raw_ajio)
    amazon_products  = dedup_products(raw_amazon)
    tenxyou_products = dedup_products(raw_tenxyou)

    print(f"[SCRAPE] Myntra:  {len(raw_myntra)} raw → {len(myntra_products)} deduped", flush=True)
    print(f"[SCRAPE] Ajio:    {len(raw_ajio)} raw → {len(ajio_products)} deduped", flush=True)
    print(f"[SCRAPE] Amazon:  {len(raw_amazon)} raw → {len(amazon_products)} deduped", flush=True)
    print(f"[SCRAPE] TenXYou: {len(raw_tenxyou)} raw → {len(tenxyou_products)} deduped", flush=True)

    results = []
    for sku, name in rows:
        tenxyou_match, t_score = find_best_match(name, tenxyou_products)

        anchor = find_tenxyou_anchor(name, tenxyou_products)
        query_name = anchor["name"] if anchor else name

        myntra_candidates = filter_by_category(sku, myntra_products)
        ajio_candidates   = filter_by_category(sku, ajio_products)
        amazon_candidates = filter_by_category(sku, amazon_products)

        myntra_match, m_score  = find_best_match(query_name, myntra_candidates)
        ajio_match,   a_score  = find_best_match(query_name, ajio_candidates)
        amazon_match, az_score = find_best_match(query_name, amazon_candidates)

        tenxyou_price       = _normalize_price(tenxyou_match["price"]) if tenxyou_match else None
        matched_tenxyou_url = tenxyou_match["url"]   if tenxyou_match else None
        myntra_price        = _normalize_price(myntra_match["price"])  if myntra_match  else None
        matched_myntra_url  = myntra_match["url"]    if myntra_match  else None
        ajio_price          = _normalize_price(ajio_match["price"])    if ajio_match    else None
        matched_ajio_url    = ajio_match["url"]      if ajio_match    else None
        amazon_price        = _normalize_price(amazon_match["price"])  if amazon_match  else None
        matched_amazon_url  = amazon_match["url"]    if amazon_match  else None

        comp_prices = [p for p in [myntra_price, ajio_price, amazon_price] if p is not None]
        lowest_comp_price = min(comp_prices) if comp_prices else None

        print(
            f"  {sku:<14} T={t_score:.2f} {(tenxyou_match['name'] if tenxyou_match else 'NO MATCH')[:30]}"
            f"  M={m_score:.2f} {(myntra_match['name'] if myntra_match else 'NO MATCH')[:25]}"
            f"  A={a_score:.2f} {(ajio_match['name'] if ajio_match else 'NO MATCH')[:25]}"
            f"  AZ={az_score:.2f} {(amazon_match['name'] if amazon_match else 'NO MATCH')[:25]}",
            flush=True,
        )

        t_ok  = tenxyou_price is not None
        m_ok  = myntra_price  is not None
        a_ok  = ajio_price    is not None
        az_ok = amazon_price  is not None

        if t_ok and (m_ok or a_ok or az_ok):
            status = "success"
        elif t_ok or m_ok or a_ok or az_ok:
            status = "partial"
        else:
            status = "error"

        results.append({
            "sku":                 sku,
            "name":                name,
            "tenxyou_price":       tenxyou_price,
            "myntra_price":        myntra_price,
            "ajio_price":          ajio_price,
            "amazon_price":        amazon_price,
            "lowest_comp_price":   lowest_comp_price,
            "matched_tenxyou_url": matched_tenxyou_url,
            "matched_myntra_url":  matched_myntra_url,
            "matched_ajio_url":    matched_ajio_url,
            "matched_amazon_url":  matched_amazon_url,
            "tenxyou_match_score": t_score,
            "myntra_match_score":  m_score,
            "ajio_match_score":    a_score,
            "amazon_match_score":  az_score,
            "status":              status,
        })

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
                    "tenxyou_match_score", "myntra_match_score", "ajio_match_score", "amazon_match_score",
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
