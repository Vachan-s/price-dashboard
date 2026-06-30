import asyncio
import csv
import json
import os
from datetime import datetime, timezone
from io import StringIO

import openpyxl
from flask import Flask, jsonify, make_response, render_template

from scrapers.myntra import scrape_myntra_price
from scrapers.tenxyou import scrape_tenxyou_price

app = Flask(__name__)

XLSX_PATH = "data/sku_mapping.xlsx"
RESULTS_PATH = "data/results.json"


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/scrape", methods=["POST"])
def scrape():
    wb = openpyxl.load_workbook(XLSX_PATH)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    sku_col = headers.index("SKU")
    name_col = headers.index("Product Name")
    myntra_col = headers.index("Myntra URL")
    tenxyou_col = headers.index("TenXYou URL") if "TenXYou URL" in headers else None

    with open("data/debug.log", "w") as _dbg:
        _dbg.write(f"Columns: {headers}\n")
        for r in ws.iter_rows(min_row=2, max_row=2, values_only=True):
            _dbg.write(f"First row: {dict(zip(headers, r))}\n")

    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        myntra_url = row[myntra_col]
        if not myntra_url:
            continue

        tenxyou_url = row[tenxyou_col] if tenxyou_col is not None else None

        myntra_result = asyncio.run(scrape_myntra_price(myntra_url))
        tenxyou_result = (
            asyncio.run(scrape_tenxyou_price(tenxyou_url))
            if tenxyou_url
            else {"price": None, "status": "no url"}
        )

        m_ok = myntra_result["status"] == "success"
        t_ok = tenxyou_result["status"] == "success"
        status = "success" if (m_ok and t_ok) else ("partial" if (m_ok or t_ok) else "error")

        results.append({
            "sku": row[sku_col],
            "name": row[name_col],
            "myntra_price": myntra_result.get("price"),
            "tenxyou_price": tenxyou_result.get("price"),
            "myntra_url": myntra_url,
            "tenxyou_url": tenxyou_url,
            "status": status,
        })

    output = {
        "timestamp": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "count": len(results),
        "results": results,
    }
    with open(RESULTS_PATH, "w") as f:
        json.dump(output, f, indent=2)

    return jsonify(output)


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
        fieldnames=["sku", "name", "tenxyou_price", "myntra_price", "myntra_url", "tenxyou_url", "status"],
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
