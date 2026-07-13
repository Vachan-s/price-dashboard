"""
TenXYou Price Dashboard — VIEWER ONLY

This is a stripped-down version of the full dashboard app, meant to be
deployed on free hosting (e.g. PythonAnywhere). It has NO scraping code at
all — no Playwright, no browser automation. It only reads and displays
whatever is in data/results.json.

Weekly workflow:
  1. Run the FULL app (with scraping) locally on your own machine as usual.
  2. Copy the fresh data/results.json produced by that run.
  3. Upload/overwrite data/results.json in this deployed app's file storage.
  4. Anyone with the hosted URL immediately sees the fresh prices — no
     redeploy needed, just overwrite the one JSON file.
"""

import json
import os
from datetime import datetime, timezone
from io import StringIO
import csv

from flask import Flask, jsonify, make_response, render_template, request

app = Flask(__name__)

RESULTS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "results.json")

# Shared secret for the /upload endpoint. Anyone with this key can push a
# fresh results.json — no PythonAnywhere account needed. Set this as an
# environment variable on PythonAnywhere (Web tab -> "Environment variables")
# rather than leaving the default in place.
UPLOAD_SECRET = os.environ.get("DASHBOARD_UPLOAD_SECRET", "change-this-secret-before-deploying")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/results")
def results():
    if not os.path.exists(RESULTS_PATH):
        return jsonify({"timestamp": None, "count": 0, "results": []})
    with open(RESULTS_PATH, encoding="utf-8") as f:
        return jsonify(json.load(f))


@app.route("/export")
def export():
    if not os.path.exists(RESULTS_PATH):
        return make_response("No results available yet.", 404)
    with open(RESULTS_PATH, encoding="utf-8") as f:
        data = json.load(f)

    si = StringIO()
    writer = csv.DictWriter(
        si,
        fieldnames=["sku", "name", "tenxyou_price", "myntra_price", "ajio_price", "amazon_price",
                    "lowest_comp_price", "status"],
        extrasaction="ignore",
    )
    writer.writeheader()
    writer.writerows(data.get("results", []))

    response = make_response(si.getvalue())
    response.headers["Content-Disposition"] = "attachment; filename=price_results.csv"
    response.headers["Content-Type"] = "text/csv"
    return response


@app.route("/upload", methods=["POST"])
def upload():
    """Team-friendly upload endpoint — any team member can push a fresh
    results.json here (via upload_results.py) using the shared secret key,
    with no PythonAnywhere account of their own required.
    """
    provided_key = request.headers.get("X-Upload-Key", "")
    if provided_key != UPLOAD_SECRET:
        return jsonify({"error": "Invalid or missing upload key"}), 401

    if "file" not in request.files:
        return jsonify({"error": "No file provided (expected form field 'file')"}), 400

    file = request.files["file"]
    try:
        raw = file.read()
        data = json.loads(raw)
    except Exception as e:
        return jsonify({"error": f"Uploaded file is not valid JSON: {e}"}), 400

    # Basic structure validation so a malformed file can't silently corrupt
    # the live dashboard.
    if not isinstance(data, dict) or "results" not in data or not isinstance(data["results"], list):
        return jsonify({"error": "JSON must be an object with a 'results' list — got something else"}), 400

    os.makedirs(os.path.dirname(RESULTS_PATH), exist_ok=True)
    with open(RESULTS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

    return jsonify({
        "status": "ok",
        "count": len(data["results"]),
        "uploaded_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    })


if __name__ == "__main__":
    # Local test run only. On PythonAnywhere, the WSGI file imports `app`
    # directly and this block never executes.
    app.run(host="0.0.0.0", port=5000, debug=True)
