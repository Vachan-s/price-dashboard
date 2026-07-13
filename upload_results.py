"""
Run this right after your local "Scrape Now" finishes to push the fresh
data/results.json to the deployed dashboard — no PythonAnywhere account
needed, just the shared secret key.

Setup (one-time per person):
  1. Set the DASHBOARD_URL below to your deployed dashboard's address,
     e.g. "https://yourusername.pythonanywhere.com"
  2. Set UPLOAD_KEY to the same secret configured on the server
     (ask whoever set up the deployment for this key, or set it yourself
     via PythonAnywhere's Web tab -> Environment variables ->
     DASHBOARD_UPLOAD_SECRET).
  3. Both values can also be set as environment variables instead of
     editing this file directly:
       set DASHBOARD_URL=https://yourusername.pythonanywhere.com
       set DASHBOARD_UPLOAD_KEY=your-secret-key

Usage:
  python upload_results.py
"""

import os
import sys

import requests

DASHBOARD_URL = os.environ.get("DASHBOARD_URL", "https://yourusername.pythonanywhere.com")
UPLOAD_KEY    = os.environ.get("DASHBOARD_UPLOAD_KEY", "change-this-secret-before-deploying")
LOCAL_RESULTS_PATH = "data/results.json"


def main():
    if not os.path.exists(LOCAL_RESULTS_PATH):
        print(f"Local file not found: {LOCAL_RESULTS_PATH}")
        print("Run a scrape first (Scrape Now) before uploading.")
        sys.exit(1)

    if DASHBOARD_URL == "https://yourusername.pythonanywhere.com":
        print("Set DASHBOARD_URL first (edit this script or set the environment variable).")
        sys.exit(1)

    if UPLOAD_KEY == "change-this-secret-before-deploying":
        print("Set DASHBOARD_UPLOAD_KEY first — ask whoever deployed the dashboard for the secret.")
        sys.exit(1)

    url = f"{DASHBOARD_URL.rstrip('/')}/upload"
    with open(LOCAL_RESULTS_PATH, "rb") as f:
        resp = requests.post(
            url,
            headers={"X-Upload-Key": UPLOAD_KEY},
            files={"file": ("results.json", f, "application/json")},
            timeout=30,
        )

    if resp.status_code == 200:
        data = resp.json()
        print(f"Success — {data['count']} products uploaded, dashboard updated at {data['uploaded_at']}.")
    else:
        print(f"Upload failed ({resp.status_code}): {resp.text}")
        sys.exit(1)


if __name__ == "__main__":
    main()
