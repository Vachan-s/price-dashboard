# Deploying the Viewer Dashboard on PythonAnywhere (Free)

This viewer app has **no scraping code** — it only reads `data/results.json`
and displays it. That's what makes it safe and easy to host for free.

---

## One-time setup (~15 minutes)

### 1. Create a free PythonAnywhere account
Go to https://www.pythonanywhere.com/ → **Pricing & signup** → **Create a
Beginner account** (free, no credit card).

### 2. Upload the viewer app files
On the PythonAnywhere dashboard:
1. Click **Files** (top menu)
2. Create a folder, e.g. `price-dashboard-viewer`
3. Upload these 4 files/folders into it, keeping the same structure:
   ```
   price-dashboard-viewer/
   ├── app.py
   ├── requirements.txt
   ├── templates/
   │   └── index.html
   └── data/
       └── results.json
   ```
   (You can drag-and-drop files, or use "Upload a file" button per file.)

### 3. Install Flask
Click **Consoles** → **Bash** → run:
```bash
pip install --user flask
```

### 4. Create the web app
1. Click **Web** (top menu) → **Add a new web app**
2. Choose **Flask** and the Python version shown (any recent 3.x is fine)
3. When asked for the path, point it to your uploaded `app.py`
   (e.g. `/home/yourusername/price-dashboard-viewer/app.py`)
4. PythonAnywhere auto-generates a WSGI configuration file — open it
   (link is on the **Web** tab) and make sure the path at the top points to
   your project folder, e.g.:
   ```python
   import sys
   path = '/home/yourusername/price-dashboard-viewer'
   if path not in sys.path:
       sys.path.append(path)
   from app import app as application
   ```
5. Click the big green **Reload** button on the Web tab.

### 5. Set the shared upload secret
This lets ANY team member push fresh data without needing a PythonAnywhere
login of their own.

1. On the **Web** tab, scroll to **Environment variables**
2. Add: `DASHBOARD_UPLOAD_SECRET` = *(pick any random password-like string,
   e.g. `tenxyou-dash-8k2m9x`)*
3. Click **Reload** on the Web tab for it to take effect

Share this secret with whoever on the team needs to be able to upload fresh
scrapes (via a password manager or however you share sensitive info
internally — not over plain chat/email ideally).

### 6. Visit your dashboard
Your app is now live at:
```
https://yourusername.pythonanywhere.com
```
Share this link with the team — no login needed to **view** it.

---

## Weekly update workflow — team-friendly, no PythonAnywhere account needed

Anyone on the team who has (a) run the local scraper and (b) the shared
secret from Step 5 can push fresh data themselves:

1. Run the full app as usual: `python app.py` → click **Scrape Now**
2. Make sure `requests` is installed locally (one-time):
   `pip install requests`
3. Set two environment variables (one-time, per person, per machine):
   ```bash
   set DASHBOARD_URL=https://yourusername.pythonanywhere.com
   set DASHBOARD_UPLOAD_KEY=the-secret-from-step-5
   ```
4. Run:
   ```bash
   python upload_results.py
   ```
5. That's it — no PythonAnywhere login, no manual file upload through their
   website. The dashboard updates immediately; anyone with the link sees
   the fresh prices on next refresh.

This means scraping and uploading can happen from **any team member's
machine** — nothing is tied to one specific person or computer.

---

## Notes

- **Free tier limits**: PythonAnywhere's free tier gives you one always-on
  web app with a `yourusername.pythonanywhere.com` subdomain — no sleep/cold
  start issue, unlike some other free hosts.
- **No scraping runs on PythonAnywhere** — all scraping still happens on
  whichever team member's machine runs it. This viewer only ever reads the
  one JSON file, updated via the `/upload` endpoint.
- **Security note**: the upload secret is a simple shared-password style
  protection — adequate for an internal team tool, not meant for
  public-internet-facing sensitive data. Don't reuse a real account
  password as this secret; treat it as a disposable shared key you can
  rotate anytime by changing the environment variable and telling the team.
- The old `push_results.py` (PythonAnywhere API token method) still works
  as an alternative if preferred, but requires a personal API token per
  person rather than one shared secret — `upload_results.py` is simpler for
  a small team.
