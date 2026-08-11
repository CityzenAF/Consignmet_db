# Setup — Gmail Order Sync

## 1. Get Gmail API credentials (one-time, ~5 min)
1. Go to https://console.cloud.google.com/ → create a project (or use an existing one).
2. Search **"Gmail API"** in the top search bar → click **Enable**.
3. Go to **APIs & Services → Credentials → Create Credentials → OAuth client ID**.
   - If prompted, configure the consent screen first: choose **External**, fill in an app name (anything), your email, and add yourself as a test user.
   - Application type: **Desktop app**
   - Name it anything (e.g. "Order Sync")
4. Click **Create**, then **Download JSON**. Rename the downloaded file to `credentials.json` and put it in this same folder as `gmail_order_sync.py`.

## 2. Install dependencies
```bash
pip install -r requirements.txt
```

## 3. Run it
```bash
python gmail_order_sync.py
```
A browser tab opens asking you to sign in and approve **read-only** Gmail access. Approve it. This creates a `token.json` so you won't need to log in again next time.

## 4. See your orders
- `orders.csv` — raw parsed data, open in Excel/Sheets to fix anything the parser got wrong.
- `dashboard.html` — double-click to open in your browser. Filterable, searchable ledger view.

## Re-running later
Just run `python gmail_order_sync.py` again anytime — it re-scans and regenerates both files. Change `DAYS_BACK` near the top of the script to widen/narrow the window (default 90 days).

## If it finds 0 orders
- Check `credentials.json` is in this folder and you approved access in the browser popup.
- Widen `DAYS_BACK`.
- Open Gmail and manually search `from:amazon.in subject:order` (etc.) to confirm matching emails actually exist and aren't sitting in Trash/Spam (excluded by default).
- Print `full_query` inside `fetch_orders()` temporarily to see exactly what's being searched, and try that same string in Gmail's search bar.
