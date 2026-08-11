"""
Gmail Order Sync
================
Pulls order / shipping / delivery emails for Amazon, Flipkart, Myntra, and Ajio
out of your Gmail account, parses them, and writes:
  - orders.csv        (raw data, editable in Excel/Sheets)
  - dashboard.html     (a styled, filterable dashboard you can open in any browser)

ONE-TIME SETUP
--------------
1. Go to https://console.cloud.google.com/ and create a project (or reuse one).
2. APIs & Services -> Enable APIs and Services -> search "Gmail API" -> Enable.
3. APIs & Services -> Credentials -> Create Credentials -> OAuth client ID.
     Application type: Desktop app
   Download the JSON it gives you and save it as `credentials.json`
   in this same folder.
4. Install dependencies:
     pip install --upgrade google-auth google-auth-oauthlib google-api-python-client
5. Run:
     python gmail_order_sync.py
   A browser tab will open asking you to sign in and grant READ-ONLY Gmail
   access. After the first run, a `token.json` is cached locally so you won't
   need to log in again.

This script only requests gmail.readonly — it cannot send, delete, or modify
anything in your inbox. Nothing is uploaded anywhere; everything stays on
your machine.

RUNNING AGAIN LATER
--------------------
Just run `python gmail_order_sync.py` again — it re-scans the configured
window (default 90 days) and regenerates both output files. Adjust DAYS_BACK
below to widen or narrow the search.
"""

import os
import re
import csv
import json
import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
DAYS_BACK = 90
MAX_PER_PLATFORM = 150

PLATFORM_QUERIES = {
    "Amazon": '(from:amazon.in OR from:amazon.com OR from:shipment-tracking@amazon.in) '
              '(subject:order OR subject:shipped OR subject:delivered OR subject:dispatched OR subject:"out for delivery")',
    "Flipkart": 'from:flipkart.com '
                '(subject:order OR subject:shipped OR subject:delivered OR subject:dispatched OR subject:"out for delivery")',
    "Myntra": 'from:myntra.com '
              '(subject:order OR subject:shipped OR subject:delivered OR subject:dispatched OR subject:"out for delivery")',
    "Ajio": 'from:ajio.com '
            '(subject:order OR subject:shipped OR subject:delivered OR subject:dispatched OR subject:"out for delivery")',
}

STATUS_KEYWORDS = [
    ("Delivered", ["delivered", "has arrived"]),
    ("Out for Delivery", ["out for delivery"]),
    ("Shipped", ["shipped", "dispatched", "on its way", "on the way"]),
    ("Cancelled", ["cancelled", "canceled", "cancellation"]),
    ("Returned", ["returned", "refund initiated", "refunded"]),
    ("Ordered", ["order confirmed", "order placed", "thank you for your order",
                 "order confirmation", "we've received your order"]),
]

ORDER_ID_PATTERNS = [
    r"\b\d{3}-\d{7}-\d{7}\b",          # Amazon: 402-1234567-1234567
    r"\bOD\d{15,20}\b",                 # Flipkart: OD123456789012345
    r"\b[A-Z]{2,5}\d{4,18}\b",          # generic alnum order id (Myntra/Ajio style, etc.)
]

PRICE_PATTERN = r"(?:₹|Rs\.?|INR)\s?([\d,]+(?:\.\d{1,2})?)"


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
def get_gmail_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                raise FileNotFoundError(
                    "credentials.json not found. See the setup steps in the "
                    "docstring at the top of this file."
                )
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as f:
            f.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------
def detect_status(text):
    text_l = text.lower()
    for status, keywords in STATUS_KEYWORDS:
        if any(k in text_l for k in keywords):
            return status
    return "Unknown"


def extract_order_id(text):
    for pattern in ORDER_ID_PATTERNS:
        m = re.search(pattern, text)
        if m:
            return m.group(0)
    return None


def extract_price(text):
    m = re.search(PRICE_PATTERN, text)
    if m:
        try:
            return float(m.group(1).replace(",", ""))
        except ValueError:
            return None
    return None


BOILERPLATE = re.compile(
    r"(?i)\b(your|order|has been|shipped|delivered|confirmed|confirmation|"
    r"dispatched|out for delivery|amazon\.in|amazon\.com|flipkart|myntra|ajio|"
    r"is on (its|it's) way|update|notification|thank you|for shopping with us|"
    r"shopping with us|^for\b|\bis\b|\bfor\b|cancelled|canceled|cancellation|"
    r"returned|refund initiated|refunded)"
)

ORDER_ID_PAREN = re.compile(r"\(?\s*order\s*id\s*[:#]?\s*[A-Z0-9-]+\s*\)?", re.I)


def extract_item(subject):
    m = re.search(r'["“]([^"”]{4,80})["”]', subject)
    if m:
        return m.group(1).strip()
    cleaned = ORDER_ID_PAREN.sub("", subject)
    cleaned = re.sub(PRICE_PATTERN, "", cleaned, flags=re.I)
    cleaned = BOILERPLATE.sub("", cleaned)
    cleaned = re.sub(r"[:#|!]+", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" -–—.")
    return cleaned[:80] if cleaned else subject[:80]


def parse_email_date(date_header):
    if not date_header:
        return None
    # Trim trailing timezone name in parens, e.g. "(UTC)"
    cleaned = re.sub(r"\s*\([^)]*\)\s*$", "", date_header.strip())
    fmts = ["%a, %d %b %Y %H:%M:%S %z", "%d %b %Y %H:%M:%S %z"]
    for fmt in fmts:
        try:
            return datetime.datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Fetch + parse
# ---------------------------------------------------------------------------
def fetch_orders():
    service = get_gmail_service()
    after_date = (datetime.date.today() - datetime.timedelta(days=DAYS_BACK)).strftime("%Y/%m/%d")
    orders = []
    seen = set()

    for platform, query in PLATFORM_QUERIES.items():
        full_query = f"{query} after:{after_date}"
        print(f"Searching {platform}...")

        messages = []
        page_token = None
        while True:
            resp = service.users().messages().list(
                userId="me", q=full_query, maxResults=50, pageToken=page_token
            ).execute()
            messages.extend(resp.get("messages", []))
            page_token = resp.get("nextPageToken")
            if not page_token or len(messages) >= MAX_PER_PLATFORM:
                break

        print(f"  found {len(messages)} matching emails")

        for m in messages:
            msg = service.users().messages().get(
                userId="me", id=m["id"], format="metadata",
                metadataHeaders=["Subject", "Date"]
            ).execute()
            headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
            subject = headers.get("Subject", "")
            date_header = headers.get("Date", "")
            snippet = msg.get("snippet", "")
            combined = f"{subject} {snippet}"

            parsed_date = parse_email_date(date_header)

            order = {
                "platform": platform,
                "item": extract_item(subject),
                "orderId": extract_order_id(combined),
                "price": extract_price(combined),
                "status": detect_status(combined),
                "orderDate": parsed_date.isoformat() if parsed_date else None,
                "subject": subject,
            }

            key = (platform, order["orderId"] or (order["item"], order["orderDate"]))
            if key in seen:
                continue
            seen.add(key)
            orders.append(order)

    orders.sort(key=lambda o: o["orderDate"] or "", reverse=True)
    return orders


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------
def write_csv(orders, path="orders.csv"):
    fields = ["platform", "item", "orderId", "price", "status", "orderDate", "subject"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for o in orders:
            writer.writerow(o)
    print(f"Wrote {len(orders)} orders to {path}")


DASHBOARD_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Consignment — Order Ledger</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{{
    --bg:#12151F; --surface:#1B1F2C; --surface-2:#232838; --line:#2E3446; --text:#ECEDF3; --muted:#8B90A5;
    --gold:#E3A54D; --teal:#4FB0A0; --rose:#E0637F;
    --amazon:#F0A94E; --flipkart:#5B8DEF; --myntra:#E85D8A; --ajio:#8B6BC7; --other:#7C8194;
  }}
  *{{box-sizing:border-box; margin:0; padding:0;}}
  body{{background:var(--bg); color:var(--text); font-family:'Inter',sans-serif; padding:32px 20px 80px;}}
  .wrap{{max-width:920px; margin:0 auto;}}
  h1{{font-family:'Space Grotesk',sans-serif; font-size:28px;}}
  .eyebrow{{font-family:'IBM Plex Mono',monospace; font-size:11px; letter-spacing:2.5px; text-transform:uppercase; color:var(--gold); margin-bottom:6px;}}
  .subtitle{{color:var(--muted); font-size:13.5px; margin:4px 0 24px;}}
  .summary{{display:grid; grid-template-columns:repeat(4,1fr); gap:1px; background:var(--line); border:1px solid var(--line); border-radius:10px; overflow:hidden; margin-bottom:22px;}}
  .summary .cell{{background:var(--surface); padding:16px 18px;}}
  .summary .label{{font-size:11px; letter-spacing:1.5px; text-transform:uppercase; color:var(--muted); margin-bottom:8px;}}
  .summary .value{{font-family:'Space Grotesk',sans-serif; font-size:22px; font-weight:600;}}
  .controls{{display:flex; gap:10px; margin-bottom:18px; flex-wrap:wrap;}}
  .tab{{font-family:'IBM Plex Mono',monospace; font-size:12px; padding:7px 12px; border-radius:100px; border:1px solid var(--line); background:transparent; color:var(--muted); cursor:pointer;}}
  .tab.active{{background:var(--surface-2); color:var(--text); border-color:#454d68;}}
  .search-input{{flex:1; min-width:160px; background:var(--surface); border:1px solid var(--line); border-radius:8px; padding:9px 12px; color:var(--text); font-size:13px;}}
  .order-list{{display:flex; flex-direction:column; gap:10px;}}
  .order-card{{position:relative; background:var(--surface); border:1px solid var(--line); border-radius:10px; padding:16px 18px; display:flex; align-items:center; gap:14px;}}
  .order-card::before{{content:""; position:absolute; left:0; top:0; bottom:0; width:4px; border-radius:4px 0 0 4px; background:var(--platform-color, var(--other));}}
  .order-main{{flex:1; min-width:0;}}
  .order-top{{display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-bottom:4px;}}
  .platform-tag{{font-family:'IBM Plex Mono',monospace; font-size:10.5px; letter-spacing:1px; text-transform:uppercase; padding:2px 7px; border-radius:4px; color:#0e0e12; background:var(--platform-color, var(--other)); font-weight:600;}}
  .order-item{{font-weight:600; font-size:14.5px;}}
  .order-meta{{font-family:'IBM Plex Mono',monospace; font-size:11.5px; color:var(--muted); display:flex; gap:12px; flex-wrap:wrap;}}
  .order-right{{display:flex; align-items:center; gap:14px; flex-shrink:0;}}
  .order-price{{font-family:'IBM Plex Mono',monospace; font-size:15px; font-weight:600;}}
  .status-pill{{font-size:11px; padding:4px 10px; border-radius:100px; font-weight:600; white-space:nowrap;}}
  .status-Delivered{{background:rgba(79,176,160,0.15); color:var(--teal);}}
  .status-Shipped, .status-OutforDelivery{{background:rgba(227,165,77,0.15); color:var(--gold);}}
  .status-Ordered{{background:rgba(91,141,239,0.15); color:var(--flipkart);}}
  .status-Cancelled, .status-Returned{{background:rgba(224,99,127,0.15); color:var(--rose);}}
  .status-Unknown{{background:rgba(140,145,165,0.15); color:var(--muted);}}
  .empty{{text-align:center; padding:60px 20px; color:var(--muted); border:1px dashed var(--line); border-radius:10px;}}
  .note{{font-size:12px; color:var(--muted); margin-bottom:18px;}}
</style>
</head>
<body>
<div class="wrap">
  <div class="eyebrow">Order Ledger</div>
  <h1>Consignment</h1>
  <div class="subtitle">Generated {generated_at} — {total} orders across {days} days.</div>
  <div class="note">Parsed from Gmail with regex heuristics — double check prices/order IDs against orders.csv if something looks off. Re-run gmail_order_sync.py to refresh.</div>
  <div id="summary"></div>
  <div class="controls">
    <div id="tabs"></div>
    <input class="search-input" id="search" placeholder="Search item or order ID…">
  </div>
  <div id="list"></div>
</div>
<script>
const ORDERS = {orders_json};
const PLATFORM_COLORS = {{Amazon:'var(--amazon)',Flipkart:'var(--flipkart)',Myntra:'var(--myntra)',Ajio:'var(--ajio)',Other:'var(--other)'}};
const PLATFORMS = ['Amazon','Flipkart','Myntra','Ajio','Other'];
let filter = 'All';
let search = '';

function esc(s) {{ const d=document.createElement('div'); d.textContent=s||''; return d.innerHTML; }}

function renderSummary() {{
  const total = ORDERS.length;
  const spent = ORDERS.reduce((s,o)=>s+(o.price||0),0);
  const delivered = ORDERS.filter(o=>o.status==='Delivered').length;
  const platforms = new Set(ORDERS.map(o=>o.platform)).size;
  document.getElementById('summary').innerHTML = `
    <div class="summary">
      <div class="cell"><div class="label">Total Orders</div><div class="value">${{total}}</div></div>
      <div class="cell"><div class="label">Total Spent</div><div class="value">₹${{spent.toLocaleString('en-IN')}}</div></div>
      <div class="cell"><div class="label">Delivered</div><div class="value">${{delivered}}</div></div>
      <div class="cell"><div class="label">Platforms</div><div class="value">${{platforms}}</div></div>
    </div>`;
}}

function renderTabs() {{
  const all = ['All', ...PLATFORMS];
  document.getElementById('tabs').innerHTML = all.map(p =>
    `<button class="tab ${{filter===p?'active':''}}" onclick="setFilter('${{p}}')">${{p}}</button>`
  ).join(' ');
}}

function setFilter(p) {{ filter = p; renderTabs(); renderList(); }}

function renderList() {{
  let list = ORDERS.slice();
  if (filter !== 'All') list = list.filter(o => o.platform === filter);
  if (search.trim()) {{
    const q = search.toLowerCase();
    list = list.filter(o => (o.item||'').toLowerCase().includes(q) || (o.orderId||'').toLowerCase().includes(q));
  }}
  const el = document.getElementById('list');
  if (list.length === 0) {{
    el.innerHTML = `<div class="empty">No orders match this filter.</div>`;
    return;
  }}
  el.innerHTML = `<div class="order-list">${{list.map(o => `
    <div class="order-card" style="--platform-color:${{PLATFORM_COLORS[o.platform]||PLATFORM_COLORS.Other}}">
      <div class="order-main">
        <div class="order-top">
          <span class="platform-tag">${{esc(o.platform)}}</span>
          <span class="order-item">${{esc(o.item)}}</span>
        </div>
        <div class="order-meta">
          <span>${{esc(o.orderDate||'—')}}</span>
          ${{o.orderId ? `<span>#${{esc(o.orderId)}}</span>` : ''}}
        </div>
      </div>
      <div class="order-right">
        <span class="status-pill status-${{(o.status||'Unknown').replace(/\\s+/g,'')}}">${{esc(o.status||'Unknown')}}</span>
        <span class="order-price">${{o.price!=null ? '₹'+o.price.toLocaleString('en-IN') : '—'}}</span>
      </div>
    </div>
  `).join('')}}</div>`;
}}

document.getElementById('search').addEventListener('input', e => {{ search = e.target.value; renderList(); }});
renderSummary(); renderTabs(); renderList();
</script>
</body>
</html>
"""


def write_dashboard(orders, path="dashboard.html"):
    html = DASHBOARD_TEMPLATE.format(
        generated_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        total=len(orders),
        days=DAYS_BACK,
        orders_json=json.dumps(orders, ensure_ascii=False),
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Wrote dashboard to {path} — open it in your browser.")


# ---------------------------------------------------------------------------
if __name__ == "__main__":
    orders = fetch_orders()
    if not orders:
        print(
            "\nNo orders found. If this seems wrong, try:\n"
            "  - widening DAYS_BACK at the top of this file\n"
            "  - checking that emails aren't in Trash/Spam (excluded by default)\n"
            "  - running with a broader query, e.g. just from:amazon.in with no subject filter\n"
        )
    write_csv(orders)
    write_dashboard(orders)
