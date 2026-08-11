# Consignment — Order Ledger

A personal dashboard for tracking e-commerce orders (Amazon, Flipkart, Myntra, Ajio, and beyond) across multiple retailers in one place — with net spend broken down by delivered / pending / returned.

![status](https://img.shields.io/badge/status-personal_project-blue)
![license](https://img.shields.io/badge/license-MIT-green)

## What's in here

| Path | What it is |
|---|---|
| `order-tracker.html` | The dashboard itself — a single-file HTML app. Open it directly in a browser, or use it as a Claude.ai artifact (it uses `window.storage` for persistence and an Anthropic API + Gmail MCP connector call for syncing). |
| `gmail-sync/` | A standalone Python script that reads your Gmail via the official Gmail API, parses order confirmation/shipping/delivery emails with regex heuristics, and outputs a CSV + a matching static HTML dashboard. |

These two are independent — you don't need both. Pick whichever fits how you want to run this.

## Dashboard features

- Net spend broken into **Kept (Delivered)**, **Pending (Ordered/Shipped/Out for Delivery)**, and **Returned/Cancelled/Refunded**
- Works with any retailer, not just the big four — unrecognized senders get their brand name auto-detected
- Manual entry for anything that isn't in your inbox
- Light/dark mode
- Filterable, searchable order table

## Setup — Python sync script (`gmail-sync/`)

See [`gmail-sync/SETUP.md`](gmail-sync/SETUP.md) for full step-by-step instructions. Short version:

```bash
cd gmail-sync
pip install -r requirements.txt
# create your own Gmail API OAuth credentials.json — see SETUP.md
python gmail_order_sync.py
```

This never uploads your data anywhere — everything runs and stays on your machine. The script only requests **read-only** Gmail access.

## Setup — HTML dashboard (`order-tracker.html`)

This file can be used two ways:

1. **Standalone / manual-entry only**: just open it in a browser. The Gmail auto-sync button won't work outside Claude.ai (it depends on Claude's Gmail MCP connector + API routing), but manual entry, filtering, search, and the spend breakdown all work fine on their own.
2. **Inside Claude.ai**: paste the file into a Claude.ai conversation as an artifact with a connected Gmail account, and the "Sync from Gmail" button will search and auto-populate orders.

## Privacy

- Nothing in this repo contains credentials, tokens, or personal order data.
- `credentials.json`, `token.json`, and any generated `orders.csv`/`dashboard.html` output from the Python script are gitignored — don't commit your own.
- The example files under `gmail-sync/examples/` use fake sample data for preview purposes only.

## License

MIT — see [LICENSE](LICENSE).
