# Weekly Sales Velocity Report

Generates a weekly sales velocity report from Toast POS exports, emails it, and syncs data to Google Sheets.

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure your environment

```bash
cp .env.example .env
```

Edit `.env` and fill in:
- `GMAIL_SENDER` — the Gmail account used to send reports
- `GMAIL_RECIPIENT` — where reports are sent
- `GOOGLE_SHEET_ID` — the ID from your Google Sheet's URL (`/spreadsheets/d/THIS_PART/edit`)

### 3. Configure your locations and sheet structure

```bash
cp config.example.json config.json
```

Edit `config.json` to match your setup:
- `locations` — location display names mapped to CSV filenames
- `excluded_items` — Toast menu items to filter out (food, merch, etc.)
- `beer_name_aliases` — when a Toast item name differs from your sheet column header
- `location_section_headers` — how each location maps to its section header in the sheet
- `cans_section_header` — the section header in the Cans Inventory tab
- `can_name_aliases` — when a Toast can name differs from your sheet column header

### 4. Set up Google Cloud credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a project, enable the **Gmail API** and **Google Sheets API**
3. Create an OAuth 2.0 Client ID (Desktop app)
4. Download the credentials JSON and save it as `credentials.json` in this directory
5. On first run, a browser window will open to authorize access — `token.json` is saved after that

## Running the report

1. Export the last 7 days of Item Sales from Toast Back Office for each location
2. Rename each file to match your `config.json` (e.g., `locust_point.csv`, `timonium.csv`)
3. Drop them into `.tmp/`
4. Run:

```bash
python3 tools/run_report.py
```

Use `--overwrite` to replace the last filled week in the sheet instead of appending:

```bash
python3 tools/run_report.py --overwrite
```

## Deploying to Streamlit Cloud (shared dashboard)

Streamlit Cloud lets your team access the report tool from a URL — no local setup needed.

### 1. Set up a Google service account

Service accounts authenticate headlessly (no browser popup), which is required for hosted deployments.

1. Go to [Google Cloud Console](https://console.cloud.google.com) → **IAM & Admin → Service Accounts**
2. Create a service account, then create a JSON key for it
3. Open your Google Sheet → **Share** → add the service account's email address as an **Editor**

### 2. Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and connect your GitHub repo
2. Set **Main file path** to `app.py`
3. Under **Settings → Secrets**, add the following:

```toml
GOOGLE_SHEET_ID = "your-sheet-id"
GOOGLE_SERVICE_ACCOUNT_JSON = '{"type": "service_account", ...}'  # full key JSON, single line
INVENTORY_CONFIG = '{"locations": {...}, ...}'  # full config.json contents, single line
```

> **Tip:** To get `GOOGLE_SERVICE_ACCOUNT_JSON` as a single line, run:
> `python3 -c "import json; print(json.dumps(json.load(open('your-key.json'))))"`
>
> Same for `INVENTORY_CONFIG` from your `config.json`.

### 3. Use it

Navigate to your Streamlit URL, upload the CSVs for each location, and click **Run Report**. Data syncs directly to Google Sheets.

---

## How to export from Toast

1. Log in to Toast Back Office (pos.toasttab.com)
2. Navigate to **Reports → Menu Item Reports → Item Sales**
3. Set date range to the previous 7 days
4. Click **Export** → download CSV
5. Rename and drop into `.tmp/`
