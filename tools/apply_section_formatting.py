#!/usr/bin/env python3
"""
Copy all cell formatting from the Total Volume (master) section to the
DBC Locust Point and DBC Timonium sections, row by row.

Rows are matched by their label in column A (e.g. "Beer:", "Keg Volume",
"Sales Week 1", "# of Pours Total", "Avg Weekly Sales", etc.) so the
mapping stays correct even if row positions shift between sections.

The section header row (e.g. "Total Volume") is always mapped to each
location's header row (e.g. "DBC Locust Point") regardless of differing text.

Only FORMAT is copied — values and formulas are never touched.

Usage:
  python apply_section_formatting.py           # dry run (preview)
  python apply_section_formatting.py --write   # apply changes
"""

import argparse
import json
import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
load_dotenv(BASE_DIR / ".env")

# Cover all columns used (A through AZ = 52 columns)
MAX_COL = 52


def load_config():
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        print("ERROR: config.json not found.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)


CONFIG = load_config()
MASTER_SECTION_HEADER = CONFIG.get("master_section_header")
LOCATION_SECTION_HEADERS = CONFIG["location_section_headers"]


def get_sheets_service():
    creds = None
    token_path = BASE_DIR / "token.json"
    creds_path = BASE_DIR / "credentials.json"
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())
    return build("sheets", "v4", credentials=creds)


def find_current_tab(service, sheet_id):
    from datetime import datetime
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    all_titles = [s["properties"]["title"] for s in meta["sheets"]]
    if "Total Inventory" in all_titles:
        return "Total Inventory"
    today = datetime.now()
    for delta in [0, -1]:
        month = today.month + delta
        year = today.year
        if month <= 0:
            month += 12
            year -= 1
        month_name = datetime(year, month, 1).strftime("%B %Y")
        title = f"Total Inventory - {month_name}"
        if title in all_titles:
            return title
    tabs = sorted(t for t in all_titles if t.startswith("Total Inventory -"))
    return tabs[-1] if tabs else None


def get_sheet_numeric_id(service, sheet_id, tab):
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab:
            return sheet["properties"]["sheetId"]
    return None


def read_all_rows(service, sheet_id, tab):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:AZ100",
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    return result.get("values", [])


def normalize(name):
    return name.replace("\u2019", "'").replace("\u2018", "'").strip()


def find_section_start(rows, header):
    for i, row in enumerate(rows):
        if row and normalize(row[0]) == normalize(header):
            return i
    return None


def find_section_end(rows, section_start, all_headers):
    """Return the row index where this section ends (start of next section, or end of data)."""
    for i in range(section_start + 1, len(rows)):
        if rows[i] and any(normalize(rows[i][0]) == normalize(h) for h in all_headers):
            return i
    # Find last non-empty row
    last = section_start
    for i in range(section_start, len(rows)):
        if rows[i]:
            last = i
    return last + 1


def build_label_map(rows, section_start, section_end):
    """
    Return {normalized_label: row_index} for rows in [section_start, section_end).
    Duplicate labels are stored as label, label__2, label__3, etc. to preserve all rows.
    """
    label_map = {}
    label_counts = {}
    for i in range(section_start, section_end):
        if i >= len(rows):
            break
        label = normalize(rows[i][0]) if rows[i] else ""
        count = label_counts.get(label, 0) + 1
        label_counts[label] = count
        key = label if count == 1 else f"{label}__{count}"
        label_map[key] = i
    return label_map


def make_copy_request(numeric_id, src_row, dst_row):
    return {
        "copyPaste": {
            "source": {
                "sheetId": numeric_id,
                "startRowIndex": src_row,
                "endRowIndex": src_row + 1,
                "startColumnIndex": 0,
                "endColumnIndex": MAX_COL,
            },
            "destination": {
                "sheetId": numeric_id,
                "startRowIndex": dst_row,
                "endRowIndex": dst_row + 1,
                "startColumnIndex": 0,
                "endColumnIndex": MAX_COL,
            },
            "pasteType": "PASTE_FORMAT",
            "pasteOrientation": "NORMAL",
        }
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true", help="Apply changes (default is dry run)")
    args = parser.parse_args()
    dry_run = not args.write

    if dry_run:
        print("DRY RUN — pass --write to apply changes\n")

    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("ERROR: GOOGLE_SHEET_ID not set in .env")
        sys.exit(1)

    service = get_sheets_service()
    tab = find_current_tab(service, sheet_id)
    if not tab:
        print("ERROR: No Total Inventory tab found.")
        sys.exit(1)
    print(f"Tab: '{tab}'\n")

    numeric_id = get_sheet_numeric_id(service, sheet_id, tab)
    rows = read_all_rows(service, sheet_id, tab)

    all_headers = [MASTER_SECTION_HEADER] + list(LOCATION_SECTION_HEADERS.values())

    # ── Locate master section ────────────────────────────────────────────────
    if not MASTER_SECTION_HEADER:
        print("ERROR: master_section_header not set in config.json")
        sys.exit(1)

    master_start = find_section_start(rows, MASTER_SECTION_HEADER)
    if master_start is None:
        print(f"ERROR: Master section '{MASTER_SECTION_HEADER}' not found.")
        sys.exit(1)

    master_end = find_section_end(rows, master_start, all_headers)
    master_map = build_label_map(rows, master_start, master_end)

    print(f"Master '{MASTER_SECTION_HEADER}': rows {master_start + 1}–{master_end}")
    print(f"  {len(master_map)} labeled row(s): {[k for k in master_map if not k.startswith('__')]}\n")

    requests = []

    for loc_name, section_header in LOCATION_SECTION_HEADERS.items():
        loc_start = find_section_start(rows, section_header)
        if loc_start is None:
            print(f"WARNING: '{section_header}' not found — skipping.")
            continue

        loc_end = find_section_end(rows, loc_start, all_headers)
        loc_map = build_label_map(rows, loc_start, loc_end)

        print(f"Section '{section_header}': rows {loc_start + 1}–{loc_end}")

        # 1. Always copy section header row formatting (text differs, so handle explicitly)
        print(f"  row {master_start + 1} (section header) → row {loc_start + 1}")
        requests.append(make_copy_request(numeric_id, master_start, loc_start))

        # 2. Match remaining rows by label
        matched = 0
        unmatched = []
        master_non_header_keys = [k for k in master_map if k != normalize(MASTER_SECTION_HEADER)]
        for key in master_non_header_keys:
            master_row = master_map[key]
            if key in loc_map:
                loc_row = loc_map[key]
                print(f"  row {master_row + 1} ({key!r}) → row {loc_row + 1}")
                requests.append(make_copy_request(numeric_id, master_row, loc_row))
                matched += 1
            else:
                unmatched.append(key)

        if unmatched:
            print(f"  NOTE: no match for: {unmatched}")
        print(f"  {matched + 1} row(s) mapped (including section header)\n")

    if not requests:
        print("Nothing to do — no matching rows found.")
        return

    print(f"{'Would apply' if dry_run else 'Applying'} {len(requests)} format copy request(s).")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests}
    ).execute()
    print("\nDone — formatting applied.")


if __name__ == "__main__":
    main()
