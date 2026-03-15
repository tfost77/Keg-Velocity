#!/usr/bin/env python3
"""
Sync weekly draft beer sales to the Total Inventory tab in Google Sheets.

Weighting rules:
  - Pint:      1.0 per unit
  - Mug:       1.2 per unit
  - Half Pour: 0.5 per unit
  - To-Go:     excluded (handled by sync_cans_to_sheets.py)
  - Other:     1.0 per unit

Usage:
  python sync_to_sheets.py --location "Locust Point"            # append to next empty week
  python sync_to_sheets.py --location "Locust Point" --overwrite  # replace last filled week

Input:  .tmp/{location_slug}_parsed_items.json
Output: Writes to Google Sheet (GOOGLE_SHEET_ID in .env)
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
]

load_dotenv(BASE_DIR / ".env")

# Maps extracted Toast beer name → sheet column header
BEER_NAME_ALIASES = {
    "The Bamb": "Bamb",
    "Cascara Saison": "Cascara",
    "3:30 Amber": "3:30 Amber Ale",
}

# Sheet section header for each location
LOCATION_SECTION_HEADERS = {
    "Locust Point": "DBC Locust Point",
    "Timonium":     "DBC Timonium",
}

# Serve-type weights
SERVE_WEIGHTS = {
    "Pint":      1.0,
    "Mug":       1.2,
    "Half Pour": 0.5,
    "To-Go":     0.0,   # excluded; goes to Cans Inventory via sync_cans_to_sheets.py
    "Other":     1.0,
}


def slugify(name):
    return name.lower().replace(" ", "_")


def extract_beer_name(toast_name):
    """Strip number prefix (e.g. '2 -', '2M Gold -', '2H -') and normalize."""
    name = re.sub(r'^\d+[MHmh]?\s*(Gold|Silver)?\s*-\s*', '', toast_name).strip()
    return BEER_NAME_ALIASES.get(name, name)


def col_letter(n):
    """Convert 0-based column index to spreadsheet letter (0→A, 25→Z, 26→AA)."""
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord('A')) + result
    return result


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
            if not creds_path.exists():
                print("ERROR: credentials.json not found")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)


def find_current_tab(service, sheet_id):
    """Find the Total Inventory tab for the current month."""
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    all_titles = [s["properties"]["title"] for s in meta["sheets"]]

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

    inventory_tabs = sorted([t for t in all_titles if t.startswith("Total Inventory -")])
    return inventory_tabs[-1] if inventory_tabs else None


def read_all_rows(service, sheet_id, tab):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:Z60"
    ).execute()
    return result.get("values", [])


def find_section_start(rows, section_header):
    for i, row in enumerate(rows):
        if row and row[0] == section_header:
            return i
    return None


def find_beer_columns(rows, beer_row_idx):
    if beer_row_idx >= len(rows):
        return {}
    header_row = rows[beer_row_idx]
    return {
        cell: j
        for j, cell in enumerate(header_row)
        if cell and cell != "Beer:" and j > 0
    }


def add_new_columns(service, sheet_id, tab, header_row_idx, existing_cols, new_names):
    """Append new column headers to the product header row and return the updated col dict."""
    if not new_names:
        return existing_cols
    next_col = max(existing_cols.values()) + 1
    updated_cols = dict(existing_cols)
    for name in new_names:
        updated_cols[name] = next_col
        next_col += 1
    start_col = max(existing_cols.values()) + 1
    range_notation = f"'{tab}'!{col_letter(start_col)}{header_row_idx + 1}:{col_letter(start_col + len(new_names) - 1)}{header_row_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_notation,
        valueInputOption="RAW",
        body={"values": [new_names]}
    ).execute()
    print(f"Added new column(s): {', '.join(new_names)}")
    return updated_cols


def find_all_week_rows(rows, section_start):
    week_rows = []
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if row and row[0].startswith("Sales Week"):
            week_rows.append(i)
    return week_rows


def clear_week_rows(service, sheet_id, tab, rows, week_row_indices, beer_cols):
    max_col = max(beer_cols.values()) + 1
    data = []
    for row_idx in week_row_indices:
        current_row = list(rows[row_idx]) if row_idx < len(rows) else []
        cleared_row = current_row + [""] * (max_col - len(current_row))
        for col_idx in beer_cols.values():
            if col_idx < len(cleared_row):
                cleared_row[col_idx] = ""
        data.append({
            "range": f"'{tab}'!A{row_idx + 1}:{col_letter(max_col - 1)}{row_idx + 1}",
            "values": [cleared_row[:max_col]]
        })
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "RAW", "data": data}
    ).execute()
    print(f"Cycle reset: cleared {len(week_row_indices)} Sales Week rows.")


def find_next_empty_week_row(rows, section_start, beer_col_indices):
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if not row or not row[0].startswith("Sales Week"):
            continue
        empty = all(
            (col_idx >= len(row) or not row[col_idx] or row[col_idx] == "0")
            for col_idx in beer_col_indices
        )
        if empty:
            return i
    return None


def find_last_filled_week_row(rows, section_start, beer_col_indices):
    last = None
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if not row or not row[0].startswith("Sales Week"):
            continue
        has_data = any(
            col_idx < len(row) and row[col_idx] and row[col_idx] != "0"
            for col_idx in beer_col_indices
        )
        if has_data:
            last = i
    return last


def classify_serve_type(name):
    if re.search(r'pack', name, re.IGNORECASE):
        return "To-Go"
    if re.match(r'^\d+M[\s-]', name):
        return "Mug"
    if re.match(r'^\d+H[\s-]', name):
        return "Half Pour"
    if re.match(r'^\d+ -', name):
        return "Pint"
    return "Other"


def aggregate_by_beer(items):
    """Sum weighted draft units per beer. To-Go excluded."""
    totals = {}
    for item in items:
        serve_type = classify_serve_type(item["name"])
        weight = SERVE_WEIGHTS[serve_type]
        if weight == 0.0:
            continue
        beer = extract_beer_name(item["name"])
        totals[beer] = totals.get(beer, 0) + (item["units_sold"] * weight)
    return {beer: round(total) for beer, total in totals.items()}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--location", required=True, help='Location name, e.g. "Locust Point"')
    parser.add_argument("--overwrite", action="store_true", help="Overwrite the last filled Sales Week row instead of appending")
    args = parser.parse_args()

    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("ERROR: GOOGLE_SHEET_ID not set in .env")
        sys.exit(1)

    slug = slugify(args.location)
    in_path = TMP_DIR / f"{slug}_parsed_items.json"
    if not in_path.exists():
        print(f"ERROR: {in_path.name} not found. Run parse_toast_csv.py first.")
        sys.exit(1)

    with open(in_path) as f:
        items = json.load(f)

    beer_totals = aggregate_by_beer(items)
    service = get_sheets_service()

    tab = find_current_tab(service, sheet_id)
    if not tab:
        print("ERROR: No 'Total Inventory' tab found for current month.")
        sys.exit(1)
    print(f"Tab: '{tab}'")

    rows = read_all_rows(service, sheet_id, tab)

    section_header = LOCATION_SECTION_HEADERS.get(args.location)
    if not section_header:
        print(f"ERROR: No section header configured for '{args.location}'")
        sys.exit(1)

    section_start = find_section_start(rows, section_header)
    if section_start is None:
        print(f"ERROR: Section '{section_header}' not found in tab '{tab}'")
        sys.exit(1)

    beer_header_idx = section_start + 2
    beer_cols = find_beer_columns(rows, beer_header_idx)
    if not beer_cols:
        print(f"ERROR: No beer columns found under '{section_header}'")
        sys.exit(1)

    col_indices = list(beer_cols.values())
    if args.overwrite:
        target_row_idx = find_last_filled_week_row(rows, section_start, col_indices)
        if target_row_idx is None:
            print(f"ERROR: No filled Sales Week rows found to overwrite for '{args.location}'")
            sys.exit(1)
        mode = "Overwriting"
    else:
        target_row_idx = find_next_empty_week_row(rows, section_start, col_indices)
        if target_row_idx is None:
            print(f"All Sales Weeks filled for '{args.location}' — resetting cycle.")
            all_week_rows = find_all_week_rows(rows, section_start)
            clear_week_rows(service, sheet_id, tab, rows, all_week_rows, beer_cols)
            rows = read_all_rows(service, sheet_id, tab)
            target_row_idx = find_next_empty_week_row(rows, section_start, col_indices)
            if target_row_idx is None:
                print(f"ERROR: Could not find empty week row after cycle reset for '{args.location}'")
                sys.exit(1)
        mode = "Writing to"

    week_label = rows[target_row_idx][0] if target_row_idx < len(rows) and rows[target_row_idx] else "?"
    print(f"{mode}: {week_label} (sheet row {target_row_idx + 1})")

    # Add columns for any beers in Toast data not yet in the sheet
    new_beers = [b for b in beer_totals if b not in beer_cols]
    if new_beers:
        beer_cols = add_new_columns(service, sheet_id, tab, beer_header_idx, beer_cols, new_beers)

    max_col = max(beer_cols.values()) + 1
    current_row = list(rows[target_row_idx]) if target_row_idx < len(rows) else []
    new_row = current_row + [""] * (max_col - len(current_row))

    matched, no_data = [], []
    for beer_name, col_idx in beer_cols.items():
        if beer_name in beer_totals:
            new_row[col_idx] = beer_totals[beer_name]
            matched.append(f"{beer_name}: {beer_totals[beer_name]}")
        else:
            no_data.append(beer_name)

    range_notation = f"'{tab}'!A{target_row_idx + 1}:{col_letter(max_col - 1)}{target_row_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_notation,
        valueInputOption="RAW",
        body={"values": [new_row[:max_col]]}
    ).execute()

    print(f"Written: {', '.join(matched)}")
    if no_data:
        print(f"No Toast data for: {', '.join(no_data)}")


if __name__ == "__main__":
    main()
