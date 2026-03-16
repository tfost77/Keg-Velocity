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


def load_config():
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        print("ERROR: config.json not found. Copy config.example.json to config.json and customize it.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)

CONFIG = load_config()

# Maps extracted Toast beer name → sheet column header
BEER_NAME_ALIASES = CONFIG["beer_name_aliases"]

# Sheet section header for each location
LOCATION_SECTION_HEADERS = CONFIG["location_section_headers"]

# Header text for the master Total Volume section (used for column alignment)
MASTER_SECTION_HEADER = CONFIG.get("master_section_header")

# Serve-type weights
SERVE_WEIGHTS = {
    "Pint":      1.0,
    "Mug":       1.2,
    "Half Pour": 0.5,
    "To-Go":     0.0,   # excluded; goes to Cans Inventory via sync_cans_to_sheets.py
    "Other":     0.0,   # excluded; uncategorized items not tracked in inventory
}


def slugify(name):
    return name.lower().replace(" ", "_")


def extract_beer_name(toast_name):
    """Strip number prefix (e.g. '2 -', '2M Gold -', '2H -') and normalize."""
    name = re.sub(r'^\d+[MHmh]?\s*(Gold|Silver)?\s*-\s*', '', toast_name).strip()
    name = name.replace('\u2019', "'")  # normalize curly apostrophe → straight (Toast locations differ)
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


def expand_sheet_if_needed(service, spreadsheet_id, tab, needed_col_count):
    """Expand the sheet's column count if it would be exceeded. Returns the sheet's numeric ID."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab:
            numeric_id = sheet["properties"]["sheetId"]
            current = sheet["properties"]["gridProperties"]["columnCount"]
            if needed_col_count > current:
                new_count = needed_col_count + 10  # add buffer
                service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body={"requests": [{"updateSheetProperties": {
                        "properties": {"sheetId": numeric_id, "gridProperties": {"columnCount": new_count}},
                        "fields": "gridProperties.columnCount"
                    }}]}
                ).execute()
                print(f"Expanded sheet columns to {new_count}.")
            return numeric_id
    return None


def copy_column_formatting(service, spreadsheet_id, numeric_sheet_id, source_col_idx, dest_col_indices, row_count=60):
    """Copy formatting from the Total Volume reference column to each new column.

    Uses PASTE_FORMAT so only formatting is copied — no values or formulas are touched.
    Column B (index 1) in Total Volume is the canonical format reference.
    """
    requests = [
        {
            "copyPaste": {
                "source": {
                    "sheetId": numeric_sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": row_count,
                    "startColumnIndex": source_col_idx,
                    "endColumnIndex": source_col_idx + 1,
                },
                "destination": {
                    "sheetId": numeric_sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": row_count,
                    "startColumnIndex": dest_col,
                    "endColumnIndex": dest_col + 1,
                },
                "pasteType": "PASTE_FORMAT",
                "pasteOrientation": "NORMAL",
            }
        }
        for dest_col in dest_col_indices
    ]
    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": requests}
        ).execute()


def find_all_beer_header_rows(rows):
    """Return beer-header row indices for all sections (master + each location).

    Each section's beer header row is always section_start + 2 (the 'Beer:' row).
    Sections not found in the current sheet are silently skipped.
    """
    all_section_texts = []
    if MASTER_SECTION_HEADER:
        all_section_texts.append(MASTER_SECTION_HEADER)
    all_section_texts.extend(LOCATION_SECTION_HEADERS.values())

    header_rows = []
    for header_text in all_section_texts:
        section_start = find_section_start(rows, header_text)
        if section_start is not None:
            header_rows.append(section_start + 2)
    return header_rows


def copy_master_formula_rows(service, spreadsheet_id, numeric_sheet_id, tab, new_col_indices):
    """Copy all formula rows in the master (Total Volume) section to new columns.

    Reads col B formulas for the master section and copies any cell whose value starts
    with '=' to the new columns using PASTE_FORMULA (so relative references adjust
    automatically). This covers Keg Volume, # of Pours Total, all Sales Week aggregates,
    Avg Weekly/Daily Sales, # of Days Remaining, and Projected Kick Date.

    NOTE: '# of Pours per Keg' is a plain value (not a formula) and is intentionally
    skipped — keg size varies per beer and must be filled in by hand.
    """
    if not MASTER_SECTION_HEADER:
        return

    # Read col A+B with FORMULA rendering to detect which rows have formula cells
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab}'!A1:B60",
        valueRenderOption="FORMULA"
    ).execute()
    formula_rows = result.get("values", [])

    master_start = next(
        (i for i, r in enumerate(formula_rows) if r and r[0] == MASTER_SECTION_HEADER),
        None
    )
    if master_start is None:
        return

    # Collect rows in the section (up to 20 rows) where col B is a formula
    formula_row_indices = [
        i for i in range(master_start, min(master_start + 20, len(formula_rows)))
        if len(formula_rows[i]) > 1 and isinstance(formula_rows[i][1], str)
        and formula_rows[i][1].startswith("=")
    ]
    if not formula_row_indices:
        return

    requests = [
        {
            "copyPaste": {
                "source": {
                    "sheetId": numeric_sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": 1,  # col B as template
                    "endColumnIndex": 2,
                },
                "destination": {
                    "sheetId": numeric_sheet_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": dest_col,
                    "endColumnIndex": dest_col + 1,
                },
                "pasteType": "PASTE_FORMULA",
                "pasteOrientation": "NORMAL",
            }
        }
        for row_idx in formula_row_indices
        for dest_col in new_col_indices
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()
    print(f"Copied master section formulas for {len(formula_row_indices)} row(s) to {len(new_col_indices)} new column(s)")


def add_new_columns(service, sheet_id, tab, header_row_idx, existing_cols, new_names, all_header_row_indices=None):
    """Append new column headers to ALL section header rows and return the updated col dict.

    Writing to every section (master + all locations) at the same column indices keeps the
    sheet aligned so that beer columns line up vertically across all sections.
    Formatting is copied from column B (Total Volume reference) so new columns are uniform.
    All formula rows in the master section are extended automatically to the new columns.
    """
    if not new_names:
        return existing_cols
    next_col = max(existing_cols.values()) + 1
    updated_cols = dict(existing_cols)
    for name in new_names:
        updated_cols[name] = next_col
        next_col += 1
    start_col = max(existing_cols.values()) + 1
    new_col_indices = list(range(start_col, start_col + len(new_names)))
    numeric_sheet_id = expand_sheet_if_needed(service, sheet_id, tab, start_col + len(new_names))

    # Copy formatting from column B (Total Volume reference) to all new columns first,
    # so borders, background colors, and alignment are uniform before writing headers.
    if numeric_sheet_id is not None:
        copy_column_formatting(service, sheet_id, numeric_sheet_id,
                               source_col_idx=1, dest_col_indices=new_col_indices)

    # Write header to every section so all stay aligned; fall back to current section only.
    target_rows = all_header_row_indices if all_header_row_indices else [header_row_idx]
    updates = []
    for hrow_idx in target_rows:
        if hrow_idx is not None:
            r = f"'{tab}'!{col_letter(start_col)}{hrow_idx + 1}:{col_letter(start_col + len(new_names) - 1)}{hrow_idx + 1}"
            updates.append({"range": r, "values": [new_names]})
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "RAW", "data": updates}
    ).execute()

    # Extend all master-section formula rows to new columns (Keg Volume, # of Pours Total,
    # Sales Week aggregates, Avg Weekly/Daily Sales, # of Days Remaining, Projected Kick Date).
    # Uses PASTE_FORMULA from col B so relative references adjust automatically.
    if numeric_sheet_id is not None:
        copy_master_formula_rows(service, sheet_id, numeric_sheet_id, tab, new_col_indices)

    print(f"Added new column(s): {', '.join(new_names)} (synced across {len(target_rows)} section(s), formatting copied from col B)")
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
    parser.add_argument("--week", help='Target a specific week label, e.g. "Sales Week 4"')
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
    if args.week:
        target_row_idx = next(
            (i for i in range(section_start, min(section_start + 20, len(rows)))
             if rows[i] and rows[i][0] == args.week),
            None
        )
        if target_row_idx is None:
            print(f"ERROR: '{args.week}' not found in section '{args.location}'")
            sys.exit(1)
        mode = "Overwriting"
    elif args.overwrite:
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

    # Add columns for any beers in Toast data not yet in the sheet.
    # Pass all section beer-header rows so every section stays column-aligned.
    new_beers = [b for b in beer_totals if b not in beer_cols]
    if new_beers:
        all_header_rows = find_all_beer_header_rows(rows)
        beer_cols = add_new_columns(service, sheet_id, tab, beer_header_idx, beer_cols, new_beers, all_header_rows)

    max_col = max(beer_cols.values()) + 1
    current_row = list(rows[target_row_idx]) if target_row_idx < len(rows) else []
    new_row = current_row + [""] * (max_col - len(current_row))

    matched, no_data = [], []
    for beer_name, col_idx in beer_cols.items():
        if beer_name in beer_totals:
            new_row[col_idx] = int(beer_totals[beer_name])
            matched.append(f"{beer_name}: {beer_totals[beer_name]}")
        elif col_idx < len(current_row) and isinstance(current_row[col_idx], str):
            # current_row values come back from the API as strings; keep them as numbers
            # so RAW doesn't store them as text (which breaks AVERAGEIF formulas)
            raw = current_row[col_idx].strip()
            try:
                new_row[col_idx] = int(float(raw)) if raw else ""
            except (ValueError, TypeError):
                new_row[col_idx] = raw
            no_data.append(beer_name)
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
