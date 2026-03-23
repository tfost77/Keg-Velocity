#!/usr/bin/env python3
"""
Sync combined To-Go (cans) data from all locations to the Cans Inventory tab.

Reads all available {slug}_parsed_items.json files from .tmp/ and merges them
into a single set of can totals before writing to the sheet.

Weighting rules:
  - 12-pack:  2.0 per unit
  - Other:    1.0 per unit

Usage:
  python sync_cans_to_sheets.py            # append to next empty week
  python sync_cans_to_sheets.py --overwrite  # replace last filled week

Input:  .tmp/*_parsed_items.json (all available locations)
Output: Writes to Cans Inventory tab in Google Sheet (GOOGLE_SHEET_ID in .env)
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from auth import get_sheets_service
from config_loader import load_config

from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"

load_dotenv(BASE_DIR / ".env")

CONFIG = load_config()

# Section header in the Cans Inventory tab
CANS_SECTION_HEADER = CONFIG["cans_section_header"]

# Weights for can types
CAN_WEIGHTS = {
    "12-pack": 2.0,
    "To-Go":   1.0,
}

# Maps stripped can name → sheet column header
CAN_NAME_ALIASES = CONFIG["can_name_aliases"]


def extract_can_name(toast_name):
    """Strip pack suffix (e.g. '6-Pack', '12 Pack', '6 Pack') and normalize."""
    name = re.sub(r'\s+\d+.?[Pp]ack$', '', toast_name).strip()
    name = name.replace('\u2019', "'")  # normalize curly apostrophe → straight (Toast locations differ)
    return CAN_NAME_ALIASES.get(name, name)


def col_letter(n):
    """Convert 0-based column index to spreadsheet letter."""
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord('A')) + result
    return result



def find_cans_tab(service, sheet_id):
    """Find the Cans Inventory tab (month-based or static)."""
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
        title = f"Cans Inventory - {month_name}"
        if title in all_titles:
            return title

    if "Cans Inventory" in all_titles:
        return "Cans Inventory"

    cans_tabs = sorted([t for t in all_titles if t.startswith("Cans Inventory")])
    return cans_tabs[-1] if cans_tabs else None


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


def find_product_columns(rows, header_row_idx):
    if header_row_idx >= len(rows):
        return {}
    header_row = rows[header_row_idx]
    return {
        cell: j
        for j, cell in enumerate(header_row)
        if cell and j > 0 and cell not in ("Beer:", "Can:")
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
    """Copy formatting from the reference column (col B) to each new column."""
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


def add_new_columns(service, sheet_id, tab, header_row_idx, existing_cols, new_names):
    """Append new column headers to the product header row and return the updated col dict.

    Formatting is copied from column B (the reference column) so new columns are uniform.
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

    # Copy formatting from column B to all new columns before writing headers
    if numeric_sheet_id is not None:
        copy_column_formatting(service, sheet_id, numeric_sheet_id,
                               source_col_idx=1, dest_col_indices=new_col_indices)

    range_notation = f"'{tab}'!{col_letter(start_col)}{header_row_idx + 1}:{col_letter(start_col + len(new_names) - 1)}{header_row_idx + 1}"
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=range_notation,
        valueInputOption="RAW",
        body={"values": [new_names]}
    ).execute()
    print(f"Added new column(s): {', '.join(new_names)} (formatting copied from col B)")
    return updated_cols


def find_all_week_rows(rows, section_start):
    week_rows = []
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if row and row[0].startswith("Sales Week"):
            week_rows.append(i)
    return week_rows


def clear_week_rows(service, sheet_id, tab, rows, week_row_indices, product_cols):
    max_col = max(product_cols.values()) + 1
    data = []
    for row_idx in week_row_indices:
        current_row = list(rows[row_idx]) if row_idx < len(rows) else []
        cleared_row = current_row + [""] * (max_col - len(current_row))
        for col_idx in product_cols.values():
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


def find_next_empty_week_row(rows, section_start, col_indices):
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if not row or not row[0].startswith("Sales Week"):
            continue
        empty = all(
            (col_idx >= len(row) or not row[col_idx] or row[col_idx] == "0")
            for col_idx in col_indices
        )
        if empty:
            return i
    return None


def find_last_filled_week_row(rows, section_start, col_indices):
    last = None
    for i in range(section_start, min(section_start + 20, len(rows))):
        row = rows[i]
        if not row or not row[0].startswith("Sales Week"):
            continue
        has_data = any(
            col_idx < len(row) and row[col_idx] and row[col_idx] != "0"
            for col_idx in col_indices
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


def classify_can_type(name):
    if re.search(r'12.?pack', name, re.IGNORECASE):
        return "12-pack"
    return "To-Go"


def aggregate_cans_from_file(path):
    """Load a parsed_items.json and return {can_name: weighted_total} for To-Go items."""
    with open(path) as f:
        items = json.load(f)
    totals = {}
    for item in items:
        if classify_serve_type(item["name"]) != "To-Go":
            continue
        can_type = classify_can_type(item["name"])
        weight = CAN_WEIGHTS[can_type]
        can_name = extract_can_name(item["name"])
        totals[can_name] = totals.get(can_name, 0) + (item["units_sold"] * weight)
    return totals


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true", help="Overwrite the last filled Sales Week row instead of appending")
    args = parser.parse_args()

    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("ERROR: GOOGLE_SHEET_ID not set in .env")
        sys.exit(1)

    # Combine To-Go totals from all available location files
    parsed_files = list(TMP_DIR.glob("*_parsed_items.json"))
    if not parsed_files:
        print("ERROR: No parsed_items.json files found in .tmp/")
        sys.exit(1)

    combined_totals = {}
    for path in parsed_files:
        location_totals = aggregate_cans_from_file(path)
        for can_name, qty in location_totals.items():
            combined_totals[can_name] = combined_totals.get(can_name, 0) + qty

    # Round after combining
    combined_totals = {can: round(qty) for can, qty in combined_totals.items()}

    sources = ", ".join(p.stem.replace("_parsed_items", "") for p in parsed_files)
    print(f"Combined can totals from: {sources}")

    service = get_sheets_service()

    tab = find_cans_tab(service, sheet_id)
    if not tab:
        print("ERROR: No 'Cans Inventory' tab found in sheet.")
        sys.exit(1)
    print(f"Tab: '{tab}'")

    rows = read_all_rows(service, sheet_id, tab)

    section_start = find_section_start(rows, CANS_SECTION_HEADER)
    if section_start is None:
        print(f"ERROR: Section '{CANS_SECTION_HEADER}' not found in tab '{tab}'")
        sys.exit(1)

    # Product header is 2 rows below section header
    product_cols = find_product_columns(rows, section_start + 2)
    if not product_cols:
        print(f"ERROR: No product columns found under '{CANS_SECTION_HEADER}'")
        sys.exit(1)

    col_indices = list(product_cols.values())

    if args.overwrite:
        target_row_idx = find_last_filled_week_row(rows, section_start, col_indices)
        if target_row_idx is None:
            print("ERROR: No filled Sales Week rows found to overwrite.")
            sys.exit(1)
        mode = "Overwriting"
    else:
        target_row_idx = find_next_empty_week_row(rows, section_start, col_indices)
        if target_row_idx is None:
            print("All Sales Weeks filled — resetting cycle.")
            all_week_rows = find_all_week_rows(rows, section_start)
            clear_week_rows(service, sheet_id, tab, rows, all_week_rows, product_cols)
            rows = read_all_rows(service, sheet_id, tab)
            target_row_idx = find_next_empty_week_row(rows, section_start, col_indices)
            if target_row_idx is None:
                print("ERROR: Could not find empty week row after cycle reset.")
                sys.exit(1)
        mode = "Writing to"

    week_label = rows[target_row_idx][0] if target_row_idx < len(rows) and rows[target_row_idx] else "?"
    print(f"{mode}: {week_label} (sheet row {target_row_idx + 1})")

    # Add columns for any cans in Toast data not yet in the sheet
    product_header_idx = section_start + 2
    new_cans = [c for c in combined_totals if c not in product_cols]
    if new_cans:
        product_cols = add_new_columns(service, sheet_id, tab, product_header_idx, product_cols, new_cans)

    max_col = max(product_cols.values()) + 1
    current_row = list(rows[target_row_idx]) if target_row_idx < len(rows) else []
    new_row = current_row + [""] * (max_col - len(current_row))

    matched, no_data = [], []
    for col_name, col_idx in product_cols.items():
        if col_name in combined_totals:
            new_row[col_idx] = combined_totals[col_name]
            matched.append(f"{col_name}: {combined_totals[col_name]}")
        else:
            no_data.append(col_name)

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
