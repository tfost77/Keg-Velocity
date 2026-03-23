#!/usr/bin/env python3
"""
Right-justify all cells (columns B onward) in rows labeled:
  - Total Keg Volume
  - In Timonium
  - In LP Cold Box
  - # of Pours per Keg
  - # of Pours Total

Applies across the entire tab (all sections).

Usage:
  python right_align_data_rows.py           # dry run (preview)
  python right_align_data_rows.py --write   # apply changes
"""

import argparse
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from auth import get_sheets_service

from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent.parent
load_dotenv(BASE_DIR / ".env")

TARGET_LABELS = {
    "Total Keg Volume",
    "Keg Volume:",
    "In Timonium",
    "In LP Cold Box",
    "# of Pours per Keg",
    "# of Pours Total",
}



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


def read_col_a(service, sheet_id, tab):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:A100",
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    return [row[0] if row else "" for row in result.get("values", [])]


def normalize(s):
    return s.replace("\u2019", "'").replace("\u2018", "'").strip()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true", help="Apply changes (default is dry run)")
    args = parser.parse_args()
    dry_run = not args.write

    if dry_run:
        print("DRY RUN — pass --write to apply\n")

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
    col_a = read_col_a(service, sheet_id, tab)

    # Find all rows whose col A label matches a target
    target_rows = []
    for i, label in enumerate(col_a):
        if normalize(label) in {normalize(t) for t in TARGET_LABELS}:
            target_rows.append((i, label))

    if not target_rows:
        print("No matching rows found.")
        return

    print("Rows to right-align:")
    for row_idx, label in target_rows:
        print(f"  Row {row_idx + 1}: {label!r}")
    print()

    # Build one repeatCell request per row: right-align columns B onward (indices 1–51)
    requests = []
    for row_idx, _ in target_rows:
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": numeric_id,
                    "startRowIndex": row_idx,
                    "endRowIndex": row_idx + 1,
                    "startColumnIndex": 1,   # col B
                    "endColumnIndex": 52,    # through AZ
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "RIGHT"
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment",
            }
        })

    print(f"{'Would apply' if dry_run else 'Applying'} {len(requests)} alignment request(s).")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests}
    ).execute()
    print("Done.")


if __name__ == "__main__":
    main()
