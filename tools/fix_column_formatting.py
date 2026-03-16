#!/usr/bin/env python3
"""
Fix formatting for beer columns that are missing borders or alignment
on specific rows.

Detects which beer columns lack userEnteredFormat borders on the Beer: header
rows, then applies SOLID borders + correct alignment to bring them in line
with the existing well-formatted columns.

Usage:
  python fix_column_formatting.py              # dry run
  python fix_column_formatting.py --write      # apply
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

# Beer: header rows for each section (1-based sheet row numbers)
BEER_HEADER_ROWS = [2, 19, 37]


def load_config():
    config_path = BASE_DIR / "config.json"
    with open(config_path) as f:
        return json.load(f)


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


def col_letter(n):
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord("A")) + result
    return result


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


def get_sheet_meta(service, sheet_id, tab):
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab:
            return sheet["properties"]["sheetId"]
    return None


def beer_cols_from_row(service, sheet_id, tab, row_num):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A{row_num}:AZ{row_num}",
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    row = (result.get("values") or [[]])[0]
    return [j for j, cell in enumerate(row) if cell and j > 0]


def get_missing_border_cols(service, sheet_id, tab, beer_header_rows, all_beer_cols):
    """For each Beer: header row, return which columns are missing userEnteredFormat borders."""
    ranges = [
        f"'{tab}'!B{r}:{col_letter(max(all_beer_cols))}{r}"
        for r in beer_header_rows
    ]
    meta = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=ranges,
        includeGridData=True
    ).execute()

    # Map: row_num → set of col indices missing borders
    missing = {}
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab:
            for range_idx, range_data in enumerate(sheet.get("data", [])):
                row_num = beer_header_rows[range_idx]
                cells = (range_data.get("rowData") or [{}])[0].get("values", [])
                for cell_idx, cell in enumerate(cells):
                    col_idx = 1 + cell_idx  # col B = index 1
                    uf = cell.get("userEnteredFormat", {})
                    has_borders = bool(uf.get("borders"))
                    if not has_borders and col_idx in all_beer_cols:
                        missing.setdefault(row_num, set()).add(col_idx)
    return missing


def solid_border():
    return {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    dry_run = not args.write

    if dry_run:
        print("DRY RUN — pass --write to apply\n")

    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    if not sheet_id:
        print("ERROR: GOOGLE_SHEET_ID not set")
        sys.exit(1)

    service = get_sheets_service()
    tab = find_current_tab(service, sheet_id)
    print(f"Tab: '{tab}'\n")

    numeric_id = get_sheet_meta(service, sheet_id, tab)
    all_beer_cols = beer_cols_from_row(service, sheet_id, tab, BEER_HEADER_ROWS[0])

    missing = get_missing_border_cols(service, sheet_id, tab, BEER_HEADER_ROWS, all_beer_cols)

    if not missing:
        print("All Beer: header rows have uniform borders — nothing to fix.")
        return

    requests = []
    for row_num, cols in sorted(missing.items()):
        for col_idx in sorted(cols):
            print(f"  Row {row_num} ({col_letter(col_idx)}): adding SOLID borders + CENTER BOLD")
            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": numeric_id,
                        "startRowIndex": row_num - 1,
                        "endRowIndex": row_num,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "textFormat": {"bold": True},
                            "borders": {
                                "top":    solid_border(),
                                "bottom": solid_border(),
                                "left":   solid_border(),
                                "right":  solid_border(),
                            },
                        }
                    },
                    "fields": "userEnteredFormat(horizontalAlignment,textFormat.bold,borders)",
                }
            })

    print(f"\n{'Would apply' if dry_run else 'Applying'} {len(requests)} fix(es).")

    if dry_run:
        print("Run with --write to apply.")
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests}
    ).execute()
    print("Done.")


if __name__ == "__main__":
    main()
