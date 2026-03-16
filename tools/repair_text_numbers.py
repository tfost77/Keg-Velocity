#!/usr/bin/env python3
"""
Repair: convert text-stored numbers anywhere in the Total Inventory tab back to actual numbers.

The apostrophe prefix (') appears in Google Sheets when a number was stored as text
via valueInputOption="RAW" with a string value. This breaks AVERAGEIF and other formulas.

Scans ALL rows in the tab and rewrites numeric strings as actual numbers using USER_ENTERED.
Formula cells (values starting with '=') are always skipped.

Usage:
  python repair_text_numbers.py              # dry run (preview only)
  python repair_text_numbers.py --write      # apply changes
"""

import argparse
import json
import os
import re
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


def load_config():
    config_path = BASE_DIR / "config.json"
    with open(config_path) as f:
        return json.load(f)


def col_letter(n):
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord("A")) + result
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


def read_all_rows(service, sheet_id, tab, render="FORMATTED_VALUE"):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:AZ100",
        valueRenderOption=render
    ).execute()
    return result.get("values", [])


def is_formula(val):
    return isinstance(val, str) and val.startswith("=")


def is_numeric_string(val):
    """Return True if val is a non-empty string that represents an integer or float."""
    if not isinstance(val, str) or not val.strip():
        return False
    return bool(re.match(r'^\d+(\.\d+)?$', val.strip()))


def to_number(val):
    """Convert a numeric string to int or float."""
    v = val.strip()
    return int(v) if '.' not in v else float(v)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true", help="Apply changes (default is dry run)")
    args = parser.parse_args()
    dry_run = not args.write

    if dry_run:
        print("DRY RUN — pass --write to apply changes\n")

    CONFIG = load_config()
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

    rows_display = read_all_rows(service, sheet_id, tab, render="FORMATTED_VALUE")
    rows_formula = read_all_rows(service, sheet_id, tab, render="FORMULA")

    updates = []

    for row_idx, row in enumerate(rows_display):
        formula_row = rows_formula[row_idx] if row_idx < len(rows_formula) else []
        col_a = row[0] if row else ""

        # Scan all columns (skip col A which is the row label)
        for col_idx in range(1, len(row)):
            # Skip cells that already contain a formula
            formula_val = formula_row[col_idx] if col_idx < len(formula_row) else ""
            if is_formula(formula_val):
                continue

            val = row[col_idx]
            if is_numeric_string(val):
                num = to_number(val)
                range_str = f"'{tab}'!{col_letter(col_idx)}{row_idx + 1}"
                updates.append((range_str, num, col_a, val))

    if not updates:
        print("No text-stored numbers found in Sales Week rows.")
        return

    print(f"Found {len(updates)} text-stored number(s) to fix:")
    for range_str, num, row_label, original in updates:
        print(f"  {range_str}  '{original}' → {num}  (in {row_label})")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    # Write back as USER_ENTERED so Sheets parses strings as numbers
    data = [{"range": r, "values": [[n]]} for r, n, _, _ in updates]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data}
    ).execute()
    print("\nDone — all text-stored numbers converted to actual numbers.")


if __name__ == "__main__":
    main()
