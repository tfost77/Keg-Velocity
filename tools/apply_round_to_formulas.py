#!/usr/bin/env python3
"""
Wrap Avg Weekly Sales, Avg Daily Sales, and # of Days Remaining formula cells
with ROUND(..., 0) in all three sections (Total Volume, LP, Timonium).

Projected Kick Date is left alone — it references Days Remaining which will
already be a whole number after rounding.

Usage:
  python apply_round_to_formulas.py              # dry run
  python apply_round_to_formulas.py --write      # apply
"""

import argparse
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


def beer_cols_from_header_row(service, sheet_id, tab, row_num):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A{row_num}:AZ{row_num}",
        valueRenderOption="FORMATTED_VALUE"
    ).execute()
    row = (result.get("values") or [[]])[0]
    return [j for j, cell in enumerate(row) if cell and j > 0]


def build_rounded_formulas(tab, beer_col_indices, rows):
    """
    rows: dict with keys avg_weekly, avg_daily, days_remaining
    and each value is a dict with keys: formula_template
    formula_template is a string with {c} placeholder for column letter.
    """
    updates = []
    for col_idx in beer_col_indices:
        c = col_letter(col_idx)
        for row_num, template in rows.items():
            formula = template.format(c=c)
            updates.append((f"'{tab}'!{c}{row_num}", formula))
    return updates


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    dry_run = not args.write

    if dry_run:
        print("DRY RUN — pass --write to apply\n")

    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    service = get_sheets_service()
    tab = find_current_tab(service, sheet_id)
    print(f"Tab: '{tab}'\n")

    # Section definitions: (label, beer_header_row, formula_rows)
    # formula_rows: {sheet_row_num: formula_template_with_{c}_placeholder}
    sections = [
        (
            "Total Volume",
            2,   # Beer: header row
            {
                11: '=ROUND(AVERAGEIF({c}6:{c}10,"<>0"),0)',
                12: '=ROUND({c}11/6,0)',
                14: '=ROUND({c}5/{c}12,0)',
            }
        ),
        (
            "DBC Locust Point",
            19,
            {
                28: '=ROUND(AVERAGEIF({c}23:{c}27,"<>0"),0)',
                29: '=ROUND({c}28/6,0)',
                31: '=ROUND({c}22/{c}29,0)',
            }
        ),
        (
            "DBC Timonium",
            37,
            {
                48: '=ROUND(AVERAGEIF({c}43:{c}47,"<>0"),0)',
                49: '=ROUND({c}48/6,0)',
                51: '=ROUND({c}42/{c}49,0)',
            }
        ),
    ]

    all_updates = []
    for label, beer_header_row, formula_rows in sections:
        cols = beer_cols_from_header_row(service, sheet_id, tab, beer_header_row)
        updates = build_rounded_formulas(tab, cols, formula_rows)
        all_updates.extend(updates)
        print(f"{label} — {len(cols)} beers × {len(formula_rows)} rows = {len(updates)} cells")

    print(f"\n{'Would write' if dry_run else 'Writing'} {len(all_updates)} formula(s):")
    for range_str, formula in all_updates:
        print(f"  {range_str}  →  {formula}")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    data = [{"range": r, "values": [[f]]} for r, f in all_updates]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data}
    ).execute()
    print("\nDone.")


if __name__ == "__main__":
    main()
