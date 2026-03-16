#!/usr/bin/env python3
"""
Ensure all formula rows in the Cans Inventory tab are filled for every
product column. Fills any missing cells and is safe to re-run (skips
cells that already have a formula).

Formula rows (discovered dynamically by row label in col A):
  - # of Units Total    → ={col}<case_volume_row>*{col}<units_per_case_row>
  - Avg Weekly Sales    → =AVERAGEIF({col}<w1>:{col}<w5>,"<>0")
  - Avg Daily Sales     → ={col}<avg_weekly_row>/6
  - # of Days Remaining → ={col}<units_total_row>/{col}<avg_daily_row>
  - Projected Kick Date → =TODAY()+{col}<days_remaining_row>

Usage:
  python apply_cans_formulas.py           # dry run (preview)
  python apply_cans_formulas.py --write   # apply changes
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
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
load_dotenv(BASE_DIR / ".env")


def load_config():
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        print("ERROR: config.json not found.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)


CONFIG = load_config()
CANS_SECTION_HEADER = CONFIG["cans_section_header"]


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


def find_cans_tab(service, sheet_id):
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    all_titles = [s["properties"]["title"] for s in meta["sheets"]]
    today = datetime.now()
    for delta in [0, -1]:
        month = today.month + delta
        year = today.year
        if month <= 0:
            month += 12
            year -= 1
        title = f"Cans Inventory - {datetime(year, month, 1).strftime('%B %Y')}"
        if title in all_titles:
            return title
    if "Cans Inventory" in all_titles:
        return "Cans Inventory"
    tabs = sorted(t for t in all_titles if t.startswith("Cans Inventory"))
    return tabs[-1] if tabs else None


def read_all_rows(service, sheet_id, tab, formulas=False):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:AZ60",
        valueRenderOption="FORMULA" if formulas else "FORMATTED_VALUE"
    ).execute()
    return result.get("values", [])


def normalize(s):
    return s.replace("\u2019", "'").replace("\u2018", "'").strip()


def find_row_by_label(rows, label):
    """Return 1-based sheet row number for the first row whose col A matches label."""
    for i, row in enumerate(rows):
        if row and normalize(row[0]) == normalize(label):
            return i + 1  # 1-based
    return None


def find_sales_week_rows(rows):
    """Return list of 1-based row numbers for Sales Week 1–5 rows."""
    week_rows = []
    for i, row in enumerate(rows):
        if row and re.match(r"Sales Week \d+", normalize(row[0])):
            week_rows.append(i + 1)
    return sorted(week_rows)


def build_formulas_for_col(c, row_nums):
    """
    Return {1-based_row: formula_string} for all formula rows for column letter c.
    row_nums is a dict with keys: case_volume, units_per_case, units_total,
    sales_weeks (list), avg_weekly, avg_daily, days_remaining, projected_kick.
    """
    w1 = row_nums["sales_weeks"][0]
    w5 = row_nums["sales_weeks"][-1]
    return {
        row_nums["units_total"]:    f"={c}{row_nums['case_volume']}*{c}{row_nums['units_per_case']}",
        row_nums["avg_weekly"]:     f'=AVERAGEIF({c}{w1}:{c}{w5},"<>0")',
        row_nums["avg_daily"]:      f"={c}{row_nums['avg_weekly']}/6",
        row_nums["days_remaining"]: f"={c}{row_nums['units_total']}/{c}{row_nums['avg_daily']}",
        row_nums["projected_kick"]: f"=TODAY()+{c}{row_nums['days_remaining']}",
    }


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
    tab = find_cans_tab(service, sheet_id)
    if not tab:
        print("ERROR: No Cans Inventory tab found.")
        sys.exit(1)
    print(f"Tab: '{tab}'\n")

    rows_display = read_all_rows(service, sheet_id, tab, formulas=False)
    rows_formula = read_all_rows(service, sheet_id, tab, formulas=True)

    # Discover row numbers by label
    row_nums = {
        "case_volume":    find_row_by_label(rows_display, "Case Volume"),
        "units_per_case": find_row_by_label(rows_display, "# of Units Per Case"),
        "units_total":    find_row_by_label(rows_display, "# of Units Total"),
        "sales_weeks":    find_sales_week_rows(rows_display),
        "avg_weekly":     find_row_by_label(rows_display, "Avg Weekly Sales"),
        "avg_daily":      find_row_by_label(rows_display, "Avg Daily Sales"),
        "days_remaining": find_row_by_label(rows_display, "# of Days Remaining"),
        "projected_kick": find_row_by_label(rows_display, "Projected Kick Date"),
    }

    missing = [k for k, v in row_nums.items() if v is None or (k == "sales_weeks" and not v)]
    if missing:
        print(f"ERROR: Could not find row(s) for: {missing}")
        sys.exit(1)

    print("Row structure discovered:")
    for k, v in row_nums.items():
        print(f"  {k}: {v}")
    print()

    # Find the product header row (section_start + 2)
    section_start = find_row_by_label(rows_display, CANS_SECTION_HEADER)
    if section_start is None:
        print(f"ERROR: Section '{CANS_SECTION_HEADER}' not found.")
        sys.exit(1)
    product_header_row_idx = section_start + 1  # 0-based index (section_start is 1-based)
    product_header = rows_formula[product_header_row_idx] if product_header_row_idx < len(rows_formula) else []
    product_cols = [j for j, cell in enumerate(product_header) if cell and j > 0 and cell not in ("Beer:", "Can:")]

    print(f"Product columns (0-based indices): {product_cols}\n")

    # For each product column, check each formula row and fill if missing
    updates = []
    for col_idx in product_cols:
        c = col_letter(col_idx)
        formulas = build_formulas_for_col(c, row_nums)

        for row_num_1based, formula in formulas.items():
            row_idx = row_num_1based - 1
            # Check if cell already has a formula
            existing = ""
            if row_idx < len(rows_formula):
                r = rows_formula[row_idx]
                existing = r[col_idx] if col_idx < len(r) else ""
            if isinstance(existing, str) and existing.startswith("="):
                continue  # already has a formula
            range_str = f"'{tab}'!{c}{row_num_1based}"
            updates.append((range_str, formula))
            print(f"  {range_str}  →  {formula}")

    if not updates:
        print("All formula cells already populated — nothing to do.")
        return

    print(f"\n{'Would write' if dry_run else 'Writing'} {len(updates)} formula(s).")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    data = [{"range": r, "values": [[f]]} for r, f in updates]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data}
    ).execute()
    print("\nDone.")


if __name__ == "__main__":
    main()
