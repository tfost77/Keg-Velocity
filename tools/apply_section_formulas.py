#!/usr/bin/env python3
"""
Apply Avg Weekly Sales, Avg Daily Sales, # of Days Remaining, and Projected Kick
formulas to the DBC Locust Point and DBC Timonium sections.

Mirrors the same formula pattern already used in the Total Volume section.
Also fills any missing columns in existing formula rows.

Usage:
  python apply_section_formulas.py              # dry run (preview only)
  python apply_section_formulas.py --write      # apply changes
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

TAB = None  # resolved at runtime


def col_letter(n):
    """0-based column index → spreadsheet letter."""
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


def read_row(service, sheet_id, tab, row_num, render="FORMULA"):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A{row_num}:AZ{row_num}",
        valueRenderOption=render
    ).execute()
    rows = result.get("values", [])
    return rows[0] if rows else []


def beer_columns_from_header(row):
    """Return list of 0-based column indices for beers (skip col A)."""
    return [j for j, cell in enumerate(row) if cell and j > 0]


def build_formula_updates(tab, beer_cols, section_rows):
    """
    Build the formula updates for a section.

    section_rows: dict with keys:
      - sales_weeks: [row_num, ...] (e.g. [23,24,25,26,27])
      - keg_vol: row_num (Keg Volume / Total Keg Volume row)
      - pours_per_keg: row_num (# of Pours per Keg row)
      - pours_total: row_num (# of Pours Total row)
      - avg_weekly: row_num
      - avg_daily: row_num
      - days_remaining: row_num
      - projected_kick: row_num

    Returns list of (range_str, formula_str) tuples.
    """
    updates = []

    w1 = section_rows["sales_weeks"][0]
    w5 = section_rows["sales_weeks"][-1]
    keg_row = section_rows["keg_vol"]
    ppk_row = section_rows["pours_per_keg"]
    pours_row = section_rows["pours_total"]
    aws_row = section_rows["avg_weekly"]
    ads_row = section_rows["avg_daily"]
    drem_row = section_rows["days_remaining"]
    proj_row = section_rows["projected_kick"]

    for col_idx in beer_cols:
        c = col_letter(col_idx)

        updates.append((
            f"'{tab}'!{c}{pours_row}",
            f"={c}{keg_row}*{c}{ppk_row}"
        ))
        updates.append((
            f"'{tab}'!{c}{aws_row}",
            f'=ROUND(AVERAGEIF({c}{w1}:{c}{w5},"<>0"),0)'
        ))
        updates.append((
            f"'{tab}'!{c}{ads_row}",
            f"=ROUND({c}{aws_row}/6,0)"
        ))
        updates.append((
            f"'{tab}'!{c}{drem_row}",
            f"=ROUND({c}{pours_row}/{c}{ads_row},0)"
        ))
        updates.append((
            f"'{tab}'!{c}{proj_row}",
            f"=TODAY()+{c}{drem_row}"
        ))

    return updates


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

    def cols_needing_formula(all_cols, formula_row_data, pours_total_row_data):
        """Return cols missing a formula in avg_weekly OR having a hardcoded value in # of Pours Total."""
        aws_formula_cols = set(
            j for j, cell in enumerate(formula_row_data)
            if isinstance(cell, str) and cell.startswith("=")
        )
        pours_hardcoded_cols = set(
            j for j, cell in enumerate(pours_total_row_data)
            if j > 0 and not (isinstance(cell, str) and cell.startswith("=")) and cell != ""
        )
        # Combine: fix any col missing avg_weekly formula OR having hardcoded pours_total
        return [c for c in all_cols if c not in aws_formula_cols or c in pours_hardcoded_cols]

    # ── DBC Locust Point ──────────────────────────────────────────────────────
    # Beer header: row 19
    # Keg Volume: row 20, # of Pours per Keg: row 21, # of Pours Total: row 22
    # Sales Weeks: 23-27, Avg Weekly: 28, Avg Daily: 29, Days Rem: 31, Proj Kick: 33

    lp_beer_header = read_row(service, sheet_id, tab, 19)
    lp_all_cols = beer_columns_from_header(lp_beer_header)
    lp_aws_row_data = read_row(service, sheet_id, tab, 28)
    lp_pours_row_data = read_row(service, sheet_id, tab, 22)
    lp_fix_cols = cols_needing_formula(lp_all_cols, lp_aws_row_data, lp_pours_row_data)

    lp_section = {
        "keg_vol":       20,
        "pours_per_keg": 21,
        "sales_weeks":   [23, 24, 25, 26, 27],
        "pours_total":   22,
        "avg_weekly":    28,
        "avg_daily":     29,
        "days_remaining":31,
        "projected_kick":33,
    }

    lp_updates = []
    if lp_fix_cols:
        lp_updates = build_formula_updates(tab, lp_fix_cols, lp_section)
        print(f"DBC Locust Point — fixing formulas for cols: {', '.join(col_letter(c) for c in lp_fix_cols)}")
    else:
        print("DBC Locust Point — all formula columns already present ✓")

    # ── DBC Timonium ─────────────────────────────────────────────────────────
    # Beer header: row 37
    # Total Keg Volume: row 38, # of Pours per Keg: row 41, # of Pours Total: row 42
    # Sales Weeks: 43-47, Avg Weekly: 48, Avg Daily: 49, Days Rem: 51, Proj Kick: 53

    tim_beer_header = read_row(service, sheet_id, tab, 37)
    tim_all_cols = beer_columns_from_header(tim_beer_header)
    tim_aws_row_data = read_row(service, sheet_id, tab, 48)
    tim_pours_row_data = read_row(service, sheet_id, tab, 42)
    tim_fix_cols = cols_needing_formula(tim_all_cols, tim_aws_row_data, tim_pours_row_data)

    tim_section = {
        "keg_vol":       38,
        "pours_per_keg": 41,
        "sales_weeks":   [43, 44, 45, 46, 47],
        "pours_total":   42,
        "avg_weekly":    48,
        "avg_daily":     49,
        "days_remaining":51,
        "projected_kick":53,
    }

    tim_updates = []
    if tim_fix_cols:
        tim_updates = build_formula_updates(tab, tim_fix_cols, tim_section)
        print(f"DBC Timonium — fixing formulas for cols: {', '.join(col_letter(c) for c in tim_fix_cols)}")
    else:
        print("DBC Timonium — all formula columns already present ✓")

    all_updates = lp_updates + tim_updates

    if not all_updates:
        print("\nNothing to update.")
        return

    print(f"\n{'Would write' if dry_run else 'Writing'} {len(all_updates)} formula(s):")
    for range_str, formula in all_updates:
        print(f"  {range_str}  →  {formula}")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    # Write using USER_ENTERED so Sheets evaluates the formulas
    data = [{"range": r, "values": [[f]]} for r, f in all_updates]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=sheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": data}
    ).execute()
    print("\nDone.")


if __name__ == "__main__":
    main()
