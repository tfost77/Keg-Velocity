#!/usr/bin/env python3
"""
One-time repair script to align beer column headers across all sections
(Total Volume, DBC Locust Point, DBC Timonium) in the current inventory tab.

Canonical column order is derived from the master (Total Volume) Beer: row.
Any beers found only in a location section are appended at the end.

All data rows in each section (Keg Volume, Sales Week rows, etc.) are remapped
to match the canonical column positions.

Usage:
  python repair_column_alignment.py              # dry run (preview only)
  python repair_column_alignment.py --write      # apply changes
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
    if not config_path.exists():
        print("ERROR: config.json not found.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)


CONFIG = load_config()
MASTER_SECTION_HEADER = CONFIG.get("master_section_header")
LOCATION_SECTION_HEADERS = CONFIG["location_section_headers"]


def col_letter(n):
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(r + ord("A")) + result
    return result


def normalize_name(name):
    """Normalize quote characters so names can be compared reliably."""
    return name.replace("\u2019", "'").replace("\u2018", "'")


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


def read_all_rows(service, sheet_id, tab, formulas=False):
    result = service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"'{tab}'!A1:AZ100",
        valueRenderOption="FORMULA" if formulas else "FORMATTED_VALUE"
    ).execute()
    return result.get("values", [])


def find_section_start(rows, header):
    for i, row in enumerate(rows):
        if row and normalize_name(row[0]) == normalize_name(header):
            return i
    return None


def find_beer_header_row_idx(rows, section_start, search_range=6):
    """Find the row index with 'Beer:' in col A within search_range rows of section_start."""
    for i in range(section_start + 1, min(section_start + search_range, len(rows))):
        if rows[i] and normalize_name(rows[i][0]) == "Beer:":
            return i
    return None


def parse_beer_header_row(rows, row_idx):
    """Return {normalized_beer_name: col_index} for the Beer: header row."""
    if row_idx is None or row_idx >= len(rows):
        return {}
    return {
        normalize_name(cell): j
        for j, cell in enumerate(rows[row_idx])
        if cell and normalize_name(cell) != "Beer:" and j > 0
    }


def find_section_data_rows(rows, beer_header_row, next_section_start):
    """Return all row indices after beer_header_row (up to next section) that have col A content."""
    end = next_section_start if next_section_start is not None else len(rows)
    return [
        i for i in range(beer_header_row + 1, end)
        if i < len(rows) and rows[i] and rows[i][0]
    ]


def build_canonical_columns(master_cols, all_location_cols):
    """
    Build canonical {beer: col_index} from master order,
    appending location-only beers at the end.
    """
    canonical = dict(sorted(master_cols.items(), key=lambda x: x[1]))
    next_col = max(canonical.values()) + 1 if canonical else 1
    for loc_cols in all_location_cols:
        for beer in sorted(loc_cols, key=lambda b: loc_cols[b]):
            if beer not in canonical:
                canonical[beer] = next_col
                next_col += 1
    return canonical


def build_col_letter_map(old_cols, canonical):
    """Return {old_col_letter: new_col_letter} for beers that actually moved."""
    return {
        col_letter(old_col): col_letter(canonical[beer])
        for beer, old_col in old_cols.items()
        if beer in canonical and old_col != canonical[beer]
    }


def update_formula_refs(val, col_letter_map):
    """
    Rewrite column references in a formula string.
    e.g. with map {'L': 'M', 'M': 'N'}: =AVERAGE(L23:L27) → =AVERAGE(M23:M27)
    Replacement is done in a single pass to avoid double-substitution.
    """
    if not isinstance(val, str) or not val.startswith("=") or not col_letter_map:
        return val
    # Match longer col names first so 'AA' is replaced before 'A'
    sorted_keys = sorted(col_letter_map, key=len, reverse=True)
    pattern = r"(\$?)(" + "|".join(re.escape(c) for c in sorted_keys) + r")(\$?)(\d+)"
    return re.sub(pattern, lambda m: f"{m.group(1)}{col_letter_map[m.group(2)]}{m.group(3)}{m.group(4)}", val)


def remap_row(formula_row, old_cols, canonical, max_col, col_letter_map):
    """
    Return a new row with values (including formulas) moved to canonical columns.
    Formula column references are rewritten to match the new positions.
    col 0 (row label) is preserved.
    """
    new_row = [""] * (max_col + 1)
    if formula_row:
        new_row[0] = formula_row[0]
    for beer, old_col in old_cols.items():
        if beer in canonical:
            val = formula_row[old_col] if old_col < len(formula_row) else ""
            new_row[canonical[beer]] = update_formula_refs(val, col_letter_map)
    return new_row


def expand_sheet_if_needed(service, sheet_id, tab, needed_col_count):
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == tab:
            numeric_id = sheet["properties"]["sheetId"]
            current = sheet["properties"]["gridProperties"]["columnCount"]
            if needed_col_count > current:
                new_count = needed_col_count + 10
                service.spreadsheets().batchUpdate(
                    spreadsheetId=sheet_id,
                    body={"requests": [{"updateSheetProperties": {
                        "properties": {"sheetId": numeric_id, "gridProperties": {"columnCount": new_count}},
                        "fields": "gridProperties.columnCount"
                    }}]}
                ).execute()
                print(f"  Expanded sheet to {new_count} columns.")
            break


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

    rows = read_all_rows(service, sheet_id, tab)
    rows_with_formulas = read_all_rows(service, sheet_id, tab, formulas=True)

    # --- Locate master section ---
    if not MASTER_SECTION_HEADER:
        print("ERROR: master_section_header not set in config.json")
        sys.exit(1)
    master_start = find_section_start(rows, MASTER_SECTION_HEADER)
    if master_start is None:
        print(f"ERROR: Master section '{MASTER_SECTION_HEADER}' not found in tab.")
        sys.exit(1)
    master_beer_row = find_beer_header_row_idx(rows, master_start)
    if master_beer_row is None:
        print(f"ERROR: No 'Beer:' row found in master section.")
        sys.exit(1)
    master_cols = parse_beer_header_row(rows, master_beer_row)
    print(f"Master '{MASTER_SECTION_HEADER}' — Beer: row at sheet row {master_beer_row + 1}, {len(master_cols)} beers")

    # --- Locate location sections ---
    location_info = {}
    all_section_starts = [master_start]
    for loc_name, section_header in LOCATION_SECTION_HEADERS.items():
        sec_start = find_section_start(rows, section_header)
        if sec_start is None:
            print(f"WARNING: Section '{section_header}' not found — skipping.")
            continue
        all_section_starts.append(sec_start)
        beer_row = find_beer_header_row_idx(rows, sec_start)
        if beer_row is None:
            print(f"WARNING: No 'Beer:' row in '{section_header}' — skipping.")
            continue
        loc_cols = parse_beer_header_row(rows, beer_row)
        location_info[loc_name] = {
            "header": section_header,
            "sec_start": sec_start,
            "beer_row": beer_row,
            "cols": loc_cols,
        }
        print(f"Location '{section_header}' — Beer: row at sheet row {beer_row + 1}, {len(loc_cols)} beers")

    print()

    # Sort section starts so we can find each section's end
    all_section_starts_sorted = sorted(set(all_section_starts))

    def next_section_start_after(start):
        idx = all_section_starts_sorted.index(start)
        return all_section_starts_sorted[idx + 1] if idx + 1 < len(all_section_starts_sorted) else None

    # --- Build canonical column order ---
    canonical = build_canonical_columns(master_cols, [info["cols"] for info in location_info.values()])
    max_col = max(canonical.values())

    print(f"Canonical column order ({len(canonical)} beers):")
    for beer, col in sorted(canonical.items(), key=lambda x: x[1]):
        in_master = "M" if beer in master_cols else " "
        in_locs = " ".join(
            name[0] for name, info in location_info.items() if beer in info["cols"]
        )
        print(f"  {col_letter(col)}: {beer}  [{in_master}|{in_locs}]")
    print()

    # --- Check and fix each section ---
    updates = []

    all_sections = [
        ("master", MASTER_SECTION_HEADER, master_start, master_beer_row, master_cols),
    ] + [
        (loc_name, info["header"], info["sec_start"], info["beer_row"], info["cols"])
        for loc_name, info in location_info.items()
    ]

    for section_name, header_text, sec_start, beer_row_idx, current_cols in all_sections:
        # Check if any remapping or header additions are needed
        moved = {b for b in current_cols if b in canonical and current_cols[b] != canonical[b]}
        missing = {b for b in canonical if b not in current_cols}

        if not moved and not missing:
            print(f"[OK] '{header_text}' — already aligned")
            continue

        print(f"[FIX] '{header_text}':")
        for b in sorted(moved, key=lambda b: current_cols[b]):
            print(f"  move '{b}': col {col_letter(current_cols[b])} → {col_letter(canonical[b])}")
        for b in sorted(missing, key=lambda b: canonical[b]):
            print(f"  add  '{b}' at col {col_letter(canonical[b])} (no data)")

        # Build corrected Beer: header row
        current_header = list(rows[beer_row_idx]) if beer_row_idx < len(rows) else []
        new_header = [""] * (max_col + 1)
        new_header[0] = current_header[0] if current_header else "Beer:"
        for beer, col in canonical.items():
            new_header[col] = beer
        r = f"'{tab}'!A{beer_row_idx + 1}:{col_letter(max_col)}{beer_row_idx + 1}"
        updates.append({"range": r, "values": [new_header], "desc": f"{header_text} Beer: header"})

        # Remap data rows only if columns actually moved (missing-only doesn't need remapping)
        if moved:
            clmap = build_col_letter_map(current_cols, canonical)
            nxt = next_section_start_after(sec_start)
            data_rows = find_section_data_rows(rows, beer_row_idx, nxt)
            for row_idx in data_rows:
                formula_row = list(rows_with_formulas[row_idx]) if row_idx < len(rows_with_formulas) else []
                new_row = remap_row(formula_row, current_cols, canonical, max_col, clmap)
                r = f"'{tab}'!A{row_idx + 1}:{col_letter(max_col)}{row_idx + 1}"
                label = formula_row[0] if formula_row else "?"
                updates.append({
                    "range": r,
                    "values": [new_row],
                    "input_option": "USER_ENTERED",
                    "desc": f"  data row {row_idx + 1} ({label})"
                })
            print(f"  remapping {len(data_rows)} data row(s) (formulas rewritten)")

    print()
    if not updates:
        print("Nothing to fix — all sections already aligned.")
        return

    print(f"{'Would apply' if dry_run else 'Applying'} {len(updates)} range update(s):")
    for u in updates:
        print(f"  {u['range']}  — {u['desc']}")

    if dry_run:
        print("\nRun with --write to apply.")
        return

    expand_sheet_if_needed(service, sheet_id, tab, max_col + 1)

    # Split updates by valueInputOption — headers use RAW, data rows use USER_ENTERED
    raw_updates = [{"range": u["range"], "values": u["values"]} for u in updates if u.get("input_option") != "USER_ENTERED"]
    entered_updates = [{"range": u["range"], "values": u["values"]} for u in updates if u.get("input_option") == "USER_ENTERED"]

    if raw_updates:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id,
            body={"valueInputOption": "RAW", "data": raw_updates}
        ).execute()
    if entered_updates:
        service.spreadsheets().values().batchUpdate(
            spreadsheetId=sheet_id,
            body={"valueInputOption": "USER_ENTERED", "data": entered_updates}
        ).execute()
    print("\nDone — all sections aligned.")


if __name__ == "__main__":
    main()
