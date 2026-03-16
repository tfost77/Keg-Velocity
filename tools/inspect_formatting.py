#!/usr/bin/env python3
"""Inspect cell formatting for a column range to understand the formatting pattern."""

import json
import os
from pathlib import Path
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
load_dotenv(BASE_DIR / ".env")

def get_sheets_service():
    creds = None
    token_path = BASE_DIR / "token.json"
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
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

def main():
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    service = get_sheets_service()
    tab = find_current_tab(service, sheet_id)
    print(f"Tab: '{tab}'\n")

    # Get the sheet numeric ID and full data with formatting for rows 1-55
    meta = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        ranges=[f"'{tab}'!A1:P55"],
        includeGridData=True
    ).execute()

    sheet_data = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == tab:
            sheet_data = s
            break

    if not sheet_data:
        print("Sheet not found")
        return

    rows = sheet_data.get("data", [{}])[0].get("rowData", [])

    # Print formatting summary for each row, each cell (cols A-P)
    col_labels = "ABCDEFGHIJKLMNOP"
    for row_idx, row in enumerate(rows):
        row_num = row_idx + 1
        cells = row.get("values", [])
        row_items = []
        for col_idx, cell in enumerate(cells):
            if col_idx >= 16:
                break
            col = col_labels[col_idx]
            val = ""
            if cell.get("formattedValue"):
                val = cell["formattedValue"][:15]
            fmt = cell.get("effectiveFormat", {})
            halign = fmt.get("horizontalAlignment", "")
            borders = fmt.get("borders", {})
            border_summary = ""
            for side in ["top", "bottom", "left", "right"]:
                b = borders.get(side, {})
                style = b.get("style", "")
                if style and style != "NONE":
                    border_summary += f"{side[0]}:{style} "
            bg = fmt.get("backgroundColor", {})
            bg_str = ""
            if bg and bg != {"red": 1, "green": 1, "blue": 1} and bg != {}:
                r = bg.get("red", 1)
                g = bg.get("green", 1)
                b_val = bg.get("blue", 1)
                if not (r == 1 and g == 1 and b_val == 1):
                    bg_str = f" bg({r:.1f},{g:.1f},{b_val:.1f})"
            bold = fmt.get("textFormat", {}).get("bold", False)
            bold_str = " BOLD" if bold else ""
            if val or halign or border_summary or bg_str or bold_str:
                row_items.append(f"{col}[{val}|{halign}{bold_str}{bg_str}|{border_summary.strip()}]")
        if row_items:
            print(f"Row {row_num:3d}: {' | '.join(row_items)}")

if __name__ == "__main__":
    main()
