# Weekly Sales Velocity Report

## Objective
Generate a weekly sales velocity report from a manually exported Toast POS CSV and email it to tom@diamondbackbeer.com.

## Trigger
- **Manual**: Drop both location CSVs into `.tmp/`, then run:
  ```
  python3 "tools/run_report.py"
  ```
- The script processes whichever CSVs are present and skips any that are missing

## Required Inputs
Two CSV files dropped into `.tmp/`:

| Location | Expected filename |
|---|---|
| Locust Point | `locust_point.csv` |
| Timonium | `timonium.csv` |

Each CSV covers the 7 days prior to the current date, exported from Toast Back Office.

## How to Export from Toast (repeat for each location)
1. Log in to Toast Back Office (pos.toasttab.com)
2. Navigate to **Reports → Menu Item Reports → Item Sales**
3. Set date range to the previous 7 days
4. Click **Export** → download CSV
5. Rename the file `locust_point.csv` or `timonium.csv` and drop it in `.tmp/`

## Adding a New Location
Edit `LOCATIONS` at the top of `tools/run_report.py`:
```python
LOCATIONS = {
    "Locust Point": "locust_point.csv",
    "Timonium":     "timonium.csv",
    "New Location": "new_location.csv",   # ← add here
}
```

## Tools (in order)
1. `tools/parse_toast_csv.py` — reads `.tmp/toast_export.csv`, cleans and normalizes data
2. `tools/calculate_velocity.py` — calculates units sold, daily average, weekly total per item
3. `tools/build_report.py` — formats data into an HTML email report
4. `tools/send_email.py` — sends the report via Gmail API

## Expected Output
Two separate HTML emails, one per location:
- Subject: `Weekly Sales Velocity Report — Locust Point – Mar 08, 2026`
- Subject: `Weekly Sales Velocity Report — Timonium – Mar 08, 2026`

Each email contains:
- Location name and report period in the header
- Table of all menu items sorted by weekly units sold (descending)
- Columns: Item Name | Units Sold (Week) | Daily Avg

## Edge Cases & Known Behaviors
- If `.tmp/toast_export.csv` is missing, the script exits with a clear error message
- Items with 0 units sold are excluded from the report
- Toast CSV column names may vary by account — update `parse_toast_csv.py` if columns don't match

## Credentials
- Place `credentials.json` (downloaded from Google Cloud Console) in the project root
- On first run, a browser window will open to authorize Gmail access — `token.json` is saved after that
- Both files are gitignored

## File Locations
- Input: `.tmp/locust_point.csv`, `.tmp/timonium.csv`
- Intermediate: `.tmp/locust_point_parsed_items.json`, `.tmp/locust_point_report_data.json`, etc.
- Output: Emailed directly (no local file saved)
