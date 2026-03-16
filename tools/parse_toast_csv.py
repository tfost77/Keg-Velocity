#!/usr/bin/env python3
"""
Parse Toast POS CSV export and filter to relevant items.

Usage:
  python parse_toast_csv.py --csv .tmp/locust_point.csv --location "Locust Point"

Input:  CSV file path via --csv (or auto-detected)
Output: .tmp/{location_slug}_parsed_items.json
"""

import argparse
import csv
import json
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"


def slugify(name):
    return name.lower().replace(" ", "_")


def load_config():
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        print("ERROR: config.json not found. Copy config.example.json to config.json and customize it.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)

CONFIG = load_config()
EXCLUDED_ITEMS = set(CONFIG["excluded_items"])


def find_csv(location_slug=None):
    # 1. Named after the location slug
    if location_slug:
        named = TMP_DIR / f"{location_slug}.csv"
        if named.exists():
            return named

    # 2. Explicitly named fallback
    default = TMP_DIR / "toast_export.csv"
    if default.exists():
        return default

    # 3. Most recently modified CSV in .tmp/
    tmp_csvs = list(TMP_DIR.glob("*.csv"))
    if tmp_csvs:
        return sorted(tmp_csvs, key=lambda f: f.stat().st_mtime, reverse=True)[0]

    # 4. Most recently modified CSV in project root
    root_csvs = list(BASE_DIR.glob("*.csv"))
    if root_csvs:
        return sorted(root_csvs, key=lambda f: f.stat().st_mtime, reverse=True)[0]

    return None


def is_beer_item(name):
    """Only include items that match Toast's beer naming conventions.
    Draft beer starts with a tap number (e.g. '2 - Green Machine', '2H - ...', '2M Gold - ...').
    Packaged beer ends with a pack size (e.g. 'Green Machine 6-Pack', 'Lager 12 Pack').
    Everything else (food, merch, event fees, pitchers, etc.) is excluded.
    """
    if re.match(r'^\d+', name):
        return True
    if re.search(r'\d+.?[Pp]ack', name):
        return True
    return False


def parse(csv_path):
    items = []
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.get("Item Name", "").strip()
            category = row.get("Sales Category", "").strip()

            # Skip excluded items and non-beer items
            if not name or name in EXCLUDED_ITEMS:
                continue
            if not is_beer_item(name):
                continue

            # Only use summary rows (empty Sales Category = total row for that item)
            if category != "":
                continue

            try:
                units = int(float(row.get("Item Qty", 0) or 0))
            except (ValueError, TypeError):
                units = 0

            # Skip items with zero sales
            if units == 0:
                continue

            items.append({"name": name, "units_sold": units})

    return items


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv", help="Path to Toast CSV export")
    parser.add_argument("--location", default="", help='Location name, e.g. "Locust Point"')
    args = parser.parse_args()

    slug = slugify(args.location) if args.location else ""
    csv_path = Path(args.csv) if args.csv else find_csv(slug)

    if not csv_path or not csv_path.exists():
        print(f"ERROR: No CSV file found for location '{args.location}'. Drop {slug}.csv into {TMP_DIR}/")
        sys.exit(1)

    print(f"Parsing [{args.location or 'default'}]: {csv_path.name}")
    items = parse(csv_path)

    TMP_DIR.mkdir(exist_ok=True)
    prefix = f"{slug}_" if slug else ""
    out_path = TMP_DIR / f"{prefix}parsed_items.json"
    with open(out_path, "w") as f:
        json.dump(items, f, indent=2)

    print(f"Parsed {len(items)} items → {out_path.name}")


if __name__ == "__main__":
    main()
