#!/usr/bin/env python3
"""
Calculate sales velocity from parsed items.

Usage:
  python calculate_velocity.py --location "Locust Point"

Input:  .tmp/{location_slug}_parsed_items.json
Output: .tmp/{location_slug}_report_data.json
"""

import argparse
import json
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"

DAYS_IN_PERIOD = 7

# Category display order in the report
CATEGORY_ORDER = ["Pint", "Half Pour", "Mug", "To-Go", "Other"]


def slugify(name):
    return name.lower().replace(" ", "_")


def classify(name):
    if re.search(r'pack', name, re.IGNORECASE):
        return "To-Go"
    if re.match(r'^\d+M[\s-]', name):
        return "Mug"
    if re.match(r'^\d+H[\s-]', name):
        return "Half Pour"
    if re.match(r'^\d+ -', name):
        return "Pint"
    return "Other"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--location", default="", help='Location name, e.g. "Locust Point"')
    args = parser.parse_args()

    slug = slugify(args.location) if args.location else ""
    prefix = f"{slug}_" if slug else ""

    in_path = TMP_DIR / f"{prefix}parsed_items.json"
    if not in_path.exists():
        print(f"ERROR: {in_path.name} not found. Run parse_toast_csv.py first.")
        sys.exit(1)

    with open(in_path) as f:
        items = json.load(f)

    report_items = []
    for item in items:
        units = item["units_sold"]
        daily_avg = round(units / DAYS_IN_PERIOD, 1)
        report_items.append({
            "name": item["name"],
            "category": classify(item["name"]),
            "units_sold": units,
            "daily_avg": daily_avg,
        })

    # Sort by category order, then by units sold descending within each category
    report_items.sort(key=lambda x: (
        CATEGORY_ORDER.index(x["category"]),
        -x["units_sold"]
    ))

    out_path = TMP_DIR / f"{prefix}report_data.json"
    with open(out_path, "w") as f:
        json.dump(report_items, f, indent=2)

    print(f"Calculated velocity for {len(report_items)} items → {out_path.name}")


if __name__ == "__main__":
    main()
