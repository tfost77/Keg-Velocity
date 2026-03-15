#!/usr/bin/env python3
"""
Build HTML email report from velocity data.

Usage:
  python build_report.py --location "Locust Point"

Input:  .tmp/{location_slug}_report_data.json
Output: .tmp/{location_slug}_report.html
"""

import argparse
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"


def slugify(name):
    return name.lower().replace(" ", "_")


def get_report_period():
    today = datetime.now()
    end = today - timedelta(days=1)       # Yesterday (Wednesday)
    start = end - timedelta(days=6)       # 7 days total
    return start.strftime("%b %d"), end.strftime("%b %d, %Y")


def group_by_category(items):
    groups = {}
    for item in items:
        cat = item.get("category", "Other")
        groups.setdefault(cat, []).append(item)
    return groups


def build_html(items, location=""):
    start, end = get_report_period()
    location_label = f" — {location}" if location else ""

    groups = group_by_category(items)
    category_order = ["Pint", "Half Pour", "Mug", "To-Go", "Other"]

    rows = ""
    row_index = 0
    for category in category_order:
        if category not in groups:
            continue
        # Category header row
        rows += f"""
        <tr>
          <td colspan="3" style="padding:12px 14px 6px 14px; font-weight:bold; font-size:13px;
              text-transform:uppercase; letter-spacing:0.05em; color:#666;
              border-top:2px solid #e0e0e0; background:#fff;">{category}</td>
        </tr>"""
        for item in groups[category]:
            bg = "#f9f9f9" if row_index % 2 == 0 else "#ffffff"
            rows += f"""
        <tr style="background-color:{bg};">
          <td style="padding:9px 14px; border-bottom:1px solid #eee;">{item['name']}</td>
          <td style="padding:9px 14px; text-align:center; border-bottom:1px solid #eee;">{item['units_sold']}</td>
          <td style="padding:9px 14px; text-align:center; border-bottom:1px solid #eee;">{item['daily_avg']}</td>
        </tr>"""
            row_index += 1

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family: Arial, sans-serif; color: #333; max-width: 720px; margin: 0 auto; padding: 24px;">

  <h2 style="margin:0 0 4px 0; color:#1a1a1a;">Weekly Sales Velocity Report{location_label}</h2>
  <p style="margin:0 0 20px 0; color:#888; font-size:14px;">
    Period: <strong>{start} – {end}</strong>
  </p>

  <table style="width:100%; border-collapse:collapse; font-size:14px;">
    <thead>
      <tr style="background-color:#1a1a1a; color:#ffffff;">
        <th style="padding:10px 14px; text-align:left;">Item</th>
        <th style="padding:10px 14px; text-align:center;">Units Sold (Week)</th>
        <th style="padding:10px 14px; text-align:center;">Daily Avg</th>
      </tr>
    </thead>
    <tbody>{rows}
    </tbody>
  </table>

  <p style="font-size:12px; color:#aaa; margin-top:20px;">
    {len(items)} items tracked &nbsp;·&nbsp; Generated from Toast POS export
  </p>

</body>
</html>"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--location", default="", help='Location name, e.g. "Locust Point"')
    args = parser.parse_args()

    slug = slugify(args.location) if args.location else ""
    prefix = f"{slug}_" if slug else ""

    in_path = TMP_DIR / f"{prefix}report_data.json"
    if not in_path.exists():
        print(f"ERROR: {in_path.name} not found. Run calculate_velocity.py first.")
        sys.exit(1)

    with open(in_path) as f:
        items = json.load(f)

    html = build_html(items, location=args.location)

    out_path = TMP_DIR / f"{prefix}report.html"
    with open(out_path, "w") as f:
        f.write(html)

    print(f"Built report with {len(items)} items → {out_path.name}")


if __name__ == "__main__":
    main()
