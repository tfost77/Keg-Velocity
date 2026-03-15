#!/usr/bin/env python3
"""
Orchestrator for the weekly sales velocity report.
Runs the full pipeline for each configured location, then syncs combined
can sales to the Cans Inventory tab.

Usage:
  python3 tools/run_report.py
  python3 tools/run_report.py --overwrite   # replace last filled Sales Week

CSV files expected in .tmp/:
  locust_point.csv
  timonium.csv
"""

import argparse
import subprocess
import sys
from pathlib import Path

TOOLS_DIR = Path(__file__).parent

# Add or remove locations here. Key = display name, value = CSV filename in .tmp/
LOCATIONS = {
    "Locust Point": "locust_point.csv",
    "Timonium":     "timonium.csv",
}

STEPS = [
    "parse_toast_csv.py",
    "calculate_velocity.py",
    "build_report.py",
    "sync_to_sheets.py",
    "send_email.py",
]

STEP_LABELS = {
    "parse_toast_csv.py":    "Parsing Toast CSV",
    "calculate_velocity.py": "Calculating velocity",
    "build_report.py":       "Building report",
    "sync_to_sheets.py":     "Syncing to Google Sheets (draft beer)",
    "send_email.py":         "Sending email",
}


def run_step(script, location, csv_path=None, overwrite=False):
    label = STEP_LABELS[script]
    print(f"  → {label}...")
    cmd = [sys.executable, str(TOOLS_DIR / script), "--location", location]
    if script == "parse_toast_csv.py" and csv_path:
        cmd += ["--csv", str(csv_path)]
    if script == "sync_to_sheets.py" and overwrite:
        cmd += ["--overwrite"]
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\n  ERROR: {script} failed (exit code {result.returncode}). Skipping remaining steps for {location}.")
        return False
    return True


def run_cans_sync(overwrite=False):
    print("  → Syncing to Google Sheets (cans)...")
    cmd = [sys.executable, str(TOOLS_DIR / "sync_cans_to_sheets.py")]
    if overwrite:
        cmd += ["--overwrite"]
    result = subprocess.run(cmd)
    if result.returncode != 0:
        print(f"\n  ERROR: sync_cans_to_sheets.py failed (exit code {result.returncode}).")
        return False
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--overwrite", action="store_true", help="Overwrite the last filled Sales Week row instead of appending")
    args = parser.parse_args()

    tmp_dir = TOOLS_DIR.parent / ".tmp"
    print("=== Weekly Sales Velocity Report ===\n")

    any_ran = False
    for location, csv_filename in LOCATIONS.items():
        csv_path = tmp_dir / csv_filename
        if not csv_path.exists():
            print(f"[{location}] SKIPPED — {csv_filename} not found in .tmp/")
            continue

        print(f"[{location}]")
        any_ran = True
        for script in STEPS:
            ok = run_step(script, location, csv_path if script == "parse_toast_csv.py" else None, overwrite=args.overwrite)
            if not ok:
                break
        print()

    if not any_ran:
        print("No CSV files found. Drop locust_point.csv and/or timonium.csv into .tmp/ and re-run.")
        sys.exit(1)

    # Combined cans sync runs once after all locations are processed
    print("[Cans Inventory]")
    run_cans_sync(overwrite=args.overwrite)
    print()

    print("✓ Done.")


if __name__ == "__main__":
    main()
