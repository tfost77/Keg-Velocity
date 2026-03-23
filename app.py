#!/usr/bin/env python3
"""
Keg Velocity Report — Streamlit front-end.

Upload Toast CSV exports for each location and sync sales data to Google Sheets.
"""

import json
import os
import subprocess
import sys
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent
TMP_DIR = BASE_DIR / ".tmp"

load_dotenv(BASE_DIR / ".env")


def load_config():
    env_config = os.getenv("INVENTORY_CONFIG")
    if env_config:
        return json.loads(env_config)
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        st.error(
            "**config.json not found.**\n\n"
            "Either add `config.json` to the repo, or set the `INVENTORY_CONFIG` "
            "secret in Streamlit Cloud with the full contents of your config.json."
        )
        st.stop()
    with open(config_path) as f:
        return json.load(f)


def run_pipeline(overwrite: bool):
    cmd = [sys.executable, str(BASE_DIR / "tools" / "run_report.py")]
    if overwrite:
        cmd.append("--overwrite")

    output_box = st.empty()
    lines = []

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        cwd=str(BASE_DIR),
    )

    for line in proc.stdout:
        lines.append(line.rstrip())
        output_box.code("\n".join(lines), language=None)

    proc.wait()
    return proc.returncode


# ── Page setup ────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Keg Velocity Report", layout="centered")
st.title("Keg Velocity Report")
st.caption("Upload Toast exports for each location, then click Run Report.")

config = load_config()
locations = config.get("locations", {})

TMP_DIR.mkdir(exist_ok=True)

# ── File uploaders ─────────────────────────────────────────────────────────────

cols = st.columns(len(locations))
uploaded = {}

for col, (location, csv_filename) in zip(cols, locations.items()):
    with col:
        f = st.file_uploader(location, type="csv", key=location)
        if f:
            uploaded[location] = (csv_filename, f)

st.divider()

# ── Controls ───────────────────────────────────────────────────────────────────

overwrite = st.checkbox(
    "Overwrite last week instead of appending",
    help="Use this if you're correcting data you already synced this week.",
)

run_btn = st.button(
    "Run Report",
    type="primary",
    disabled=len(uploaded) == 0,
    help="Upload at least one CSV to enable.",
)

# ── Pipeline execution ─────────────────────────────────────────────────────────

if run_btn:
    for location, (filename, fileobj) in uploaded.items():
        (TMP_DIR / filename).write_bytes(fileobj.read())

    st.subheader("Output")

    with st.spinner("Running pipeline..."):
        returncode = run_pipeline(overwrite)

    if returncode == 0:
        st.success("Done — data synced to Google Sheets.")
    else:
        st.error("Pipeline failed. See output above for details.")
