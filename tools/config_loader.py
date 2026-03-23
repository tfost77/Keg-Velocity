#!/usr/bin/env python3
"""Shared config loader.

Checks INVENTORY_CONFIG env var first (for Streamlit Cloud where config.json
isn't committed), then falls back to config.json on disk.

Usage:
  from config_loader import load_config
  CONFIG = load_config()
"""

import json
import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent


def load_config():
    env_config = os.getenv("INVENTORY_CONFIG")
    if env_config:
        return json.loads(env_config)
    config_path = BASE_DIR / "config.json"
    if not config_path.exists():
        print("ERROR: config.json not found. Copy config.example.json to config.json and customize it.")
        sys.exit(1)
    with open(config_path) as f:
        return json.load(f)
