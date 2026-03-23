#!/usr/bin/env python3
"""Shared Google Sheets authentication helper.

Auth priority:
  1. Service account via GOOGLE_SERVICE_ACCOUNT_JSON env var (Streamlit Cloud / headless)
  2. OAuth token.json (local dev — existing token)
  3. OAuth browser flow via credentials.json (local dev — first run)

Usage:
  from auth import get_sheets_service
  service = get_sheets_service()
"""

import json
import os
import sys
from pathlib import Path

from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_sheets_service():
    """Return an authenticated Google Sheets service."""
    sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        from google.oauth2 import service_account
        info = json.loads(sa_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return build("sheets", "v4", credentials=creds)

    # OAuth fallback for local dev
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    token_path = BASE_DIR / "token.json"
    creds_path = BASE_DIR / "credentials.json"
    creds = None

    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not creds_path.exists():
                print(f"ERROR: credentials.json not found at {creds_path}")
                print("Download it from Google Cloud Console → APIs & Services → Credentials")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as f:
            f.write(creds.to_json())

    return build("sheets", "v4", credentials=creds)
