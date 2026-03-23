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
    """Return an authenticated Google Sheets service.

    Auth priority:
      1. Service account via GOOGLE_SERVICE_ACCOUNT_JSON (headless / Streamlit Cloud)
      2. OAuth refresh token via GOOGLE_CLIENT_ID + GOOGLE_CLIENT_SECRET + GOOGLE_REFRESH_TOKEN
         (Streamlit Cloud when service account keys are blocked by org policy)
      3. token.json on disk (local dev — existing OAuth session)
      4. OAuth browser flow via credentials.json (local dev — first run)
    """
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    # Option 1: service account key JSON
    sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if sa_json:
        from google.oauth2 import service_account
        info = json.loads(sa_json)
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        return build("sheets", "v4", credentials=creds)

    # Option 2: OAuth refresh token stored as env vars (no key file needed)
    client_id     = os.getenv("GOOGLE_CLIENT_ID")
    client_secret = os.getenv("GOOGLE_CLIENT_SECRET")
    refresh_token = os.getenv("GOOGLE_REFRESH_TOKEN")
    if client_id and client_secret and refresh_token:
        creds = Credentials(
            token=None,
            refresh_token=refresh_token,
            token_uri="https://oauth2.googleapis.com/token",
            client_id=client_id,
            client_secret=client_secret,
            scopes=SCOPES,
        )
        creds.refresh(Request())
        return build("sheets", "v4", credentials=creds)

    # Option 3 & 4: local dev — token.json or browser flow
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
