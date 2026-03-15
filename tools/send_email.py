#!/usr/bin/env python3
"""
Send the HTML report via Gmail API using OAuth2.

Usage:
  python send_email.py --location "Locust Point"

Input:    .tmp/{location_slug}_report.html
Requires: credentials.json in project root (from Google Cloud Console)
          .env with GMAIL_SENDER and GMAIL_RECIPIENT
"""

import argparse
import base64
import os
import sys
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent.parent
TMP_DIR = BASE_DIR / ".tmp"
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
]

load_dotenv(BASE_DIR / ".env")


def slugify(name):
    return name.lower().replace(" ", "_")


def get_gmail_service():
    creds = None
    token_path = BASE_DIR / "token.json"
    creds_path = BASE_DIR / "credentials.json"

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

    return build("gmail", "v1", credentials=creds)


def send_email(service, sender, recipient, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.attach(MIMEText(html_body, "html"))

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"Email sent → {recipient}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--location", default="", help='Location name, e.g. "Locust Point"')
    args = parser.parse_args()

    slug = slugify(args.location) if args.location else ""
    prefix = f"{slug}_" if slug else ""

    report_path = TMP_DIR / f"{prefix}report.html"
    if not report_path.exists():
        print(f"ERROR: {report_path.name} not found. Run build_report.py first.")
        sys.exit(1)

    sender = os.getenv("GMAIL_SENDER")
    recipient = os.getenv("GMAIL_RECIPIENT")
    if not sender or not recipient:
        print("ERROR: GMAIL_SENDER and GMAIL_RECIPIENT must be set in .env")
        sys.exit(1)

    with open(report_path) as f:
        html_body = f.read()

    today = datetime.now().strftime("%b %d, %Y")
    location_label = f" — {args.location}" if args.location else ""
    subject = f"Weekly Sales Velocity Report{location_label} – {today}"

    print(f"Authenticating with Gmail API...")
    service = get_gmail_service()
    send_email(service, sender, recipient, subject, html_body)


if __name__ == "__main__":
    main()
