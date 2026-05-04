#!/usr/bin/env python3
"""One-time helper to generate a Google OAuth refresh token.

Usage:
    pip install --user google-auth-oauthlib
    python tools/get_refresh_token.py /path/to/client_secret.json

The script opens a browser, asks you to grant Gmail + Calendar access, and
prints the three values you'll paste into GitHub Actions secrets.
"""
import json
import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/calendar.readonly",
]


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python get_refresh_token.py <path-to-client-secret.json>")
        return 1

    secret_path = Path(sys.argv[1]).expanduser()
    if not secret_path.exists():
        print(f"File not found: {secret_path}")
        return 1

    flow = InstalledAppFlow.from_client_secrets_file(str(secret_path), SCOPES)
    creds = flow.run_local_server(
        port=0,
        prompt="consent",
        access_type="offline",
        authorization_prompt_message="",
        success_message="Authorization complete. You can close this tab.",
    )

    if not creds.refresh_token:
        print("ERROR: no refresh token returned.")
        print("Revoke prior access at https://myaccount.google.com/permissions")
        print("then re-run this script.")
        return 1

    with open(secret_path) as f:
        client = json.load(f)
    info = client.get("installed") or client.get("web") or {}

    print()
    print("=" * 64)
    print("SUCCESS - copy these three values into GitHub Actions secrets:")
    print("=" * 64)
    print(f"GOOGLE_CLIENT_ID:      {info.get('client_id')}")
    print(f"GOOGLE_CLIENT_SECRET:  {info.get('client_secret')}")
    print(f"GOOGLE_REFRESH_TOKEN:  {creds.refresh_token}")
    print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main())
