#!/usr/bin/env python3
"""
One-time Gmail + Google Drive OAuth authentication helper.

Run this once to generate token.pickle. After that, processor.py will
use the saved token automatically and refresh it when needed.
"""

import pickle
import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
TOKEN_FILE = SCRIPT_DIR / "token.pickle"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]


def main():
    if not CREDENTIALS_FILE.exists():
        print("ERROR: credentials.json not found.")
        print("Download it from Google Cloud Console and place it in this folder.")
        print("See README.md for instructions.")
        sys.exit(1)

    flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)

    # Console-based flow — no localhost server required (works from any device/OS)
    auth_url, _ = flow.authorization_url(prompt="consent")
    print("\nVisit this URL in your browser to authorise Gmail + Drive access:")
    print(f"\n  {auth_url}\n")
    print("After approving, Google will show you an authorisation code.")
    code = input("Paste the authorisation code here: ").strip()

    flow.fetch_token(code=code)
    creds = flow.credentials

    with open(TOKEN_FILE, "wb") as f:
        pickle.dump(creds, f)

    print(f"\nAuthentication successful! Token saved to: {TOKEN_FILE}")
    print("You can now run processor.py")


if __name__ == "__main__":
    main()
