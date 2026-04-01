#!/usr/bin/env python3
"""
One-time Gmail + Google Drive OAuth authentication helper.

Run this once to generate token.pickle. After that, processor.py will
use the saved token automatically and refresh it when needed.
"""

import pickle
import sys
import urllib.request
import urllib.parse
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

    # Save URL to file to avoid terminal truncation issues
    url_file = SCRIPT_DIR / "auth_url.txt"
    with open(url_file, "w") as f:
        f.write(auth_url)

    print("\nAuth URL saved to auth_url.txt")

    # Shorten the URL so it's easy to copy on any device
    try:
        short = urllib.request.urlopen(
            "https://tinyurl.com/api-create.php?url=" + urllib.parse.quote(auth_url, safe=""),
            timeout=5
        ).read().decode()
        print(f"\nOpen this short URL in your browser:\n\n  {short}\n")
    except Exception:
        print("\nCould not shorten URL. Open auth_url.txt in the editor and copy the full URL.")

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
