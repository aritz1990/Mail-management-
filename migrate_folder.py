#!/usr/bin/env python3
"""
One-time script to move all files from the old Pitch Decks folder
to the new Pitch Decks folder in Google Drive.
"""
import pickle
from pathlib import Path
from googleapiclient.discovery import build

SCRIPT_DIR = Path(__file__).parent
TOKEN_FILE = SCRIPT_DIR / "token.pickle"

OLD_FOLDER_ID = "1SdbwIvYOFbpcqAeapalDANaILGvXZT9u"
NEW_FOLDER_ID = "1bg0NQVwuP82wkIWvXzlJCrs-WHYk12DD"

with open(TOKEN_FILE, "rb") as f:
    creds = pickle.load(f)

service = build("drive", "v3", credentials=creds)

results = service.files().list(
    q=f"'{OLD_FOLDER_ID}' in parents and trashed=false",
    fields="files(id, name)"
).execute()

files = results.get("files", [])
print(f"Found {len(files)} file(s) to move.")

for file in files:
    service.files().update(
        fileId=file["id"],
        addParents=NEW_FOLDER_ID,
        removeParents=OLD_FOLDER_ID,
        fields="id, name"
    ).execute()
    print(f"  Moved: {file['name']}")

print("Done.")
