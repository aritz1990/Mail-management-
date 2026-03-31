#!/usr/bin/env python3
"""
Pitch Deck Email Processor
Checks Gmail hourly for emails containing pitch decks (funding round decks),
analyses them with Claude, and saves qualifying attachments directly to Google Drive.

SETUP: See README.md for first-time setup instructions.
CONFIG: Set PITCH_DECKS_FOLDER_NAME env var, or edit DRIVE_FOLDER_NAME below.
"""

import os
import sys
import io
import json
import base64
import pickle
import hashlib
from datetime import datetime, timedelta, timezone
from pathlib import Path

# Gmail & Drive API
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Claude
import anthropic

# File extraction
try:
    from pypdf import PdfReader
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    from pptx import Presentation
    PPTX_SUPPORT = True
except ImportError:
    PPTX_SUPPORT = False

# ── Config ─────────────────────────────────────────────────────────────────────

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
TOKEN_FILE = SCRIPT_DIR / "token.pickle"
PROCESSED_LOG = SCRIPT_DIR / "processed.json"

# Name of the Google Drive folder where pitch decks will be saved.
# Override with the PITCH_DECKS_FOLDER_NAME environment variable if needed.
DRIVE_FOLDER_NAME = os.environ.get("PITCH_DECKS_FOLDER_NAME", "Pitch Decks")

# Only attachments with these MIME types are considered
PITCH_DECK_MIME_TYPES = {
    "application/pdf",
    "application/vnd.ms-powerpoint",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}

# Both Gmail (read) and Drive (upload) scopes
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

# ── Auth ───────────────────────────────────────────────────────────────────────

def get_credentials():
    """Load or refresh OAuth credentials, prompting via console if needed."""
    creds = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                print("ERROR: credentials.json not found.")
                print("Follow the setup guide: see README.md in this folder.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES
            )
            # Console-based flow — no localhost server required (works from any device)
            auth_url, _ = flow.authorization_url(prompt="consent")
            print("\nAuthorisation required. Visit this URL in your browser:")
            print(f"\n  {auth_url}\n")
            code = input("Paste the authorisation code here: ").strip()
            flow.fetch_token(code=code)
            creds = flow.credentials

        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return creds


def get_gmail_service():
    return build("gmail", "v1", credentials=get_credentials())


def get_drive_service():
    return build("drive", "v3", credentials=get_credentials())


# ── Drive helpers ──────────────────────────────────────────────────────────────

def get_or_create_drive_folder(drive_service, folder_name: str) -> str:
    """Return the Drive folder ID, creating the folder if it doesn't exist."""
    safe = folder_name.replace("'", "\\'")
    results = drive_service.files().list(
        q=f"name='{safe}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id, name)",
        spaces="drive",
    ).execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]
    meta = {"name": folder_name, "mimeType": "application/vnd.google-apps.folder"}
    folder = drive_service.files().create(body=meta, fields="id").execute()
    print(f"  Created Drive folder '{folder_name}'")
    return folder["id"]


def upload_to_drive(drive_service, folder_id: str, filename: str, data: bytes, mime_type: str) -> dict:
    """Upload bytes directly to a Drive folder. Returns the file resource."""
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    return drive_service.files().create(
        body=meta, media_body=media, fields="id, name, webViewLink"
    ).execute()


# ── Processed log ──────────────────────────────────────────────────────────────

def load_processed() -> set:
    if PROCESSED_LOG.exists():
        with open(PROCESSED_LOG) as f:
            return set(json.load(f))
    return set()


def save_processed(processed: set):
    with open(PROCESSED_LOG, "w") as f:
        json.dump(list(processed), f, indent=2)


# ── Text extraction ────────────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> str:
    if not PDF_SUPPORT:
        return "[PDF text extraction unavailable — install pypdf]"
    try:
        reader = PdfReader(io.BytesIO(data))
        pages = [page.extract_text() or "" for page in reader.pages[:20]]
        return "\n".join(pages)[:8000]
    except Exception as e:
        return f"[PDF extraction error: {e}]"


def extract_text_from_pptx(data: bytes) -> str:
    if not PPTX_SUPPORT:
        return "[PPTX text extraction unavailable — install python-pptx]"
    try:
        prs = Presentation(io.BytesIO(data))
        texts = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texts.append(shape.text)
        return "\n".join(texts)[:8000]
    except Exception as e:
        return f"[PPTX extraction error: {e}]"


def extract_text(data: bytes, mime_type: str) -> str:
    if mime_type == "application/pdf":
        return extract_text_from_pdf(data)
    if mime_type in (
        "application/vnd.ms-powerpoint",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ):
        return extract_text_from_pptx(data)
    return ""


# ── Claude analysis ────────────────────────────────────────────────────────────

def is_pitch_deck(email_subject: str, email_body: str, attachment_name: str, attachment_text: str) -> tuple[bool, str]:
    """
    Ask Claude whether this email contains a genuine funding pitch deck.
    Returns (is_pitch_deck: bool, reasoning: str)
    """
    client = anthropic.Anthropic()

    prompt = f"""You are analysing an email to determine whether it contains a startup or company funding pitch deck (i.e. a deck used to raise investment).

EMAIL SUBJECT:
{email_subject}

EMAIL BODY (first 2000 chars):
{email_body[:2000]}

ATTACHMENT FILENAME:
{attachment_name}

ATTACHMENT TEXT (first 4000 chars):
{attachment_text[:4000]}

Decide: is this attachment a funding/investment pitch deck for a startup or company?

Respond with a JSON object only, no other text:
{{
  "is_pitch_deck": true or false,
  "confidence": "high" | "medium" | "low",
  "reasoning": "one sentence explanation"
}}"""

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=256,
        messages=[{"role": "user", "content": prompt}],
    )

    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    result = json.loads(text)
    return result.get("is_pitch_deck", False), result.get("reasoning", "")


# ── Gmail helpers ──────────────────────────────────────────────────────────────

def get_email_body(payload) -> str:
    if payload.get("mimeType") == "text/plain":
        data = payload.get("body", {}).get("data", "")
        return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
    for part in payload.get("parts", []):
        body = get_email_body(part)
        if body:
            return body
    return ""


def download_attachment(service, message_id: str, attachment_id: str) -> bytes:
    attachment = service.users().messages().attachments().get(
        userId="me", messageId=message_id, id=attachment_id
    ).execute()
    data = attachment.get("data", "")
    return base64.urlsafe_b64decode(data + "==")


def sanitise_filename(name: str) -> str:
    return "".join(c for c in name if c not in r'\/:*?"<>|').strip()


# ── Main processing ────────────────────────────────────────────────────────────

def process_emails():
    gmail_service = get_gmail_service()
    drive_service = get_drive_service()
    folder_id = get_or_create_drive_folder(drive_service, DRIVE_FOLDER_NAME)
    processed = load_processed()

    cutoff = datetime.now(timezone.utc) - timedelta(minutes=65)
    after_epoch = int(cutoff.timestamp())
    query = f"has:attachment after:{after_epoch}"

    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Searching Gmail: {query}")

    results = gmail_service.users().messages().list(
        userId="me", q=query, maxResults=50
    ).execute()

    messages = results.get("messages", [])
    print(f"  Found {len(messages)} emails with attachments in the last hour.")

    saved_count = 0

    for msg_ref in messages:
        msg_id = msg_ref["id"]

        if msg_id in processed:
            continue

        msg = gmail_service.users().messages().get(
            userId="me", id=msg_id, format="full"
        ).execute()

        headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
        subject = headers.get("Subject", "(no subject)")
        sender = headers.get("From", "unknown")
        body = get_email_body(msg["payload"])

        print(f"\n  Email: '{subject}' from {sender}")

        parts = msg["payload"].get("parts", [])
        for part in parts:
            filename = part.get("filename", "")
            mime_type = part.get("mimeType", "")
            attachment_id = part.get("body", {}).get("attachmentId")

            if not attachment_id or mime_type not in PITCH_DECK_MIME_TYPES:
                continue

            print(f"    Attachment: {filename} ({mime_type})")

            attachment_data = download_attachment(gmail_service, msg_id, attachment_id)
            attachment_text = extract_text(attachment_data, mime_type)

            result, reasoning = is_pitch_deck(subject, body, filename, attachment_text)
            print(f"    Claude says pitch deck: {result} — {reasoning}")

            if result:
                safe_name = sanitise_filename(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                upload_name = f"{timestamp}_{safe_name}"

                uploaded = upload_to_drive(
                    drive_service, folder_id, upload_name, attachment_data, mime_type
                )
                print(f"    Saved to Drive: {uploaded.get('name')} — {uploaded.get('webViewLink', '')}")
                saved_count += 1

        processed.add(msg_id)

    save_processed(processed)
    print(f"\nDone. Saved {saved_count} pitch deck(s).")


if __name__ == "__main__":
    process_emails()
