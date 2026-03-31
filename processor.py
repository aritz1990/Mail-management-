#!/usr/bin/env python3
"""
Pitch Deck Email Processor
Checks Gmail hourly for emails containing pitch decks (funding round decks),
analyses them with Claude, and saves qualifying attachments to Google Drive.

SETUP: See README.md for first-time setup instructions.
CONFIG: Edit the DRIVE_PITCH_DECKS_FOLDER path below to match your Google Drive.
"""

import os
import sys
import json
import base64
import pickle
import hashlib
from datetime import datetime, timedelta, timezone
from pathlib import Path

# Gmail API
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# Claude
import anthropic

# File extraction
try:
    from pypdf import PdfReader
    import io as _io
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    from pptx import Presentation
    import io as _io
    PPTX_SUPPORT = True
except ImportError:
    PPTX_SUPPORT = False

# ── Config — edit these to match your environment ─────────────────────────────

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
TOKEN_FILE = SCRIPT_DIR / "token.pickle"
PROCESSED_LOG = SCRIPT_DIR / "processed.json"

# !! Change this to your own Google Drive Pitch Decks folder path !!
# On macOS with Google Drive for Desktop, it will look something like:
# /Users/YOUR_NAME/Library/CloudStorage/GoogleDrive-YOUR_EMAIL/Shared drives/FOLDER/Pitch Decks
DRIVE_PITCH_DECKS_FOLDER = Path(
    os.environ.get(
        "PITCH_DECKS_FOLDER",
        str(Path.home() / "Google Drive" / "Pitch Decks")  # fallback default
    )
)

# Only attachments with these MIME types are considered
PITCH_DECK_MIME_TYPES = {
    "application/pdf",
    "application/vnd.ms-powerpoint",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_gmail_service():
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
                str(CREDENTIALS_FILE), GMAIL_SCOPES
            )
            creds = flow.run_local_server(port=0, open_browser=False)
        with open(TOKEN_FILE, "wb") as f:
            pickle.dump(creds, f)

    return build("gmail", "v1", credentials=creds)


# ── Processed log ─────────────────────────────────────────────────────────────

def load_processed() -> set:
    if PROCESSED_LOG.exists():
        with open(PROCESSED_LOG) as f:
            return set(json.load(f))
    return set()


def save_processed(processed: set):
    with open(PROCESSED_LOG, "w") as f:
        json.dump(list(processed), f, indent=2)


# ── Text extraction ───────────────────────────────────────────────────────────

def extract_text_from_pdf(data: bytes) -> str:
    if not PDF_SUPPORT:
        return "[PDF text extraction unavailable — install pypdf]"
    try:
        reader = PdfReader(_io.BytesIO(data))
        pages = [page.extract_text() or "" for page in reader.pages[:20]]
        return "\n".join(pages)[:8000]
    except Exception as e:
        return f"[PDF extraction error: {e}]"


def extract_text_from_pptx(data: bytes) -> str:
    if not PPTX_SUPPORT:
        return "[PPTX text extraction unavailable — install python-pptx]"
    try:
        prs = Presentation(_io.BytesIO(data))
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


# ── Claude analysis ───────────────────────────────────────────────────────────

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


# ── Gmail helpers ─────────────────────────────────────────────────────────────

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


# ── Main processing ───────────────────────────────────────────────────────────

def process_emails():
    service = get_gmail_service()
    processed = load_processed()

    cutoff = datetime.now(timezone.utc) - timedelta(minutes=65)
    after_epoch = int(cutoff.timestamp())
    query = f"has:attachment after:{after_epoch}"

    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M')}] Searching Gmail: {query}")

    results = service.users().messages().list(
        userId="me", q=query, maxResults=50
    ).execute()

    messages = results.get("messages", [])
    print(f"  Found {len(messages)} emails with attachments in the last hour.")

    saved_count = 0

    for msg_ref in messages:
        msg_id = msg_ref["id"]

        if msg_id in processed:
            continue

        msg = service.users().messages().get(
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

            attachment_data = download_attachment(service, msg_id, attachment_id)
            attachment_text = extract_text(attachment_data, mime_type)

            result, reasoning = is_pitch_deck(subject, body, filename, attachment_text)
            print(f"    Claude says pitch deck: {result} — {reasoning}")

            if result:
                safe_name = sanitise_filename(filename)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M")
                dest = DRIVE_PITCH_DECKS_FOLDER / f"{timestamp}_{safe_name}"

                content_hash = hashlib.md5(attachment_data).hexdigest()[:8]
                if dest.exists():
                    stem = dest.stem
                    dest = DRIVE_PITCH_DECKS_FOLDER / f"{stem}_{content_hash}{dest.suffix}"

                DRIVE_PITCH_DECKS_FOLDER.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(attachment_data)
                print(f"    ✓ Saved to: {dest.name}")
                saved_count += 1

        processed.add(msg_id)

    save_processed(processed)
    print(f"\nDone. Saved {saved_count} pitch deck(s).")


if __name__ == "__main__":
    process_emails()
