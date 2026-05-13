# Pitch Deck Processor — Claude Code Context

## What this project does

This automation runs daily via GitHub Actions and checks Gmail for emails containing startup funding pitch decks. When it finds one, it:

1. Scans emails for PDF/PPTX attachments **and** DocSend links (`docsend.com/view/...`)
2. Downloads the deck (directly for attachments, via docsend2pdf.com API for DocSend links)
3. Extracts text and asks Claude (Opus 4.6) whether it is a genuine funding pitch deck
4. If yes, saves it to **two Google Drive folders** simultaneously (ar@angelinvest.ventures and anna.ritz@legata.cc)
5. Matches the company in **Attio CRM** and updates the Pitch deck URL field
6. Sends an email notification if the Attio match is ambiguous, with View and Confirm buttons

## File structure

```
Mail-management-/
├── CLAUDE.md           ← you are here
├── processor.py        ← main automation script
├── attio.py            ← Attio CRM API client (never deletes data)
├── authenticate.py     ← one-time Gmail OAuth helper (run from Cloud Shell)
├── credentials.json    ← Gmail OAuth credentials (not committed)
├── token.pickle        ← saved Gmail auth token (not committed)
├── processed.json      ← log of processed Gmail message IDs (not committed, cached by GitHub Actions)
└── .github/workflows/processor.yml  ← daily GitHub Actions cron
```

## Key configuration (top of processor.py)

| Variable | What to change |
|---|---|
| `DRIVE_FOLDER_IDS` | List of Google Drive folder IDs to save decks into |
| `NOTIFICATION_EMAIL` | Email address for ambiguous match notifications |
| `DOCSEND_EMAIL` | Email passed to docsend2pdf.com for access-controlled links |
| Gmail query in `process_emails()` | To filter by sender, date range, label, etc. |
| Claude model in `analyze_deck()` | Currently `claude-opus-4-6` |

## GitHub Actions schedule

The processor runs daily at **9am UTC** via `.github/workflows/processor.yml`.
To trigger a manual run: GitHub → Actions → Pitch Deck Processor → Run workflow.

## GitHub secrets required

| Secret | Purpose |
|---|---|
| `ANTHROPIC_API_KEY` | Claude API access |
| `ATTIO_API_KEY` | Attio CRM API access |
| `CREDENTIALS_JSON` | Google OAuth credentials JSON |
| `TOKEN_PICKLE_B64` | Base64-encoded Gmail token (regenerate with authenticate.py) |
| `APPS_SCRIPT_URL` | Google Apps Script endpoint for Confirm buttons in emails |
| `CONFIRM_TOKEN` | Secret token validating confirm button clicks |

## Re-authenticating Gmail

The token needs to be regenerated if it expires or if OAuth scopes change.
Run from Google Cloud Shell (shell.cloud.google.com):

```bash
cd Mail-management-
git pull
OAUTHLIB_INSECURE_TRANSPORT=1 python3 authenticate.py
base64 -w 0 token.pickle > token_b64.txt
cloudshell download token_b64.txt
```

Then update the `TOKEN_PICKLE_B64` GitHub secret with the downloaded file contents.

Current OAuth scopes: `gmail.readonly`, `gmail.send`, `drive` (full Drive access).

## Attio integration

- `attio.py` matches companies by: domain → exact name → partial name
- Single confident match owned by Anna Ritz → updates Pitch deck URL automatically
- Single match owned by someone else → sends notification email
- No match → creates a new company record
- Ambiguous match → sends notification email with View and Confirm buttons per candidate
- The Confirm button calls a Google Apps Script web app (script.google.com, project: "Attio Confirm")

## Drive folders

| Folder | Owner |
|---|---|
| `1bg0NQVwuP82wkIWvXzlJCrs-WHYk12DD` | ar@angelinvest.ventures |
| `13CKApFKyLmlcl90Sa-xEXMcaqHnkz3tB` | anna.ritz@legata.cc (shared with ar@angelinvest.ventures) |

Files are uploaded to folder 1, then copied to folder 2.
