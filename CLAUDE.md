# Pitch Deck Processor — Claude Code Context

## What this project does

This automation checks Gmail hourly for emails containing startup funding pitch decks. When it finds one, it:

1. Downloads the attachment (PDF or PPTX)
2. Extracts the text
3. Asks Claude (Opus 4.6) whether it is genuinely a funding pitch deck
4. If yes, saves it to a designated Google Drive folder
5. Renames the file to the company's URL domain (e.g. `sprive.com.pdf`)

## Your role as Claude Code

When the user asks you to work on this project you should be able to:

- **Run the processor manually** — execute `processor.py` to scan Gmail right now
- **Check processed companies** — read `contacts.md` to see what has already been saved
- **Re-authenticate** — run `authenticate.py` if the Gmail token expires
- **Adjust the detection logic** — edit the Claude prompt inside `processor.py` if the user wants stricter or looser pitch deck filtering
- **Process a specific sender** — modify the Gmail query in `processor.py` to target a specific email address
- **Rename or reorganise files** — rename files in the Pitch Decks Google Drive folder based on user instruction
- **Schedule or reschedule** — create or update a Claude Code scheduled task to run `processor.py` on an hourly cron

## File structure

```
pitch-deck-processor/
├── CLAUDE.md           ← you are here
├── README.md           ← human setup guide
├── processor.py        ← main automation script
├── authenticate.py     ← one-time Gmail OAuth helper
├── contacts.md         ← list of companies already processed
├── credentials.json    ← Gmail OAuth credentials (user must supply)
├── token.pickle        ← saved Gmail auth token (auto-created after auth)
├── processed.json      ← log of processed Gmail message IDs
└── venv/               ← Python virtual environment
```

## Key configuration (top of processor.py)

| Variable | What to change |
|---|---|
| `DRIVE_PITCH_DECKS_FOLDER` | Path to the Google Drive Pitch Decks folder |
| Gmail query in `process_emails()` | To filter by sender, date range, label, etc. |
| Claude model | Currently `claude-opus-4-6` — change if needed |

## Running the processor

```bash
./venv/bin/python3 processor.py
```

The `ANTHROPIC_API_KEY` environment variable must be set. Add it to `~/.zshrc` if not already there:

```bash
echo 'export ANTHROPIC_API_KEY="sk-ant-..."' >> ~/.zshrc
source ~/.zshrc
```

## Scheduled task

The processor runs hourly via a Claude Code scheduled task. To create or recreate it, ask Claude Code:
> "Create a scheduled task that runs processor.py every hour"

## Re-authenticating Gmail

If the token expires, run:

```bash
./venv/bin/python3 authenticate.py
```

This will print a URL. Visit it, approve Gmail access, and the browser will redirect to `http://localhost:8888` which the script catches automatically.

## Already-processed companies

See `contacts.md` for the full list of companies whose pitch decks have already been saved. The processor tracks processed email IDs in `processed.json` to avoid duplicates.
