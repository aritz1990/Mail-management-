# Pitch Deck Processor — Setup Guide

This automation watches your Gmail inbox and automatically saves startup funding pitch decks to a Google Drive folder, using Claude AI to identify them.

---

## What you need

- A Mac with Google Drive for Desktop installed and signed in
- A Google Cloud account (free) to create Gmail API credentials
- An Anthropic API key (get one at console.anthropic.com)
- Claude Code installed

---

## Step 1 — Place these files

Create a folder somewhere on your Mac (e.g. `~/pitch-deck-processor/`) and put all of these files inside it:

```
pitch-deck-processor/
├── CLAUDE.md
├── README.md
├── processor.py
├── authenticate.py
├── contacts.md
```

Then open Claude Code with this folder as your working directory.

---

## Step 2 — Create your Pitch Decks folder in Google Drive

Create a folder in your Google Drive where pitch decks will be saved. Note the full local path — on macOS with Google Drive for Desktop it will look like:

```
/Users/YOUR_NAME/Library/CloudStorage/GoogleDrive-YOUR_EMAIL/Shared drives/FOLDER/Pitch Decks
```

Set this as an environment variable so the script knows where to save:

```bash
echo 'export PITCH_DECKS_FOLDER="/Users/YOUR_NAME/Library/CloudStorage/..."' >> ~/.zshrc
source ~/.zshrc
```

---

## Step 3 — Install Python dependencies

```bash
python3 -m venv venv
./venv/bin/pip install google-auth-oauthlib google-api-python-client anthropic pypdf python-pptx
```

---

## Step 4 — Get Gmail API credentials

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (name it anything)
3. Enable the **Gmail API**: APIs & Services → Library → search "Gmail API" → Enable
4. Create credentials: APIs & Services → Credentials → **+ Create Credentials** → OAuth client ID
   - Configure consent screen if prompted: External, fill in app name and your email
   - Application type: **Desktop app**
5. Download the JSON file, rename it **`credentials.json`**, and put it in your project folder

---

## Step 5 — Authenticate Gmail (one time only)

Run:

```bash
./venv/bin/python3 authenticate.py
```

A URL will be printed. Visit it in your browser, approve Gmail access. Your browser will redirect to `http://localhost:8888` — the script catches this automatically and saves a `token.pickle` file. You won't need to do this again unless the token expires.

---

## Step 6 — Add your Anthropic API key

```bash
echo 'export ANTHROPIC_API_KEY="sk-ant-..."' >> ~/.zshrc
source ~/.zshrc
```

---

## Step 7 — Test it

```bash
./venv/bin/python3 processor.py
```

It will scan your Gmail for emails with PDF/PPTX attachments received in the last hour and process them.

---

## Step 8 — Set up the hourly schedule in Claude Code

Ask Claude Code:
> "Create a scheduled task that runs `./venv/bin/python3 processor.py` every hour from this folder"

Claude Code will set up the automation. From that point on it runs automatically in the background.

---

## Processing a specific sender's history

To process all historical emails from a specific person, ask Claude Code:
> "Process all emails from name@example.com and save any pitch decks"

Claude Code can run a modified version of the script targeting that sender across your full email history.

---

## Already-processed companies

See `contacts.md` for the list of companies whose pitch decks have already been saved by Tom's setup. If you want to avoid duplicates when processing shared contacts, cross-reference against this list.

---

## Troubleshooting

| Error | Fix |
|---|---|
| `credentials.json not found` | Complete Step 4 above |
| `ANTHROPIC_API_KEY not set` | Complete Step 6 above |
| `ModuleNotFoundError` | Run `./venv/bin/pip install ...` from Step 3 |
| Token expired | Run `./venv/bin/python3 authenticate.py` again |
| Wrong folder path | Check `PITCH_DECKS_FOLDER` env var or edit `processor.py` directly |
