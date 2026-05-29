# GTD Email Manager

A Python CLI application that applies the **Getting Things Done (GTD)** methodology to your
Outlook/Hotmail inbox using the Microsoft Graph API.

## Features

| GTD Phase   | Command          | Description |
|-------------|------------------|-------------|
| Capture     | `gtd capture`    | Fetch new emails from Outlook inbox into local DB |
| Clarify     | `gtd clarify`    | Interactively categorise each unprocessed email |
| Organise    | `gtd next-actions` / `gtd waiting` / etc. | View each GTD bucket |
| Reflect     | `gtd review`     | Weekly review dashboard with counts per bucket |
| Engage      | `gtd engage`     | Work through your Next Actions list |

### GTD Buckets

- **Inbox** — captured but not yet clarified
- **Trash** — not actionable, not useful
- **Reference** — keep for future reference (non-actionable)
- **Someday / Maybe** — good idea, but not now
- **Waiting For** — delegated or awaiting a reply
- **Next Actions** — concrete next step, do it or schedule it
- **Projects** — requires more than one action
- **Calendar** — has a specific date/time

---

## Installation

```bash
pip install -e .
```

Or install dependencies directly:

```bash
pip install -r requirements.txt
```

---

## Quick Start

### 1. Register an Azure Application

1. Go to <https://portal.azure.com> and sign in with **any** Microsoft account.
2. Search for **App registrations** → **New registration**.
3. Give it a name (e.g. `GTD Email Tool`).
4. Under *Supported account types* choose  
   **Accounts in any organizational directory … and personal Microsoft accounts**.
5. No Redirect URI needed.
6. After creation, copy the **Application (client) ID**.
7. Under **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**,  
   add `Mail.ReadWrite` and `Mail.Send`. Grant admin consent if prompted.

### 2. Configure GTD Email Tool

```bash
gtd setup
```

You will be prompted for:
- **Application (client) ID** — from Azure portal
- **Tenant ID** — leave as `consumers` for personal Hotmail/Outlook accounts
- **Your email address** — e.g. `jhalecpa@hotmail.com`

Configuration is stored at `~/.gtd/config.json` (permissions `600`).

### 3. Authenticate

```bash
gtd auth
```

This uses **device flow**: you'll be given a short code and a URL.  
Open <https://microsoft.com/devicelogin> in any browser, enter the code, and sign in.  
The access token is cached at `~/.gtd/token_cache.bin`.

### 4. Capture Emails

```bash
gtd capture          # fetch up to 50 emails
gtd capture --max 100  # fetch up to 100 emails
```

### 5. Clarify

```bash
gtd clarify
```

For each unprocessed email you will be shown its subject, sender, date, and preview.
Press the corresponding key to categorise it:

| Key | Category |
|-----|----------|
| `i` | Keep in Inbox (skip) |
| `t` | Trash / Delete |
| `r` | Reference |
| `s` | Someday / Maybe |
| `w` | Waiting For |
| `n` | Next Action |
| `p` | Project |
| `c` | Calendar |
| `q` | Quit session |

### 6. View Your Lists

```bash
gtd next-actions   # Next Actions
gtd waiting        # Waiting For
gtd projects       # Projects (+ add with --add)
gtd someday        # Someday / Maybe
gtd reference      # Reference archive
gtd status         # Quick summary of all buckets
```

### 7. Weekly Review

```bash
gtd review           # dashboard only
gtd review --save    # dashboard + save snapshot to DB
```

### 8. Work Through Actions

```bash
gtd engage
```

Walks you through each Next Action and lets you mark them done (archives in Outlook).

---

## Data Storage

| File | Purpose |
|------|---------|
| `~/.gtd/config.json` | Client ID, tenant ID, email address |
| `~/.gtd/token_cache.bin` | Cached OAuth2 tokens |
| `~/.gtd/gtd.db` | SQLite database with all GTD state |

Both sensitive files are stored with `600` permissions (owner read/write only).

---

## GTD Outlook Folders

When an email is categorised, GTD Email Tool automatically moves it to a corresponding
Outlook folder (created on first use):

| GTD Category  | Outlook Folder |
|---------------|----------------|
| Reference     | GTD-Reference |
| Someday/Maybe | GTD-Someday-Maybe |
| Waiting For   | GTD-Waiting-For |
| Next Actions  | GTD-Next-Actions |
| Projects      | GTD-Projects |

---

## Commands Reference

```
gtd --help
gtd setup           Configure Microsoft Graph credentials
gtd auth            Sign in via device flow
gtd capture         Fetch emails from Outlook inbox
gtd clarify         Process unprocessed emails interactively
gtd next-actions    Show Next Actions list
gtd waiting         Show Waiting For list
gtd projects        Show Projects list  (--add to create a project)
gtd someday         Show Someday/Maybe list
gtd reference       Show Reference archive
gtd review          Weekly review dashboard  (--save to persist snapshot)
gtd engage          Work through Next Actions interactively
gtd status          Quick bucket summary
```

---

## Requirements

- Python 3.10+
- `msal` — Microsoft Authentication Library
- `requests` — HTTP client
- `click` — CLI framework
- `rich` — Beautiful terminal output
