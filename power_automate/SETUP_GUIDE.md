# GTD Email — Power Automate Setup Guide

This flow automatically classifies every new email in your Outlook inbox using
Claude AI and moves it to the correct GTD folder — no terminal required.

---

## What it does

| Email type | What happens automatically |
|---|---|
| Spam / marketing | Moved to Deleted Items |
| Receipts / confirmations | Moved to GTD-Reference |
| Needs your action | Moved to GTD-Next-Actions + flagged |
| Waiting on someone | Moved to GTD-Waiting-For |
| Someday idea | Moved to GTD-Someday-Maybe |
| Multi-step project | Moved to GTD-Projects |
| Meeting / deadline | Moved to GTD-Calendar |
| Unclear | Left in Inbox for you to decide |

---

## Step 1 — Create the GTD folders in Outlook

Open Outlook and create these folders inside your Inbox:

- `GTD-Reference`
- `GTD-Next-Actions`
- `GTD-Waiting-For`
- `GTD-Someday-Maybe`
- `GTD-Projects`
- `GTD-Calendar`

Right-click **Inbox** → **Create new subfolder** for each one.

---

## Step 2 — Get your Anthropic API key

1. Go to **console.anthropic.com** → Sign in
2. Click **API Keys** → **Create Key**
3. Copy the key (starts with `sk-ant-...`) — you'll need it in step 5

---

## Step 3 — Open Power Automate

Go to **make.powerautomate.com** and sign in with your Microsoft account
(`jhalecpa@hotmail.com`).

---

## Step 4 — Create a new flow

1. Click **+ Create** → **Automated cloud flow**
2. Name it: `GTD Email Classifier`
3. Search for trigger: **"When a new email arrives (V3)"** → Select it → Click **Create**

---

## Step 5 — Build the flow

Add these steps in order using **+ New step**:

### Step 5a — Compose (email preview)
- Search: **Compose**
- Inputs: click in the box → **Expression** tab → paste:
  ```
  substring(triggerBody()?['body'], 0, 300)
  ```

### Step 5b — HTTP (call Claude AI)
- Search: **HTTP**
- Method: `POST`
- URI: `https://api.anthropic.com/v1/messages`
- Headers (add each one):
  | Key | Value |
  |---|---|
  | `x-api-key` | `sk-ant-YOUR_KEY_HERE` |
  | `anthropic-version` | `2023-06-01` |
  | `content-type` | `application/json` |
- Body (paste exactly):
```json
{
  "model": "claude-opus-4-8",
  "max_tokens": 256,
  "system": "You are a GTD expert. Classify emails into exactly one: trash, reference, someday, waiting, next_action, project, calendar, inbox. Rules: 1) Marketing/promo→trash 2) Receipt/confirmation→reference 3) Meeting/deadline→calendar 4) User must act→next_action 5) Multi-step→project 6) Waiting on someone→waiting 7) Interesting idea→someday 8) Useful info→reference 9) Unclear→inbox. Respond with ONLY JSON: {\"category\": \"...\", \"reasoning\": \"one sentence\", \"next_action\": \"specific action or empty\"}",
  "messages": [
    {
      "role": "user",
      "content": "Subject: @{triggerBody()?['subject']}\nFrom: @{triggerBody()?['from']}\nDate: @{triggerBody()?['receivedDateTime']}\nPreview: @{outputs('Compose')}\n\nClassify this email."
    }
  ]
}
```
> **Note:** Replace `sk-ant-YOUR_KEY_HERE` with your actual Anthropic API key.

### Step 5c — Parse JSON
- Search: **Parse JSON**
- Content: Click **Expression** → paste:
  ```
  body('HTTP')?['content']?[0]?['text']
  ```
- Schema: Click **Generate from sample** and paste:
  ```json
  {"category": "next_action", "reasoning": "example", "next_action": "Reply to email"}
  ```

### Step 5d — Switch
- Search: **Switch**
- On: Click **Expression** → paste:
  ```
  body('Parse_JSON')?['category']
  ```
- Add a **Case** for each category and the action inside it:

| Case value | Action to add inside |
|---|---|
| `trash` | **Delete email (V2)** → Message Id: `Id` from trigger |
| `reference` | **Move email (V2)** → Message Id: `Id`, Folder: `GTD-Reference` |
| `next_action` | **Move email (V2)** → `GTD-Next-Actions`, then **Flag email** → Flagged |
| `waiting` | **Move email (V2)** → `GTD-Waiting-For` |
| `someday` | **Move email (V2)** → `GTD-Someday-Maybe` |
| `project` | **Move email (V2)** → `GTD-Projects` |
| `calendar` | **Move email (V2)** → `GTD-Calendar` |
| (Default) | Leave empty — email stays in Inbox |

---

## Step 6 — Save and turn on

1. Click **Save**
2. Click **Test** → **Manually** → send yourself a test email
3. Watch it get sorted automatically
4. Once confirmed working, the flow runs 24/7 automatically

---

## Costs

| Service | Cost |
|---|---|
| Power Automate | Free (personal plan includes 750 runs/month) |
| Claude AI (Anthropic) | ~$0.001 per email (very cheap) |

For 1,000 emails/month → about **$1/month** in AI costs.

---

## Troubleshooting

**Flow didn't trigger:** Check the flow is turned **On** (toggle in top right).

**HTTP step failed:** Double-check your Anthropic API key is correct and has credits.

**Email not moving:** Make sure the GTD folder names match exactly (case-sensitive).

**Parse JSON failed:** The AI returned something unexpected — check the HTTP step's
output in the run history to see the raw response.

---

## Using the CLI alongside Power Automate

Power Automate handles **new** emails automatically going forward.
Use the CLI tool to process your **existing backlog** first:

```bash
gtd capture --max 500
gtd ai-clarify
```

After that, Power Automate takes over for everything new.
