"""Claude AI-powered batch GTD email classification."""

from __future__ import annotations

import json
from typing import Callable, Optional

from .database import GTDCategory

BATCH_SIZE = 20

_BASE_SYSTEM = """\
You are a GTD (Getting Things Done) expert. Classify emails into the correct GTD bucket.

Buckets and when to use them:
- trash       : Spam, marketing, newsletters the user doesn't read, low-value automated notifications
- reference   : Receipts, confirmations, documentation, useful newsletters — no action needed
- someday     : Interesting opportunities or ideas but no current commitment
- waiting     : User is expecting a reply or someone else must act before they can proceed
- next_action : Requires a specific action from the USER soon — reply, decision, task, approval
- project     : Involves multiple steps over time or an ongoing thread requiring coordination
- calendar    : Has a specific meeting time, deadline, or time-sensitive event
- inbox       : Truly ambiguous — cannot determine without more context

Decision rules (apply in order):
1. Marketing / promotional / mass-sent → trash
2. Receipt / order confirmation / bank statement / travel itinerary → reference
3. Meeting invite or time-boxed event → calendar
4. User is explicitly asked to do something → next_action
5. Ongoing multi-email thread about a goal → project
6. User sent something and is waiting → waiting
7. Interesting but no commitment now → someday
8. Keep everything else that is useful → reference
9. When genuinely unclear → inbox

Be decisive. The user has thousands of emails to triage.\
"""


def _build_system_prompt(examples: list[dict]) -> str:
    """Build the system prompt, optionally including the user's past decisions as examples."""
    if not examples:
        return _BASE_SYSTEM

    lines = [_BASE_SYSTEM, "\n\n--- USER'S PAST DECISIONS (learn from these patterns) ---\n"]
    for ex in examples:
        preview = (ex.get("body_preview") or "")[:150].replace("\n", " ")
        lines.append(
            f"Subject: {ex.get('subject') or '(no subject)'}\n"
            f"From: {ex.get('sender') or 'Unknown'}\n"
            f"Preview: {preview}\n"
            f"→ User classified as: {ex.get('category')}"
            + (f" | Notes: {ex['notes']}" if ex.get("notes") else "")
            + "\n"
        )
    lines.append("\nApply these patterns when classifying new emails.")
    return "\n".join(lines)


def _make_user_message(emails: list[dict]) -> str:
    lines = []
    for i, e in enumerate(emails, 1):
        preview = (e.get("body_preview") or "")[:250].replace("\n", " ")
        lines.append(
            f"Email {i}:\n"
            f"  Subject : {e.get('subject') or '(no subject)'}\n"
            f"  From    : {e.get('sender') or 'Unknown'} <{e.get('sender_email') or ''}>\n"
            f"  Date    : {e.get('received_at') or ''}\n"
            f"  Preview : {preview}\n"
        )

    return (
        f"Classify these {len(emails)} emails.\n\n"
        + "\n".join(lines)
        + "\n\nRespond with a JSON array — one object per email in the same order:\n"
        "[\n"
        "  {\n"
        '    "category": "trash|reference|someday|waiting|next_action|project|calendar|inbox",\n'
        '    "reasoning": "one sentence",\n'
        '    "next_action": "specific next action if next_action or project, else empty string",\n'
        '    "notes": "any useful context to capture, else empty string"\n'
        "  }\n"
        "]\n\n"
        "Return ONLY the JSON array. No markdown fences, no extra text."
    )


def _parse_response(text: str, expected: int) -> list[dict]:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        end = next((i for i, l in enumerate(lines) if i > 0 and l.startswith("```")), len(lines))
        text = "\n".join(lines[1:end])
    results: list[dict] = json.loads(text)
    if len(results) != expected:
        raise ValueError(f"AI returned {len(results)} results for {expected} emails")
    valid_cats = set(GTDCategory.ALL)
    for r in results:
        if r.get("category") not in valid_cats:
            r["category"] = GTDCategory.INBOX
    return results


def classify_batch(
    emails: list[dict],
    api_key: str,
    model: str = "claude-haiku-4-5-20251001",
    examples: Optional[list[dict]] = None,
) -> list[dict]:
    """Classify a batch of emails via Claude. Returns one result dict per email."""
    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    system = _build_system_prompt(examples or [])
    try:
        message = client.messages.create(
            model=model,
            max_tokens=4096,
            system=system,
            messages=[{"role": "user", "content": _make_user_message(emails)}],
        )
    except anthropic.AuthenticationError:
        raise RuntimeError(
            "Invalid Anthropic API key. Run 'gtd setup' to update it, "
            "or check console.anthropic.com for a valid key (starts with sk-ant-)."
        )
    except anthropic.BadRequestError as exc:
        raise RuntimeError(
            f"API request rejected (400): {exc}. "
            f"Try a different model with --model, e.g. --model claude-haiku-4-5-20251001"
        )
    except anthropic.APIError as exc:
        raise RuntimeError(f"Anthropic API error: {exc}")
    return _parse_response(message.content[0].text, len(emails))


def classify_all(
    emails: list[dict],
    api_key: str,
    examples: Optional[list[dict]] = None,
    model: str = "claude-haiku-4-5-20251001",
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> list[tuple[dict, dict]]:
    """
    Classify all emails in batches.
    Returns list of (email, result) pairs.
    progress_callback(done, total) is called after each batch.
    """
    pairs: list[tuple[dict, dict]] = []
    total = len(emails)

    for start in range(0, total, BATCH_SIZE):
        batch = emails[start : start + BATCH_SIZE]
        results = classify_batch(batch, api_key, model=model, examples=examples)
        for email, result in zip(batch, results):
            pairs.append((email, result))
        if progress_callback:
            progress_callback(min(start + BATCH_SIZE, total), total)

    return pairs
