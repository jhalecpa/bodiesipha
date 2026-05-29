"""SQLite database operations for GTD Email Tool."""

from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Generator, Optional


DB_DIR = Path.home() / ".gtd"
DB_PATH = DB_DIR / "gtd.db"

# GTD categories / buckets
class GTDCategory:
    INBOX = "inbox"               # Captured, not yet clarified
    TRASH = "trash"               # Not actionable, not useful
    REFERENCE = "reference"       # Non-actionable but keep for reference
    SOMEDAY = "someday"           # Not now, possible future action
    WAITING = "waiting"           # Delegated / awaiting reply
    NEXT_ACTION = "next_action"   # Actionable next actions
    PROJECT = "project"           # Multi-step projects
    CALENDAR = "calendar"         # Has a specific date/time

    ALL = [INBOX, TRASH, REFERENCE, SOMEDAY, WAITING, NEXT_ACTION, PROJECT, CALENDAR]

    DISPLAY_NAMES = {
        INBOX: "Inbox (Unprocessed)",
        TRASH: "Trash",
        REFERENCE: "Reference",
        SOMEDAY: "Someday / Maybe",
        WAITING: "Waiting For",
        NEXT_ACTION: "Next Actions",
        PROJECT: "Projects",
        CALENDAR: "Calendar",
    }


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS emails (
    id               TEXT PRIMARY KEY,
    message_id       TEXT NOT NULL,
    subject          TEXT,
    sender           TEXT,
    sender_email     TEXT,
    received_at      TEXT,
    body_preview     TEXT,
    is_read          INTEGER DEFAULT 0,
    has_attachments  INTEGER DEFAULT 0,
    category         TEXT NOT NULL DEFAULT 'inbox',
    next_action      TEXT,
    project_name     TEXT,
    due_date         TEXT,
    notes            TEXT,
    waiting_for      TEXT,
    processed_at     TEXT,
    folder_id        TEXT,
    raw_json         TEXT,
    created_at       TEXT NOT NULL,
    updated_at       TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_emails_category ON emails (category);
CREATE INDEX IF NOT EXISTS idx_emails_message_id ON emails (message_id);

CREATE TABLE IF NOT EXISTS projects (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    name         TEXT NOT NULL,
    description  TEXT,
    status       TEXT DEFAULT 'active',
    created_at   TEXT NOT NULL,
    updated_at   TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS weekly_reviews (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    review_date TEXT NOT NULL,
    notes       TEXT,
    inbox_count INTEGER,
    next_count  INTEGER,
    waiting_count INTEGER,
    projects_count INTEGER,
    someday_count  INTEGER,
    created_at  TEXT NOT NULL
);
"""


def _now() -> str:
    """Return current UTC time as an ISO 8601 string."""
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def get_db() -> Generator[sqlite3.Connection, None, None]:
    """Context manager that yields a database connection."""
    DB_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def initialize_db() -> None:
    """Create database tables if they don't exist."""
    with get_db() as conn:
        conn.executescript(SCHEMA_SQL)


def upsert_email(msg: dict, category: str = GTDCategory.INBOX) -> str:
    """
    Insert or update an email record from a Graph API message dict.
    Returns the internal row ID (message_id).
    """
    message_id = msg["id"]
    sender_name = ""
    sender_email = ""
    from_field = msg.get("from", {})
    if from_field:
        ea = from_field.get("emailAddress", {})
        sender_name = ea.get("name", "")
        sender_email = ea.get("address", "")

    now = _now()

    with get_db() as conn:
        existing = conn.execute(
            "SELECT id, category FROM emails WHERE message_id = ?", (message_id,)
        ).fetchone()

        if existing:
            # Don't overwrite the category if it's already been processed
            existing_cat = existing["category"]
            conn.execute(
                """UPDATE emails SET subject=?, sender=?, sender_email=?,
                   received_at=?, body_preview=?, is_read=?, has_attachments=?,
                   raw_json=?, updated_at=?
                   WHERE message_id=?""",
                (
                    msg.get("subject", "(no subject)"),
                    sender_name,
                    sender_email,
                    msg.get("receivedDateTime", ""),
                    msg.get("bodyPreview", ""),
                    1 if msg.get("isRead") else 0,
                    1 if msg.get("hasAttachments") else 0,
                    json.dumps(msg),
                    now,
                    message_id,
                ),
            )
        else:
            conn.execute(
                """INSERT INTO emails
                   (id, message_id, subject, sender, sender_email, received_at,
                    body_preview, is_read, has_attachments, category, raw_json,
                    created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    message_id,
                    message_id,
                    msg.get("subject", "(no subject)"),
                    sender_name,
                    sender_email,
                    msg.get("receivedDateTime", ""),
                    msg.get("bodyPreview", ""),
                    1 if msg.get("isRead") else 0,
                    1 if msg.get("hasAttachments") else 0,
                    category,
                    json.dumps(msg),
                    now,
                    now,
                ),
            )
    return message_id


def update_email_category(
    message_id: str,
    category: str,
    next_action: str = "",
    project_name: str = "",
    due_date: str = "",
    notes: str = "",
    waiting_for: str = "",
    folder_id: str = "",
) -> None:
    """Update the GTD category and metadata for an email."""
    now = _now()
    with get_db() as conn:
        conn.execute(
            """UPDATE emails SET category=?, next_action=?, project_name=?,
               due_date=?, notes=?, waiting_for=?, folder_id=?,
               processed_at=?, updated_at=?
               WHERE message_id=?""",
            (
                category,
                next_action,
                project_name,
                due_date,
                notes,
                waiting_for,
                folder_id,
                now,
                now,
                message_id,
            ),
        )


def get_emails_by_category(category: str) -> list[sqlite3.Row]:
    """Return all emails in a given GTD category, newest first."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT * FROM emails WHERE category=? ORDER BY received_at DESC",
            (category,),
        ).fetchall()
    return rows


def get_unprocessed_emails() -> list[sqlite3.Row]:
    """Return emails still in the inbox (not yet clarified)."""
    return get_emails_by_category(GTDCategory.INBOX)


def get_email_by_message_id(message_id: str) -> Optional[sqlite3.Row]:
    """Return a single email row by its Graph message ID."""
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM emails WHERE message_id=?", (message_id,)
        ).fetchone()


def count_by_category() -> dict[str, int]:
    """Return a dict mapping each GTD category to its email count."""
    with get_db() as conn:
        rows = conn.execute(
            "SELECT category, COUNT(*) as cnt FROM emails GROUP BY category"
        ).fetchall()
    counts = {cat: 0 for cat in GTDCategory.ALL}
    for row in rows:
        counts[row["category"]] = row["cnt"]
    return counts


def delete_email_record(message_id: str) -> None:
    """Remove an email record from the local database."""
    with get_db() as conn:
        conn.execute("DELETE FROM emails WHERE message_id=?", (message_id,))


def add_project(name: str, description: str = "") -> int:
    """Add a new project and return its ID."""
    now = _now()
    with get_db() as conn:
        cursor = conn.execute(
            "INSERT INTO projects (name, description, created_at, updated_at) VALUES (?, ?, ?, ?)",
            (name, description, now, now),
        )
        return cursor.lastrowid


def get_projects(status: str = "active") -> list[sqlite3.Row]:
    """Return projects with the given status."""
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM projects WHERE status=? ORDER BY created_at DESC",
            (status,),
        ).fetchall()


def save_weekly_review(notes: str = "") -> int:
    """Save a weekly review snapshot and return its ID."""
    counts = count_by_category()
    now = _now()
    with get_db() as conn:
        cursor = conn.execute(
            """INSERT INTO weekly_reviews
               (review_date, notes, inbox_count, next_count, waiting_count,
                projects_count, someday_count, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                now[:10],
                notes,
                counts.get(GTDCategory.INBOX, 0),
                counts.get(GTDCategory.NEXT_ACTION, 0),
                counts.get(GTDCategory.WAITING, 0),
                counts.get(GTDCategory.PROJECT, 0),
                counts.get(GTDCategory.SOMEDAY, 0),
                now,
            ),
        )
        return cursor.lastrowid
