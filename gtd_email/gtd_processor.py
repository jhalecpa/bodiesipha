"""GTD logic and categorization for GTD Email Tool."""

from __future__ import annotations

import sqlite3
from typing import Optional

from .database import (
    GTDCategory,
    get_email_by_message_id,
    get_emails_by_category,
    get_unprocessed_emails,
    update_email_category,
    upsert_email,
)
from .email_client import (
    GraphAPIError,
    archive_email,
    ensure_gtd_folders,
    fetch_inbox_emails,
    mark_as_read,
    move_email_to_folder,
    trash_email,
)


# Map GTD categories to Outlook folder keys
GTD_FOLDER_MAP = {
    GTDCategory.REFERENCE: "gtd_reference",
    GTDCategory.SOMEDAY: "gtd_someday",
    GTDCategory.WAITING: "gtd_waiting",
    GTDCategory.NEXT_ACTION: "gtd_next_actions",
    GTDCategory.PROJECT: "gtd_projects",
}


class GTDProcessor:
    """Orchestrates GTD operations against the local DB and Outlook."""

    def __init__(self, offline: bool = False):
        """
        offline=True: Only use local DB, do not call Graph API.
        This is useful for unit tests or when offline.
        """
        self._offline = offline
        self._folders: Optional[dict[str, str]] = None

    @property
    def folders(self) -> dict[str, str]:
        """Lazily load GTD Outlook folders."""
        if self._folders is None and not self._offline:
            self._folders = ensure_gtd_folders()
        return self._folders or {}

    # ------------------------------------------------------------------
    # CAPTURE
    # ------------------------------------------------------------------

    def capture(self, max_emails: int = 50) -> tuple[int, int]:
        """
        Fetch emails from Outlook inbox and save new ones to the local DB.
        Returns (new_count, total_fetched).
        """
        if self._offline:
            raise RuntimeError("Cannot capture in offline mode.")

        messages = fetch_inbox_emails(max_count=max_emails)
        new_count = 0
        for msg in messages:
            from .database import get_email_by_message_id
            existing = get_email_by_message_id(msg["id"])
            if existing is None:
                upsert_email(msg, category=GTDCategory.INBOX)
                new_count += 1
            else:
                # Refresh metadata but keep the existing category
                upsert_email(msg, category=existing["category"])

        return new_count, len(messages)

    # ------------------------------------------------------------------
    # CLARIFY
    # ------------------------------------------------------------------

    def categorize(
        self,
        message_id: str,
        category: str,
        next_action: str = "",
        project_name: str = "",
        due_date: str = "",
        notes: str = "",
        waiting_for: str = "",
        sync_to_outlook: bool = True,
    ) -> None:
        """
        Set the GTD category for an email and optionally sync to Outlook.
        """
        if category not in GTDCategory.ALL:
            raise ValueError(f"Unknown GTD category: {category!r}")

        folder_id = ""
        if sync_to_outlook and not self._offline:
            folder_key = GTD_FOLDER_MAP.get(category)
            if folder_key:
                folder_id = self.folders.get(folder_key, "")
                if folder_id:
                    try:
                        move_email_to_folder(message_id, folder_id)
                    except GraphAPIError as exc:
                        # Non-fatal: local DB update still proceeds
                        print(f"[Warning] Could not move email in Outlook: {exc}")
            elif category == GTDCategory.TRASH:
                try:
                    trash_email(message_id)
                except GraphAPIError as exc:
                    print(f"[Warning] Could not trash email in Outlook: {exc}")
            # CALENDAR and INBOX categories stay in the Outlook inbox folder

        update_email_category(
            message_id=message_id,
            category=category,
            next_action=next_action,
            project_name=project_name,
            due_date=due_date,
            notes=notes,
            waiting_for=waiting_for,
            folder_id=folder_id,
        )

    # ------------------------------------------------------------------
    # REVIEW helpers
    # ------------------------------------------------------------------

    def get_bucket(self, category: str) -> list[sqlite3.Row]:
        """Return emails in a GTD bucket."""
        return get_emails_by_category(category)

    def get_unprocessed(self) -> list[sqlite3.Row]:
        """Return emails not yet clarified."""
        return get_unprocessed_emails()

    def get_two_minute_actions(self) -> list[sqlite3.Row]:
        """
        Return Next Actions that are quick (heuristic: body_preview is short,
        or the email itself is short / flagged as quick).
        A proper implementation would let the user flag them; here we
        surface all Next Actions so the user decides in the moment.
        """
        return get_emails_by_category(GTDCategory.NEXT_ACTION)

    def mark_action_done(self, message_id: str) -> None:
        """Mark a Next Action as completed (moves to Reference/Archive)."""
        if not self._offline:
            try:
                archive_email(message_id)
            except GraphAPIError as exc:
                print(f"[Warning] Could not archive email in Outlook: {exc}")
        update_email_category(message_id, GTDCategory.REFERENCE)

    def format_email_summary(self, row: sqlite3.Row) -> dict:
        """Return a friendly dict representation of a DB row."""
        return {
            "message_id": row["message_id"],
            "subject": row["subject"] or "(no subject)",
            "sender": row["sender"] or row["sender_email"] or "Unknown",
            "sender_email": row["sender_email"] or "",
            "received_at": (row["received_at"] or "")[:19].replace("T", " "),
            "body_preview": (row["body_preview"] or "")[:200],
            "is_read": bool(row["is_read"]),
            "has_attachments": bool(row["has_attachments"]),
            "category": row["category"],
            "next_action": row["next_action"] or "",
            "project_name": row["project_name"] or "",
            "due_date": row["due_date"] or "",
            "notes": row["notes"] or "",
            "waiting_for": row["waiting_for"] or "",
        }
