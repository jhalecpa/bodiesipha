"""Microsoft Graph API email operations for GTD Email Tool."""

from __future__ import annotations

import time
from typing import Any, Optional

import requests

from .auth import GRAPH_BASE_URL, get_access_token


class GraphAPIError(Exception):
    """Raised when the Microsoft Graph API returns an error."""

    def __init__(self, status_code: int, message: str, code: str = ""):
        self.status_code = status_code
        self.code = code
        super().__init__(f"Graph API error {status_code} ({code}): {message}")


def _headers() -> dict[str, str]:
    """Build HTTP headers with a fresh access token."""
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type": "application/json",
    }


def _get(url: str, params: Optional[dict] = None) -> Any:
    """Perform a GET request against the Graph API."""
    resp = requests.get(url, headers=_headers(), params=params, timeout=30)
    if not resp.ok:
        _raise_for_status(resp)
    return resp.json()


def _post(url: str, data: dict) -> Any:
    """Perform a POST request against the Graph API."""
    resp = requests.post(url, headers=_headers(), json=data, timeout=30)
    if not resp.ok:
        _raise_for_status(resp)
    return resp.json() if resp.text else {}


def _patch(url: str, data: dict) -> Any:
    """Perform a PATCH request against the Graph API."""
    resp = requests.patch(url, headers=_headers(), json=data, timeout=30)
    if not resp.ok:
        _raise_for_status(resp)
    return resp.json() if resp.text else {}


def _delete(url: str) -> None:
    """Perform a DELETE request against the Graph API."""
    resp = requests.delete(url, headers=_headers(), timeout=30)
    if not resp.ok:
        _raise_for_status(resp)


def _raise_for_status(resp: requests.Response) -> None:
    """Parse and raise a GraphAPIError from a failed response."""
    try:
        body = resp.json()
        err = body.get("error", {})
        message = err.get("message", resp.text)
        code = err.get("code", "")
    except Exception:
        message = resp.text
        code = ""
    raise GraphAPIError(resp.status_code, message, code)


# ---------------------------------------------------------------------------
# Email reading
# ---------------------------------------------------------------------------

def fetch_inbox_emails(max_count: int = 50, skip: int = 0) -> list[dict]:
    """
    Fetch emails from the inbox.

    Returns a list of message dicts with keys:
      id, subject, from, receivedDateTime, bodyPreview, isRead, hasAttachments
    """
    url = f"{GRAPH_BASE_URL}/me/mailFolders/Inbox/messages"
    params = {
        "$top": min(max_count, 50),
        "$skip": skip,
        "$select": (
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
            "bodyPreview,isRead,hasAttachments,importance,internetMessageId,"
            "conversationId,body"
        ),
        "$orderby": "receivedDateTime desc",
    }
    data = _get(url, params=params)
    return data.get("value", [])


def fetch_email_by_id(message_id: str) -> dict:
    """Fetch a single email by its Graph message ID."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
    return _get(url)


def get_message_body(message_id: str) -> str:
    """Return the plain-text body of a message."""
    msg = fetch_email_by_id(message_id)
    body = msg.get("body", {})
    content_type = body.get("contentType", "text")
    content = body.get("content", "")
    if content_type.lower() == "html":
        # Very basic HTML-to-text stripping
        import re
        content = re.sub(r"<[^>]+>", " ", content)
        content = re.sub(r"\s+", " ", content).strip()
    return content


# ---------------------------------------------------------------------------
# Folder management
# ---------------------------------------------------------------------------

def list_mail_folders() -> list[dict]:
    """List top-level mail folders."""
    url = f"{GRAPH_BASE_URL}/me/mailFolders"
    data = _get(url, params={"$top": 100})
    return data.get("value", [])


def get_or_create_folder(display_name: str) -> str:
    """
    Find an existing mail folder by display name or create it.
    Returns the folder ID.
    """
    folders = list_mail_folders()
    for folder in folders:
        if folder.get("displayName", "").lower() == display_name.lower():
            return folder["id"]

    # Create the folder
    url = f"{GRAPH_BASE_URL}/me/mailFolders"
    result = _post(url, {"displayName": display_name})
    return result["id"]


def ensure_gtd_folders() -> dict[str, str]:
    """
    Ensure all GTD folders exist in Outlook.
    Returns a mapping of GTD category name -> folder ID.
    """
    folder_names = {
        "gtd_reference": "GTD-Reference",
        "gtd_someday": "GTD-Someday-Maybe",
        "gtd_waiting": "GTD-Waiting-For",
        "gtd_next_actions": "GTD-Next-Actions",
        "gtd_projects": "GTD-Projects",
    }
    return {key: get_or_create_folder(name) for key, name in folder_names.items()}


# ---------------------------------------------------------------------------
# Email actions
# ---------------------------------------------------------------------------

def move_email_to_folder(message_id: str, folder_id: str) -> dict:
    """Move an email to the given folder."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/move"
    return _post(url, {"destinationId": folder_id})


def delete_email(message_id: str) -> None:
    """Permanently delete an email."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
    _delete(url)


def trash_email(message_id: str) -> dict:
    """Move an email to the Deleted Items folder."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/move"
    return _post(url, {"destinationId": "deleteditems"})


def archive_email(message_id: str) -> dict:
    """Move an email to the Archive folder."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/move"
    return _post(url, {"destinationId": "archive"})


def mark_as_read(message_id: str) -> dict:
    """Mark an email as read."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
    return _patch(url, {"isRead": True})


def mark_as_unread(message_id: str) -> dict:
    """Mark an email as unread."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}"
    return _patch(url, {"isRead": False})


def reply_to_email(message_id: str, comment: str) -> None:
    """Send a reply to a message."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/reply"
    _post(url, {"comment": comment})


def forward_email(message_id: str, to_addresses: list[str], comment: str = "") -> None:
    """Forward an email to one or more addresses."""
    url = f"{GRAPH_BASE_URL}/me/messages/{message_id}/forward"
    recipients = [{"emailAddress": {"address": addr}} for addr in to_addresses]
    _post(url, {"toRecipients": recipients, "comment": comment})


def send_email(to_addresses: list[str], subject: str, body: str, body_type: str = "Text") -> None:
    """Send a new email."""
    url = f"{GRAPH_BASE_URL}/me/sendMail"
    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": body_type, "content": body},
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to_addresses
            ],
        },
        "saveToSentItems": True,
    }
    _post(url, message)


# ---------------------------------------------------------------------------
# User profile
# ---------------------------------------------------------------------------

def get_me() -> dict:
    """Return the authenticated user's profile."""
    return _get(f"{GRAPH_BASE_URL}/me")
