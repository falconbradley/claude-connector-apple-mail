"""
Apple Mail MCP Server
=====================
Exposes read-only access to Apple Mail via the Model Context Protocol so
Claude Desktop can search and read emails.

Communication with Mail.app is done through AppleScript/JXA, so there is
no need for Full Disk Access.  Mail.app must be running, and macOS
Automation permission must be granted (System Settings -> Privacy &
Security -> Automation) so this process can control Mail.app.

Tools provided
--------------
  get_stats             - Overview: total messages, unread, mailboxes, accounts
  list_mailboxes        - All accounts / folders with counts
  search_emails         - Rich search: text, sender, date, flags, mailbox
  get_email             - Full email with decoded plain-text body
  get_email_html        - HTML body of a specific email
  get_thread            - All emails in a conversation thread
  list_email_attachments - Enumerate attachments for an email
  get_email_attachment   - Download attachment as base64

Requirements
------------
  Mail.app must be running.  Automation permission must be granted to the
  host process (typically Claude Desktop, or Terminal if testing manually).
"""

from __future__ import annotations

import base64
import email as email_lib
import logging
import sys
from datetime import datetime
from typing import Optional
from urllib.parse import quote

from mcp.server.fastmcp import FastMCP

from .applescript import MailBridge
from .emlx import get_html_body
from .models import (
    Attachment,
    AttachmentData,
    EmailDetail,
    EmailSummary,
    Mailbox,
    MailboxStats,
    SearchResult,
)

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    stream=sys.stderr,
)
logger = logging.getLogger("apple_mail_mcp")

# ---------------------------------------------------------------------------
# Lazy-initialised shared state.  MailBridge init takes ~12-18s (mailbox
# prescan) so we MUST NOT run it at import time — the MCP client would
# time out waiting for the initialize response.
# ---------------------------------------------------------------------------

_bridge: Optional[MailBridge] = None

# ---------------------------------------------------------------------------
# FastMCP app
# ---------------------------------------------------------------------------

mcp = FastMCP(
    "Apple Mail",
    instructions=(
        "Read-only access to Apple Mail on this Mac via Mail.app. "
        "You can list mailboxes, search emails, read message bodies, "
        "list and retrieve attachments."
    ),
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _require_bridge() -> MailBridge:
    """Return the MailBridge instance, initialising on first call.

    Retries on every call if init previously failed (Mail.app may have
    been started or become responsive since the last attempt).
    """
    global _bridge
    if _bridge is not None:
        return _bridge
    try:
        _bridge = MailBridge()
        logger.info("Apple Mail MCP ready (AppleScript bridge).")
        return _bridge
    except (RuntimeError, OSError) as exc:
        raise RuntimeError(
            "Could not connect to Mail.app. Make sure Mail.app is open and "
            "that this process has Automation permission in System Settings "
            "-> Privacy & Security -> Automation."
            f"\n\nUnderlying error: {exc}"
        )


def _make_mail_link(rfc_id: Optional[str]) -> Optional[str]:
    """Build a message:// URL from an RFC 2822 Message-ID."""
    if not rfc_id:
        return None
    return f"message://{quote(f'<{rfc_id}>', safe='')}"


def _dict_to_summary(d: dict) -> EmailSummary:
    """Convert a MailBridge result dict to an EmailSummary model."""
    date_sent: Optional[datetime] = None
    date_received: Optional[datetime] = None

    if d.get("date_sent"):
        try:
            date_sent = datetime.fromisoformat(d["date_sent"])
        except (ValueError, TypeError):
            pass
    if d.get("date_received"):
        try:
            date_received = datetime.fromisoformat(d["date_received"])
        except (ValueError, TypeError):
            pass

    rfc_id = d.get("message_id") or None
    return EmailSummary(
        id=d["id"],
        mailbox=d.get("mailbox_name", ""),
        account=d.get("account_name", ""),
        subject=d.get("subject") or "(no subject)",
        sender=d.get("sender") or "",
        date_sent=date_sent,
        date_received=date_received,
        is_read=bool(d.get("is_read", True)),
        is_flagged=bool(d.get("is_flagged", False)),
        has_attachments=bool(d.get("has_attachments", False)),
        size=d.get("size", 0),
        message_id=rfc_id,
        in_reply_to=d.get("in_reply_to") or None,
        mail_link=_make_mail_link(rfc_id),
    )


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def get_stats() -> MailboxStats:
    """Return overall statistics: total messages, unread count, mailbox and account counts."""
    bridge = _require_bridge()
    stats = bridge.get_stats()
    return MailboxStats(
        total_messages=stats["total_messages"],
        unread_messages=stats["unread_messages"],
        mailbox_count=stats["mailbox_count"],
        account_count=stats["account_count"],
    )


@mcp.tool()
def list_mailboxes() -> list[Mailbox]:
    """List every mailbox (folder) Apple Mail knows about, with message counts."""
    bridge = _require_bridge()
    rows = bridge.list_mailboxes()
    return [
        Mailbox(
            name=r["name"],
            account=r["account_name"],
            full_name=f"{r['account_name']}/{r['name']}",
            unread_count=r.get("unread_count", 0),
            message_count=r.get("message_count", 0),
        )
        for r in rows
    ]


@mcp.tool()
def search_emails(
    query: Optional[str] = None,
    mailbox: Optional[str] = None,
    account: Optional[str] = None,
    from_address: Optional[str] = None,
    to_address: Optional[str] = None,
    subject: Optional[str] = None,
    since: Optional[str] = None,
    before: Optional[str] = None,
    unread_only: bool = False,
    flagged_only: bool = False,
    has_attachments: Optional[bool] = None,
    limit: int = 25,
    offset: int = 0,
) -> SearchResult:
    """Search Apple Mail messages with flexible filters.

    Results don't include mail_link for performance. Use get_email_link or
    get_email on a specific result to get a clickable message:// URL.

    Args:
        query:           Free-text search applied to the subject line.
        mailbox:         Filter to a specific mailbox by name (e.g. "INBOX", "Sent").
        account:         Filter by account name (e.g. "user@icloud.com").
        from_address:    Substring match on sender name or address.
        to_address:      Substring match on To/CC recipient addresses (e.g. "bill@example.com").
        subject:         Substring match on subject line.
        since:           ISO-8601 date/datetime -- only messages after this time.
        before:          ISO-8601 date/datetime -- only messages before this time.
        unread_only:     If true, return only unread messages.
        flagged_only:    If true, return only flagged messages.
        has_attachments: If true/false, filter on attachment presence.
        limit:           Max results per page (default 25, max 200).
        offset:          Pagination offset.
    """
    bridge = _require_bridge()

    # Parse dates
    since_dt: Optional[datetime] = None
    before_dt: Optional[datetime] = None
    if since:
        try:
            since_dt = datetime.fromisoformat(since.replace("Z", "+00:00"))
        except ValueError:
            raise ValueError(f"Invalid 'since' date: {since!r}. Use ISO-8601 format.")
    if before:
        try:
            before_dt = datetime.fromisoformat(before.replace("Z", "+00:00"))
        except ValueError:
            raise ValueError(f"Invalid 'before' date: {before!r}. Use ISO-8601 format.")

    # The 'query' param is a convenience that searches both subject and sender
    subject_contains: Optional[str] = query
    sender_contains: Optional[str] = query
    # Explicit subject/from_address override query
    if subject:
        subject_contains = subject
    if from_address:
        sender_contains = from_address

    limit = min(limit, 200)

    total, rows = bridge.search_messages(
        mailbox_name=mailbox,
        account_name=account,
        subject_contains=subject_contains,
        sender_contains=sender_contains,
        to_address_contains=to_address,
        since=since_dt,
        before=before_dt,
        is_unread=True if unread_only else None,
        is_flagged=True if flagged_only else None,
        has_attachments=has_attachments,
        limit=limit,
        offset=offset,
    )

    return SearchResult(
        total=total,
        offset=offset,
        limit=limit,
        messages=[_dict_to_summary(r) for r in rows],
    )


@mcp.tool()
def get_email(message_id: int) -> EmailDetail:
    """Fetch a single email with full plain-text body and header details.

    The response includes a mail_link field with a message:// URL that
    opens the email directly in Mail.app when clicked.

    Args:
        message_id: The integer ID from search_emails results.
    """
    bridge = _require_bridge()

    d = bridge.get_message(message_id)
    if d is None:
        raise ValueError(f"Message {message_id} not found.")

    summary = _dict_to_summary(d)

    # Get attachment count
    attachments = bridge.list_attachments(message_id)
    attachment_count = len(attachments)

    return EmailDetail(
        **summary.model_dump(),
        to_addresses=d.get("to_recipients", []),
        cc_addresses=d.get("cc_recipients", []),
        body_text=d.get("body_text"),
        attachment_count=attachment_count,
    )


@mcp.tool()
def get_email_link(message_id: int) -> dict:
    """Get a message:// URL that opens an email directly in Mail.app.

    Lightweight alternative to get_email when you only need the link.

    Args:
        message_id: The integer ID from search_emails results.
    """
    bridge = _require_bridge()
    rfc_id = bridge.get_message_id_header(message_id)
    if rfc_id is None:
        raise ValueError(f"Message {message_id} not found.")
    return {
        "message_id": message_id,
        "mail_link": _make_mail_link(rfc_id),
    }


@mcp.tool()
def get_email_html(message_id: int) -> dict:
    """Get the HTML body of an email.

    Returns a dict with keys:
      - message_id: int
      - has_html: bool
      - html: str or null
      - error: str (only if something went wrong)
    """
    bridge = _require_bridge()

    source = bridge.get_message_source(message_id)
    if not source:
        return {
            "message_id": message_id,
            "has_html": False,
            "html": None,
            "error": "Could not retrieve message source.",
        }

    try:
        msg = email_lib.message_from_string(source)
        html = get_html_body(msg)
        return {
            "message_id": message_id,
            "has_html": html is not None,
            "html": html,
        }
    except Exception as exc:
        return {
            "message_id": message_id,
            "has_html": False,
            "html": None,
            "error": str(exc),
        }


@mcp.tool()
def get_thread(message_id: int) -> list[EmailSummary]:
    """Return all emails in the same conversation thread as the given message.

    Messages are returned in chronological order (oldest first).

    Args:
        message_id: Any email ID in the thread.
    """
    bridge = _require_bridge()
    rows = bridge.get_thread_messages(message_id)
    if not rows:
        # Fall back to returning the single message
        d = bridge.get_message(message_id)
        if d:
            return [_dict_to_summary(d)]
        raise ValueError(f"Message {message_id} not found.")
    return [_dict_to_summary(r) for r in rows]


@mcp.tool()
def list_email_attachments(message_id: int) -> list[Attachment]:
    """List all attachments for a given email.

    Args:
        message_id: Email ID from search_emails results.
    """
    bridge = _require_bridge()
    raw = bridge.list_attachments(message_id)
    return [
        Attachment(
            message_id=message_id,
            index=a["index"],
            filename=a["name"],
            content_type=a["mime_type"],
            size=a["file_size"],
        )
        for a in raw
    ]


@mcp.tool()
def get_email_attachment(message_id: int, attachment_index: int) -> AttachmentData:
    """Retrieve an attachment as base64-encoded data.

    Args:
        message_id:       Email ID from search_emails.
        attachment_index:  The index from list_email_attachments.

    Returns an AttachmentData object with base64-encoded content.
    The caller can decode it with: base64.b64decode(result.data_base64)
    """
    bridge = _require_bridge()
    result = bridge.get_attachment(message_id, attachment_index)
    if result is None:
        raise ValueError(
            f"Attachment index {attachment_index} not found in message {message_id}. "
            "Use list_email_attachments to see available attachments."
        )

    filename, content_type, data = result
    return AttachmentData(
        filename=filename,
        content_type=content_type,
        size=len(data),
        data_base64=base64.b64encode(data).decode("ascii"),
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
