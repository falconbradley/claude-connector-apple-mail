"""Pydantic models for Apple Mail MCP server."""

from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel


class Mailbox(BaseModel):
    name: str               # e.g. "INBOX", "Sent Messages"
    account: str            # e.g. "iCloud", "Gmail"
    full_name: str          # "account/name"
    unread_count: int = 0
    message_count: int = 0


class EmailSummary(BaseModel):
    id: int                 # Mail.app message id
    mailbox: str            # mailbox name
    account: str            # account name
    subject: str
    sender: str             # "Name <email>" or just email address
    date_sent: Optional[datetime] = None
    date_received: Optional[datetime] = None
    is_read: bool = True
    is_flagged: bool = False
    has_attachments: bool = False
    size: int = 0
    message_id: Optional[str] = None   # RFC 2822 Message-ID header
    in_reply_to: Optional[str] = None
    mail_link: Optional[str] = None    # message:// URL to open in Mail.app


class EmailDetail(EmailSummary):
    """Full email with decoded body and recipient lists."""
    to_addresses: list[str] = []
    cc_addresses: list[str] = []
    bcc_addresses: list[str] = []
    body_text: Optional[str] = None
    attachment_count: int = 0


class Attachment(BaseModel):
    message_id: int         # parent email Mail.app id
    index: int              # 0-based index in attachment list
    filename: str
    content_type: str
    size: int               # bytes


class AttachmentData(BaseModel):
    filename: str
    content_type: str
    size: int
    data_base64: str        # base64-encoded file content


class SearchResult(BaseModel):
    total: int
    offset: int
    limit: int
    messages: list[EmailSummary]


class MailboxStats(BaseModel):
    total_messages: int
    unread_messages: int
    mailbox_count: int
    account_count: int


class DraftResult(BaseModel):
    subject: str
    to_addresses: list[str]
    cc_addresses: list[str] = []
    bcc_addresses: list[str] = []
    draft_link: Optional[str] = None  # message:// URL to open draft in Mail.app
