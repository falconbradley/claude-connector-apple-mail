"""
Unit tests for the fast Envelope Index read path.

Builds a synthetic ~/Library/Mail replica (SQLite Envelope Index +
.emlx files) in a temp dir, then exercises EnvelopeIndexBridge against
it. Two schema variants are covered:

  - "modern": dedicated read/flagged/deleted columns, recipients.message
  - "legacy": flags bitfield only, recipients.message_id

Run:  uv run --with pytest pytest tests/test_envelope.py -q
"""

from __future__ import annotations

import sqlite3
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

from apple_mail_mcp.envelope import EnvelopeIndexBridge, EnvelopeUnavailable

# Fixed reference time so tests are deterministic
NOW = datetime(2026, 8, 1, 12, 0, 0, tzinfo=timezone.utc)


def ts(days_ago: float) -> float:
    return (NOW - timedelta(days=days_ago)).timestamp()


def write_emlx(path: Path, message: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    plist = b'<?xml version="1.0"?><plist version="1.0"><dict/></plist>\n'
    path.write_bytes(str(len(message)).encode() + b"\n" + message + plist)


def simple_message(
    subject: str,
    sender: str = "Alice Example <alice@example.com>",
    to: str = "brad@icloud.com",
    cc: str = "",
    message_id: str = "<msg-1@example.com>",
    body: str = "Hello Brad,\n\nThis is the body.\n",
) -> bytes:
    cc_header = f"Cc: {cc}\r\n" if cc else ""
    return (
        f"From: {sender}\r\n"
        f"To: {to}\r\n"
        f"{cc_header}"
        f"Subject: {subject}\r\n"
        f"Message-ID: {message_id}\r\n"
        f"Date: Sat, 01 Aug 2026 12:00:00 +0000\r\n"
        f"Content-Type: text/plain; charset=utf-8\r\n"
        f"\r\n"
        f"{body}"
    ).encode()


MULTIPART_MESSAGE = (
    b"From: Carol <carol@example.com>\r\n"
    b"To: brad@icloud.com\r\n"
    b"Subject: Report attached\r\n"
    b"Message-ID: <msg-attach@example.com>\r\n"
    b"MIME-Version: 1.0\r\n"
    b'Content-Type: multipart/mixed; boundary="BOUND"\r\n'
    b"\r\n"
    b"--BOUND\r\n"
    b"Content-Type: text/html; charset=utf-8\r\n"
    b"\r\n"
    b"<html><body><p>See the <b>attached</b> report.</p></body></html>\r\n"
    b"--BOUND\r\n"
    b"Content-Type: application/pdf; name=report.pdf\r\n"
    b"Content-Disposition: attachment; filename=report.pdf\r\n"
    b"Content-Transfer-Encoding: base64\r\n"
    b"\r\n"
    b"JVBERi0xLjQgZmFrZQ==\r\n"
    b"--BOUND--\r\n"
)


def build_store(root: Path, modern: bool = True) -> None:
    """Create V10/MailData/Envelope Index + .emlx files under root."""
    maildata = root / "V10" / "MailData"
    maildata.mkdir(parents=True)
    db = sqlite3.connect(maildata / "Envelope Index")

    rec_msg_col = "message" if modern else "message_id"
    extra_msg_cols = (
        "read INTEGER DEFAULT 0, flagged INTEGER DEFAULT 0, "
        "deleted INTEGER DEFAULT 0, "
        if modern
        else ""
    )
    db.executescript(
        f"""
        CREATE TABLE mailboxes (ROWID INTEGER PRIMARY KEY, url TEXT);
        CREATE TABLE subjects (ROWID INTEGER PRIMARY KEY, subject TEXT,
                               normalized_subject TEXT);
        CREATE TABLE addresses (ROWID INTEGER PRIMARY KEY, address TEXT,
                                comment TEXT);
        CREATE TABLE recipients (ROWID INTEGER PRIMARY KEY,
                                 {rec_msg_col} INTEGER, type INTEGER,
                                 address INTEGER, position INTEGER);
        CREATE TABLE attachments (ROWID INTEGER PRIMARY KEY,
                                  {rec_msg_col} INTEGER, name TEXT);
        CREATE TABLE messages (
            ROWID INTEGER PRIMARY KEY,
            message_id INTEGER,
            in_reply_to INTEGER,
            sender INTEGER,
            subject_prefix TEXT,
            subject INTEGER,
            date_sent REAL,
            date_received REAL,
            mailbox INTEGER,
            flags INTEGER DEFAULT 0,
            {extra_msg_cols}
            size INTEGER DEFAULT 0,
            conversation_id INTEGER
        );
        """
    )

    db.executemany(
        "INSERT INTO mailboxes (ROWID, url) VALUES (?, ?)",
        [
            (1, "imap://brad%40icloud.com@p58-imap.mail.me.com/INBOX"),
            (2, "imap://brad%40icloud.com@p58-imap.mail.me.com/Sent%20Messages"),
            (3, "local:///Drafts"),
        ],
    )
    db.executemany(
        "INSERT INTO subjects (ROWID, subject, normalized_subject) VALUES (?, ?, ?)",
        [
            (1, "Quarterly invoice", "quarterly invoice"),
            (2, "Lunch on Friday?", "lunch on friday?"),
            (3, "Report attached", "report attached"),
        ],
    )
    db.executemany(
        "INSERT INTO addresses (ROWID, address, comment) VALUES (?, ?, ?)",
        [
            (1, "alice@example.com", "Alice Example"),
            (2, "bob@example.com", ""),
            (3, "carol@example.com", "Carol"),
            (4, "brad@icloud.com", "Brad"),
            (5, "bill@partners.com", "Bill Partner"),
        ],
    )

    def flags(read: bool, flagged: bool = False, deleted: bool = False,
              attach: int = 0) -> int:
        value = 0
        if read:
            value |= 1
        if deleted:
            value |= 1 << 1
        if flagged:
            value |= 1 << 4
        value |= (attach & 0x3F) << 10
        return value

    read_cols = ", read, flagged, deleted" if modern else ""

    def msg_row(rowid, sender, prefix, subject, days_ago, mailbox,
                read, flagged=False, deleted=False, attach=0, conv=None):
        base = [rowid, sender, prefix, subject, ts(days_ago), ts(days_ago),
                mailbox, flags(read, flagged, deleted, attach), 2048,
                conv if conv is not None else rowid * 100]
        if modern:
            base += [1 if read else 0, 1 if flagged else 0, 1 if deleted else 0]
        return base

    placeholders = ", ".join("?" * (13 if modern else 10))
    db.executemany(
        f"""INSERT INTO messages
            (ROWID, sender, subject_prefix, subject, date_sent, date_received,
             mailbox, flags, size, conversation_id{read_cols})
            VALUES ({placeholders})""",
        [
            # 101: unread invoice from Alice, 1 day old, in INBOX, thread 900
            msg_row(101, 1, "", 1, 1, 1, read=False, conv=900),
            # 102: read reply "Re: Quarterly invoice" from Bob, same thread
            msg_row(102, 2, "Re: ", 1, 0.5, 1, read=True, conv=900),
            # 103: flagged lunch mail from Carol, 10 days old
            msg_row(103, 3, "", 2, 10, 1, read=True, flagged=True),
            # 104: sent mail (Sent Messages mailbox)
            msg_row(104, 4, "", 2, 3, 2, read=True),
            # 105: deleted message — must never appear
            msg_row(105, 1, "", 1, 2, 1, read=True, deleted=True),
            # 106: message with attachment
            msg_row(106, 3, "", 3, 0.2, 1, read=False, attach=1),
        ],
    )

    db.executemany(
        f"INSERT INTO recipients ({rec_msg_col}, type, address, position) "
        "VALUES (?, ?, ?, ?)",
        [
            (101, 0, 4, 0),   # to brad
            (101, 1, 5, 1),   # cc bill
            (103, 0, 4, 0),
            (104, 0, 5, 0),   # sent to bill
            (106, 0, 4, 0),
        ],
    )
    db.execute(
        f"INSERT INTO attachments ({rec_msg_col}, name) VALUES (?, ?)",
        (106, "report.pdf"),
    )
    db.commit()
    db.close()

    # .emlx files (nested like the real store: .../<mbox>/<uuid>/Data/.../Messages)
    messages_dir = (
        root / "V10" / "ACCT-UUID" / "INBOX.mbox" / "MBOX-UUID" / "Data"
        / "1" / "Messages"
    )
    write_emlx(
        messages_dir / "101.emlx",
        simple_message(
            "Quarterly invoice",
            to="Brad <brad@icloud.com>",
            cc="Bill Partner <bill@partners.com>",
            message_id="<invoice-101@example.com>",
            body="Please find the invoice.\n",
        ),
    )
    write_emlx(
        messages_dir / "102.emlx",
        simple_message(
            "Re: Quarterly invoice",
            sender="bob@example.com",
            message_id="<reply-102@example.com>",
        ),
    )
    write_emlx(messages_dir / "106.partial.emlx", MULTIPART_MESSAGE)


@pytest.fixture(params=["modern", "legacy"])
def bridge(request, tmp_path):
    build_store(tmp_path, modern=(request.param == "modern"))
    return EnvelopeIndexBridge(mail_root=tmp_path)


def test_unavailable_when_missing(tmp_path):
    with pytest.raises(EnvelopeUnavailable):
        EnvelopeIndexBridge(mail_root=tmp_path / "nope")


def test_stats(bridge):
    stats = bridge.get_stats()
    assert stats["total_messages"] == 5  # deleted 105 excluded
    assert stats["unread_messages"] == 2
    assert stats["account_count"] == 2  # icloud + On My Mac


def test_list_mailboxes(bridge):
    boxes = {(b["account_name"], b["name"]): b for b in bridge.list_mailboxes()}
    inbox = boxes[("brad@icloud.com", "INBOX")]
    assert inbox["message_count"] == 4
    assert inbox["unread_count"] == 2
    assert boxes[("brad@icloud.com", "Sent Messages")]["message_count"] == 1
    assert ("On My Mac", "Drafts") in boxes


def test_search_recent_sorted(bridge):
    total, rows = bridge.search_messages(limit=10)
    assert total == 5
    assert [r["id"] for r in rows] == [106, 102, 101, 104, 103]
    assert rows[0]["has_attachments"] is True
    assert rows[0]["mailbox_name"] == "INBOX"
    assert rows[0]["account_name"] == "brad@icloud.com"


def test_search_query_mode_subject_or_sender(bridge):
    # "query" convenience: subject OR sender
    total, rows = bridge.search_messages(
        subject_contains="invoice", sender_contains="invoice", limit=10
    )
    assert {r["id"] for r in rows} == {101, 102}
    # sender-side match of query mode
    total, rows = bridge.search_messages(
        subject_contains="carol", sender_contains="carol", limit=10
    )
    assert {r["id"] for r in rows} == {103, 106}


def test_search_filters_and_pagination(bridge):
    total, rows = bridge.search_messages(is_unread=True, limit=10)
    assert {r["id"] for r in rows} == {101, 106}

    total, rows = bridge.search_messages(is_flagged=True, limit=10)
    assert [r["id"] for r in rows] == [103]
    assert rows[0]["is_flagged"] is True

    total, rows = bridge.search_messages(sender_contains="alice", limit=10)
    assert [r["id"] for r in rows] == [101]
    assert rows[0]["sender"] == "Alice Example <alice@example.com>"

    total, rows = bridge.search_messages(to_address_contains="bill@partners", limit=10)
    assert {r["id"] for r in rows} == {101, 104}

    total, rows = bridge.search_messages(mailbox_name="sent", limit=10)
    assert [r["id"] for r in rows] == [104]

    total, rows = bridge.search_messages(account_name="icloud", limit=10)
    assert total == 5

    total, rows = bridge.search_messages(since=NOW - timedelta(days=2), limit=10)
    assert {r["id"] for r in rows} == {101, 102, 106}

    total, rows = bridge.search_messages(before=NOW - timedelta(days=2), limit=10)
    assert {r["id"] for r in rows} == {103, 104}

    total, rows = bridge.search_messages(has_attachments=True, limit=10)
    assert [r["id"] for r in rows] == [106]

    total, page1 = bridge.search_messages(limit=2, offset=0)
    total, page2 = bridge.search_messages(limit=2, offset=2)
    assert total == 5
    assert [r["id"] for r in page1] == [106, 102]
    assert [r["id"] for r in page2] == [101, 104]


def test_search_like_escaping(bridge):
    total, rows = bridge.search_messages(subject_contains="100%", limit=10)
    assert total == 0


def test_get_message_with_body(bridge):
    d = bridge.get_message(101)
    assert d is not None
    assert d["subject"] == "Quarterly invoice"
    assert "Please find the invoice." in d["body_text"]
    assert d["message_id"] == "invoice-101@example.com"
    assert d["to_recipients"] == ["Brad <brad@icloud.com>"]
    assert d["cc_recipients"] == ["Bill Partner <bill@partners.com>"]
    assert d["is_read"] is False


def test_get_message_html_only_body(bridge):
    d = bridge.get_message(106)
    assert d is not None
    # HTML-only message: body derived from HTML part
    assert "attached" in d["body_text"]
    assert "<b>" not in d["body_text"]


def test_get_message_missing_emlx_falls_back_to_db_recipients(bridge):
    d = bridge.get_message(104)  # no .emlx written for this one
    assert d is not None
    assert d["body_text"] == ""
    assert d["to_recipients"] == ["Bill Partner <bill@partners.com>"]


def test_get_message_not_found(bridge):
    assert bridge.get_message(99999) is None


def test_subject_prefix_concatenation(bridge):
    d = bridge.get_message(102)
    assert d["subject"] == "Re: Quarterly invoice"


def test_thread_via_conversation_id(bridge):
    rows = bridge.get_thread_messages(101)
    assert [r["id"] for r in rows] == [101, 102]  # oldest first
    rows = bridge.get_thread_messages(102)
    assert [r["id"] for r in rows] == [101, 102]


def test_message_source_and_id_header(bridge):
    src = bridge.get_message_source(101)
    assert src is not None and "Message-ID: <invoice-101@example.com>" in src
    assert bridge.get_message_id_header(101) == "invoice-101@example.com"
    assert bridge.get_message_id_header(104) is None


def test_attachments(bridge):
    atts = bridge.list_attachments(106)
    assert len(atts) == 1
    assert atts[0]["name"] == "report.pdf"
    assert atts[0]["mime_type"] == "application/pdf"

    result = bridge.get_attachment(106, 0)
    assert result is not None
    name, mime, data = result
    assert name == "report.pdf"
    assert data.startswith(b"%PDF")

    assert bridge.get_attachment(106, 5) is None


def test_get_flag(bridge):
    assert bridge.get_flag(103)["is_flagged"] is True
    assert bridge.get_flag(101)["is_flagged"] is False
    with pytest.raises(ValueError):
        bridge.get_flag(99999)


def test_external_attachment_from_partial_emlx(bridge, tmp_path):
    # Strip the payload out of the partial emlx and park the body on disk
    # the way Mail does for .partial.emlx messages.
    stub = MULTIPART_MESSAGE.replace(b"JVBERi0xLjQgZmFrZQ==\r\n", b"")
    messages_dir = (
        tmp_path / "V10" / "ACCT-UUID" / "INBOX.mbox" / "MBOX-UUID" / "Data"
        / "1" / "Messages"
    )
    write_emlx(messages_dir / "106.partial.emlx", stub)
    external = (
        messages_dir.parent / "Attachments" / "106" / "2" / "report.pdf"
    )
    external.parent.mkdir(parents=True)
    external.write_bytes(b"%PDF-1.4 external")

    # Force re-read (cache maps id->path already; content is re-parsed per call)
    result = bridge.get_attachment(106, 0)
    assert result is not None
    assert result[2] == b"%PDF-1.4 external"
