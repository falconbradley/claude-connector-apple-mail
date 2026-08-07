"""
Fast read path: Mail.app's local message store.

Mail.app continuously mirrors every account (including iCloud) to disk:

  ~/Library/Mail/V*/MailData/Envelope Index   - SQLite index of all messages
  ~/Library/Mail/V*/.../Messages/<id>.emlx    - raw RFC 2822 message files

Querying the SQLite index takes milliseconds where the AppleScript/JXA
bridge takes tens of seconds, because it avoids Apple Events entirely.
Mail.app remains the sync + auth engine; this module is a read-only
consumer of the data it has already synced.

Requires Full Disk Access for the host process (System Settings ->
Privacy & Security -> Full Disk Access). When FDA is missing this module
raises EnvelopeUnavailable and callers fall back to the JXA bridge.

Schema notes: the Envelope Index schema varies across macOS releases, so
column usage is decided by runtime introspection (PRAGMA table_info) with
documented fallbacks (e.g. read/flagged/deleted bits in the `flags`
bitfield when the dedicated columns are absent).
"""

from __future__ import annotations

import logging
import os
import re
import sqlite3
import threading
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote, unquote, urlparse

from .emlx import (
    get_html_body,
    get_text_body,
    html_to_text,
    read_emlx,
    read_emlx_message_bytes,
)

logger = logging.getLogger("apple_mail_mcp.envelope")

# Seconds between Unix epoch (1970) and Mac absolute time epoch (2001).
_MAC_EPOCH_OFFSET = 978307200

# Bit positions in the messages.flags bitfield (Apple Mail forensics,
# stable since 10.x). Used only when the dedicated column is absent.
_FLAG_BIT_READ = 1 << 0
_FLAG_BIT_DELETED = 1 << 1
_FLAG_BIT_FLAGGED = 1 << 4
_FLAG_ATTACH_SHIFT = 10  # bits 10-15: attachment count
_FLAG_ATTACH_MASK = 0x3F

# Minimum interval between full re-walks of the Mail dir for .emlx lookup.
_EMLX_RESCAN_INTERVAL = 30.0


class EnvelopeUnavailable(RuntimeError):
    """The Envelope Index cannot be used (no FDA, missing, schema drift)."""


def _like_pattern(term: str) -> str:
    """Build a case-insensitive LIKE pattern with escaped wildcards."""
    escaped = (
        term.lower().replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    )
    return f"%{escaped}%"


class EnvelopeIndexBridge:
    """Read-only bridge to Mail.app's Envelope Index + .emlx store.

    Mirrors the read API of applescript.MailBridge so the two are
    interchangeable from the server's point of view. Message ids are the
    Envelope Index ROWIDs, which are the same integer ids Mail.app exposes
    through its scripting interface (and the .emlx filenames), so ids from
    this bridge can be handed to the JXA bridge for write operations.
    """

    def __init__(self, mail_root: Optional[Path] = None) -> None:
        root = mail_root or Path(
            os.environ.get("APPLE_MAIL_MCP_MAIL_ROOT", Path.home() / "Library" / "Mail")
        )
        self.mail_root = Path(root)
        self.version_dir = self._locate_version_dir(self.mail_root)
        self.db_path = self.version_dir / "MailData" / "Envelope Index"
        if not self._exists(self.db_path):
            raise EnvelopeUnavailable(
                f"Envelope Index not found at {self.db_path}."
            )

        self._lock = threading.Lock()
        self._conn = self._connect(self.db_path)
        self._introspect_schema()
        self._detect_epoch()

        # mailbox ROWID -> (account_display, mailbox_name, url)
        self._mailboxes: dict[int, tuple[str, str, str]] = {}
        self._load_mailboxes()

        # message ROWID -> emlx path cache
        self._emlx_index: dict[int, Path] = {}
        self._emlx_last_scan = 0.0

        logger.info(
            "EnvelopeIndexBridge ready: %s (%d mailboxes, epoch=%s)",
            self.db_path,
            len(self._mailboxes),
            "unix" if self._epoch_offset == 0 else "mac-absolute",
        )

    # ------------------------------------------------------------------
    # Setup helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _exists(path: Path) -> bool:
        try:
            return path.exists()
        except OSError:
            return False

    @staticmethod
    def _locate_version_dir(root: Path) -> Path:
        try:
            candidates = [
                p for p in root.iterdir()
                if p.is_dir() and re.fullmatch(r"V\d+", p.name)
            ]
        except PermissionError as exc:
            raise EnvelopeUnavailable(
                f"Cannot read {root} — the host process needs Full Disk Access "
                "(System Settings -> Privacy & Security -> Full Disk Access). "
                f"Underlying error: {exc}"
            )
        except FileNotFoundError:
            raise EnvelopeUnavailable(f"Mail data directory not found: {root}")
        candidates = [
            p for p in candidates
            if (p / "MailData" / "Envelope Index").exists()
        ]
        if not candidates:
            raise EnvelopeUnavailable(
                f"No V*/MailData/Envelope Index found under {root}."
            )
        return max(candidates, key=lambda p: int(p.name[1:]))

    @staticmethod
    def _connect(db_path: Path) -> sqlite3.Connection:
        """Open the live Envelope Index without ever writing to it.

        Try a strict read-only URI first; that can fail on a WAL database
        when Mail.app is not holding the shared-memory index, so fall back
        to a normal open hardened with PRAGMA query_only.
        """
        uri = f"file:{quote(str(db_path))}?mode=ro"
        for attempt, kwargs in enumerate(
            ({"database": uri, "uri": True}, {"database": str(db_path)})
        ):
            try:
                conn = sqlite3.connect(
                    timeout=5.0, check_same_thread=False, **kwargs
                )
                conn.execute("PRAGMA query_only = ON")
                conn.execute("SELECT COUNT(*) FROM sqlite_master")
                return conn
            except sqlite3.Error as exc:
                last_exc = exc
                logger.debug("Envelope Index open attempt %d failed: %s", attempt, exc)
        raise EnvelopeUnavailable(
            f"Could not open Envelope Index at {db_path}: {last_exc}"
        )

    def _introspect_schema(self) -> None:
        tables = {
            row[0]
            for row in self._conn.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            )
        }
        self._cols: dict[str, set[str]] = {}
        for table in ("messages", "subjects", "addresses", "recipients",
                      "mailboxes", "attachments"):
            if table in tables:
                self._cols[table] = {
                    row[1]
                    for row in self._conn.execute(f"PRAGMA table_info({table})")
                }

        mcols = self._cols.get("messages", set())
        required = {"subject", "sender", "mailbox", "date_received"}
        if not required <= mcols:
            raise EnvelopeUnavailable(
                "Envelope Index schema not recognised — messages table is "
                f"missing {sorted(required - mcols)}. "
                f"Present columns: {sorted(mcols)}"
            )
        for table in ("subjects", "addresses", "mailboxes"):
            if table not in self._cols:
                raise EnvelopeUnavailable(
                    f"Envelope Index schema not recognised — no {table} table."
                )

        self._read_expr = (
            "m.read" if "read" in mcols else f"(m.flags & {_FLAG_BIT_READ})"
        )
        self._flagged_expr = (
            "m.flagged" if "flagged" in mcols
            else f"((m.flags & {_FLAG_BIT_FLAGGED}) != 0)"
        )
        self._not_deleted_expr = (
            "m.deleted = 0" if "deleted" in mcols
            else f"(m.flags & {_FLAG_BIT_DELETED}) = 0"
        )
        self._size_expr = "m.size" if "size" in mcols else "0"
        self._date_sent_expr = "m.date_sent" if "date_sent" in mcols else "NULL"
        self._subject_prefix_expr = (
            "COALESCE(m.subject_prefix, '')" if "subject_prefix" in mcols else "''"
        )
        self._has_conversation = "conversation_id" in mcols
        self._flag_color_col = next(
            (c for c in ("flag_color",) if c in mcols), None
        )

        # recipients / attachments linkage columns vary across versions
        rcols = self._cols.get("recipients", set())
        self._rec_msg_col = (
            "message" if "message" in rcols
            else "message_id" if "message_id" in rcols
            else None
        )
        self._rec_type_col = "type" if "type" in rcols else None
        acols = self._cols.get("attachments", set())
        self._att_msg_col = (
            "message" if "message" in acols
            else "message_id" if "message_id" in acols
            else None
        )

        subj_cols = self._cols.get("subjects", set())
        self._normalized_subject_col = (
            "normalized_subject" if "normalized_subject" in subj_cols else None
        )

        if "flags" not in mcols:
            # Every known schema has flags; fall back gracefully anyway.
            self._read_expr = "m.read" if "read" in mcols else "1"
            self._flagged_expr = "m.flagged" if "flagged" in mcols else "0"
            self._not_deleted_expr = (
                "m.deleted = 0" if "deleted" in mcols else "1 = 1"
            )

        self._has_attachments_expr = self._build_has_attachments_expr(mcols)

    def _build_has_attachments_expr(self, mcols: set[str]) -> str:
        if self._att_msg_col:
            return (
                "EXISTS (SELECT 1 FROM attachments att "
                f"WHERE att.{self._att_msg_col} = m.ROWID)"
            )
        if "flags" in mcols:
            return (
                f"(((m.flags >> {_FLAG_ATTACH_SHIFT}) & {_FLAG_ATTACH_MASK}) > 0)"
            )
        return "0"

    def _detect_epoch(self) -> None:
        """Envelope dates are Unix epoch on modern macOS; older builds used
        Mac absolute time (seconds since 2001). Decide from the data."""
        row = self._conn.execute("SELECT MAX(date_received) FROM messages").fetchone()
        max_date = row[0] or 0
        # Any post-2001 mailbox has Unix values > 1e9; Mac-absolute values
        # for the same era are ~7.6e8 or lower.
        self._epoch_offset = 0 if max_date > 1.0e9 else _MAC_EPOCH_OFFSET

    def _load_mailboxes(self) -> None:
        rows = self._query("SELECT ROWID, url FROM mailboxes")
        mapping: dict[int, tuple[str, str, str]] = {}
        for rowid, url in rows:
            mapping[rowid] = self._parse_mailbox_url(url or "")
        self._mailboxes = mapping

    @staticmethod
    def _parse_mailbox_url(url: str) -> tuple[str, str, str]:
        """Return (account_display, mailbox_name, url).

        Typical urls:
          imap://brad%40icloud.com@p58-imap.mail.me.com/INBOX
          imap://.../Sent%20Messages
          local:///Drafts   (On My Mac)
        """
        try:
            parsed = urlparse(url)
        except ValueError:
            return ("", url, url)
        name = unquote(parsed.path.strip("/").split("/")[-1]) if parsed.path else url
        if parsed.scheme in ("local", "file") or not parsed.netloc:
            account = "On My Mac"
        elif "@" in parsed.netloc:
            account = unquote(parsed.netloc.rsplit("@", 1)[0])
        else:
            account = unquote(parsed.netloc)
        return (account, name or url, url)

    # ------------------------------------------------------------------
    # Query plumbing
    # ------------------------------------------------------------------

    def _query(self, sql: str, params: tuple = ()) -> list[tuple]:
        with self._lock:
            try:
                return list(self._conn.execute(sql, params))
            except sqlite3.OperationalError as exc:
                # WAL handoff hiccups are transient; retry once on a
                # fresh connection before giving up.
                logger.warning("Envelope query failed (%s); reconnecting.", exc)
                try:
                    self._conn.close()
                except sqlite3.Error:
                    pass
                self._conn = self._connect(self.db_path)
                return list(self._conn.execute(sql, params))

    def _mailbox_info(self, mailbox_rowid: int) -> tuple[str, str, str]:
        info = self._mailboxes.get(mailbox_rowid)
        if info is None:
            self._load_mailboxes()
            info = self._mailboxes.get(mailbox_rowid, ("", "", ""))
        return info

    def _to_iso(self, value: Any) -> Optional[str]:
        if not value:
            return None
        try:
            ts = float(value) + self._epoch_offset
            return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()
        except (ValueError, OverflowError, OSError):
            return None

    def _to_db_time(self, dt: datetime) -> float:
        return dt.timestamp() - self._epoch_offset

    # ------------------------------------------------------------------
    # .emlx location
    # ------------------------------------------------------------------

    def _scan_emlx(self) -> None:
        t0 = time.monotonic()
        index: dict[int, Path] = {}
        for dirpath, dirnames, filenames in os.walk(self.version_dir):
            for fname in filenames:
                if not fname.endswith(".emlx"):
                    continue
                stem = fname.split(".", 1)[0]
                if stem.isdigit():
                    rowid = int(stem)
                    path = Path(dirpath) / fname
                    # Prefer full .emlx over .partial.emlx when both exist
                    existing = index.get(rowid)
                    if existing is None or existing.name.endswith(".partial.emlx"):
                        index[rowid] = path
        self._emlx_index = index
        self._emlx_last_scan = time.monotonic()
        logger.info(
            "Indexed %d .emlx files in %.2fs", len(index), time.monotonic() - t0
        )

    def _find_emlx(self, message_rowid: int) -> Optional[Path]:
        path = self._emlx_index.get(message_rowid)
        if path is not None and self._exists(path):
            return path
        if time.monotonic() - self._emlx_last_scan > _EMLX_RESCAN_INTERVAL:
            self._scan_emlx()
            path = self._emlx_index.get(message_rowid)
            if path is not None and self._exists(path):
                return path
        return None

    def _read_message_file(self, message_rowid: int):
        path = self._find_emlx(message_rowid)
        if path is None:
            return None
        try:
            return read_emlx(path)
        except (OSError, ValueError) as exc:
            logger.warning("Failed to parse %s: %s", path, exc)
            return None

    # ------------------------------------------------------------------
    # Public API (mirrors MailBridge reads)
    # ------------------------------------------------------------------

    def get_stats(self) -> dict:
        rows = self._query(
            f"""
            SELECT COUNT(*),
                   SUM(CASE WHEN {self._read_expr} = 0 THEN 1 ELSE 0 END),
                   COUNT(DISTINCT m.mailbox)
            FROM messages m
            WHERE {self._not_deleted_expr}
            """
        )
        total, unread, mailbox_count = rows[0] if rows else (0, 0, 0)
        accounts = {info[0] for info in self._mailboxes.values() if info[0]}
        return {
            "total_messages": total or 0,
            "unread_messages": unread or 0,
            "mailbox_count": mailbox_count or 0,
            "account_count": len(accounts),
        }

    def list_mailboxes(self) -> list[dict]:
        counts: dict[int, tuple[int, int]] = {}
        for mbox, total, unread in self._query(
            f"""
            SELECT m.mailbox, COUNT(*),
                   SUM(CASE WHEN {self._read_expr} = 0 THEN 1 ELSE 0 END)
            FROM messages m
            WHERE {self._not_deleted_expr}
            GROUP BY m.mailbox
            """
        ):
            counts[mbox] = (total or 0, unread or 0)

        result = []
        for rowid, (account, name, _url) in sorted(
            self._mailboxes.items(), key=lambda kv: (kv[1][0], kv[1][1])
        ):
            total, unread = counts.get(rowid, (0, 0))
            result.append(
                {
                    "name": name,
                    "account_name": account,
                    "unread_count": unread,
                    "message_count": total,
                }
            )
        return result

    def search_messages(
        self,
        *,
        mailbox_name: Optional[str] = None,
        account_name: Optional[str] = None,
        subject_contains: Optional[str] = None,
        sender_contains: Optional[str] = None,
        to_address_contains: Optional[str] = None,
        since: Optional[datetime] = None,
        before: Optional[datetime] = None,
        is_unread: Optional[bool] = None,
        is_flagged: Optional[bool] = None,
        has_attachments: Optional[bool] = None,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[int, list[dict]]:
        t0 = time.monotonic()
        where: list[str] = [self._not_deleted_expr]
        params: list[Any] = []

        subject_expr = f"lower({self._subject_prefix_expr} || COALESCE(s.subject, ''))"
        sender_expr = (
            "lower(COALESCE(a.comment, '') || ' <' || COALESCE(a.address, '') || '>')"
        )

        # Mailbox / account filters resolve to rowid sets in Python
        if mailbox_name or account_name:
            wanted = [
                rowid
                for rowid, (account, name, url) in self._mailboxes.items()
                if (
                    not mailbox_name
                    or mailbox_name.lower() in name.lower()
                )
                and (
                    not account_name
                    or account_name.lower() in account.lower()
                    or account_name.lower() in url.lower()
                )
            ]
            if not wanted:
                return 0, []
            where.append(
                f"m.mailbox IN ({','.join(str(r) for r in wanted)})"
            )

        query_mode = (
            subject_contains
            and sender_contains
            and subject_contains == sender_contains
        )
        if query_mode:
            pattern = _like_pattern(subject_contains)
            where.append(
                f"({subject_expr} LIKE ? ESCAPE '\\' "
                f"OR {sender_expr} LIKE ? ESCAPE '\\')"
            )
            params += [pattern, pattern]
        else:
            if subject_contains:
                where.append(f"{subject_expr} LIKE ? ESCAPE '\\'")
                params.append(_like_pattern(subject_contains))
            if sender_contains:
                where.append(f"{sender_expr} LIKE ? ESCAPE '\\'")
                params.append(_like_pattern(sender_contains))

        if to_address_contains and self._rec_msg_col:
            pattern = _like_pattern(to_address_contains)
            type_filter = (
                f"AND r.{self._rec_type_col} IN (0, 1) " if self._rec_type_col else ""
            )
            where.append(
                "EXISTS (SELECT 1 FROM recipients r "
                "LEFT JOIN addresses ra ON ra.ROWID = r.address "
                f"WHERE r.{self._rec_msg_col} = m.ROWID {type_filter}"
                "AND (lower(COALESCE(ra.address, '')) LIKE ? ESCAPE '\\' "
                "OR lower(COALESCE(ra.comment, '')) LIKE ? ESCAPE '\\'))"
            )
            params += [pattern, pattern]

        if since is not None:
            where.append("m.date_received >= ?")
            params.append(self._to_db_time(since))
        if before is not None:
            where.append("m.date_received <= ?")
            params.append(self._to_db_time(before))

        if is_unread is True:
            where.append(f"{self._read_expr} = 0")
        elif is_unread is False:
            where.append(f"{self._read_expr} != 0")

        if is_flagged is True:
            where.append(f"{self._flagged_expr} != 0")
        elif is_flagged is False:
            where.append(f"{self._flagged_expr} = 0")

        if has_attachments is True:
            where.append(self._has_attachments_expr)
        elif has_attachments is False:
            where.append(f"NOT ({self._has_attachments_expr})")

        where_sql = " AND ".join(where)
        from_sql = (
            "FROM messages m "
            "LEFT JOIN subjects s ON m.subject = s.ROWID "
            "LEFT JOIN addresses a ON m.sender = a.ROWID"
        )

        total = self._query(
            f"SELECT COUNT(*) {from_sql} WHERE {where_sql}", tuple(params)
        )[0][0]

        rows = self._query(
            f"""
            SELECT m.ROWID,
                   {self._subject_prefix_expr} || COALESCE(s.subject, ''),
                   CASE WHEN COALESCE(a.comment, '') != ''
                        THEN a.comment || ' <' || COALESCE(a.address, '') || '>'
                        ELSE COALESCE(a.address, '') END,
                   m.date_received,
                   {self._date_sent_expr},
                   {self._read_expr},
                   {self._flagged_expr},
                   {self._has_attachments_expr},
                   m.mailbox,
                   {self._size_expr}
            {from_sql}
            WHERE {where_sql}
            ORDER BY (m.date_received IS NULL OR m.date_received = 0),
                     m.date_received DESC
            LIMIT ? OFFSET ?
            """,
            tuple(params) + (limit, offset),
        )

        results = []
        for (
            rowid, subject, sender, date_received, date_sent,
            read, flagged, has_att, mailbox_rowid, size,
        ) in rows:
            account, mbox_name, _url = self._mailbox_info(mailbox_rowid)
            results.append(
                {
                    "id": rowid,
                    "subject": subject or "",
                    "sender": sender or "",
                    "date_received": self._to_iso(date_received),
                    "date_sent": self._to_iso(date_sent) or self._to_iso(date_received),
                    "is_read": bool(read),
                    "is_flagged": bool(flagged),
                    "has_attachments": bool(has_att),
                    "mailbox_name": mbox_name,
                    "account_name": account,
                    "message_id": None,
                    "in_reply_to": None,
                    "size": size or 0,
                }
            )

        logger.info(
            "Envelope search: %d total, %d returned in %.0fms",
            total, len(results), (time.monotonic() - t0) * 1000,
        )
        return total, results

    def get_message(self, message_id: int) -> Optional[dict]:
        rows = self._query(
            f"""
            SELECT m.ROWID,
                   {self._subject_prefix_expr} || COALESCE(s.subject, ''),
                   CASE WHEN COALESCE(a.comment, '') != ''
                        THEN a.comment || ' <' || COALESCE(a.address, '') || '>'
                        ELSE COALESCE(a.address, '') END,
                   m.date_received,
                   {self._date_sent_expr},
                   {self._read_expr},
                   {self._flagged_expr},
                   {self._has_attachments_expr},
                   m.mailbox,
                   {self._size_expr}
            FROM messages m
            LEFT JOIN subjects s ON m.subject = s.ROWID
            LEFT JOIN addresses a ON m.sender = a.ROWID
            WHERE m.ROWID = ?
            """,
            (message_id,),
        )
        if not rows:
            return None
        (
            rowid, subject, sender, date_received, date_sent,
            read, flagged, has_att, mailbox_rowid, size,
        ) = rows[0]
        account, mbox_name, _url = self._mailbox_info(mailbox_rowid)

        body_text = ""
        rfc_message_id: Optional[str] = None
        in_reply_to: Optional[str] = None
        to_recipients: list[str] = []
        cc_recipients: list[str] = []

        msg = self._read_message_file(rowid)
        if msg is not None:
            body_text = get_text_body(msg)
            if not body_text:
                html = get_html_body(msg)
                if html:
                    body_text = html_to_text(html)
            rfc_message_id = (msg.get("Message-ID") or "").strip().strip("<>") or None
            in_reply_to = (msg.get("In-Reply-To") or "").strip().strip("<>") or None
            to_recipients = self._header_addresses(msg, "To")
            cc_recipients = self._header_addresses(msg, "Cc")
        else:
            to_recipients, cc_recipients = self._recipients_from_db(rowid)

        return {
            "id": rowid,
            "subject": subject or "",
            "sender": sender or "",
            "date_received": self._to_iso(date_received),
            "date_sent": self._to_iso(date_sent) or self._to_iso(date_received),
            "is_read": bool(read),
            "is_flagged": bool(flagged),
            "has_attachments": bool(has_att),
            "mailbox_name": mbox_name,
            "account_name": account,
            "message_id": rfc_message_id,
            "in_reply_to": in_reply_to,
            "size": size or 0,
            "body_text": body_text,
            "to_recipients": to_recipients,
            "cc_recipients": cc_recipients,
        }

    @staticmethod
    def _header_addresses(msg, header: str) -> list[str]:
        import email.utils as email_utils

        out = []
        for name, addr in email_utils.getaddresses(msg.get_all(header, [])):
            if not addr:
                continue
            out.append(f"{name} <{addr}>" if name else addr)
        return out

    def _recipients_from_db(self, rowid: int) -> tuple[list[str], list[str]]:
        if not self._rec_msg_col:
            return [], []
        type_col = f"r.{self._rec_type_col}" if self._rec_type_col else "0"
        rows = self._query(
            f"""
            SELECT {type_col}, COALESCE(ra.comment, ''), COALESCE(ra.address, '')
            FROM recipients r
            LEFT JOIN addresses ra ON ra.ROWID = r.address
            WHERE r.{self._rec_msg_col} = ?
            ORDER BY r.ROWID
            """,
            (rowid,),
        )
        to_list, cc_list = [], []
        for rtype, comment, address in rows:
            if not address:
                continue
            display = f"{comment} <{address}>" if comment else address
            if rtype == 1:
                cc_list.append(display)
            elif rtype in (0, None):
                to_list.append(display)
        return to_list, cc_list

    def get_message_id_header(self, message_id: int) -> Optional[str]:
        msg = self._read_message_file(message_id)
        if msg is None:
            return None
        rfc_id = (msg.get("Message-ID") or "").strip().strip("<>")
        return rfc_id or None

    def get_message_source(self, message_id: int) -> Optional[str]:
        path = self._find_emlx(message_id)
        if path is None:
            return None
        try:
            return read_emlx_message_bytes(path).decode("utf-8", errors="replace")
        except (OSError, ValueError) as exc:
            logger.warning("Failed to read source for %d: %s", message_id, exc)
            return None

    def get_thread_messages(self, message_id: int) -> list[dict]:
        if self._has_conversation:
            rows = self._query(
                "SELECT conversation_id FROM messages WHERE ROWID = ?",
                (message_id,),
            )
            if not rows:
                return []
            conversation_id = rows[0][0]
            if conversation_id and conversation_id > 0:
                thread_rows = self._query(
                    f"""
                    SELECT m.ROWID FROM messages m
                    WHERE m.conversation_id = ? AND {self._not_deleted_expr}
                    ORDER BY m.date_received ASC
                    LIMIT 200
                    """,
                    (conversation_id,),
                )
                ids = [r[0] for r in thread_rows]
                if ids:
                    return self._summaries_for_ids(ids)
        # Fallback: single message (server falls back to get_message anyway)
        d = self.get_message(message_id)
        return [d] if d else []

    def _summaries_for_ids(self, ids: list[int]) -> list[dict]:
        placeholders = ",".join("?" for _ in ids)
        rows = self._query(
            f"""
            SELECT m.ROWID,
                   {self._subject_prefix_expr} || COALESCE(s.subject, ''),
                   CASE WHEN COALESCE(a.comment, '') != ''
                        THEN a.comment || ' <' || COALESCE(a.address, '') || '>'
                        ELSE COALESCE(a.address, '') END,
                   m.date_received,
                   {self._date_sent_expr},
                   {self._read_expr},
                   {self._flagged_expr},
                   {self._has_attachments_expr},
                   m.mailbox,
                   {self._size_expr}
            FROM messages m
            LEFT JOIN subjects s ON m.subject = s.ROWID
            LEFT JOIN addresses a ON m.sender = a.ROWID
            WHERE m.ROWID IN ({placeholders})
            ORDER BY (m.date_received IS NULL OR m.date_received = 0),
                     m.date_received ASC
            """,
            tuple(ids),
        )
        results = []
        for (
            rowid, subject, sender, date_received, date_sent,
            read, flagged, has_att, mailbox_rowid, size,
        ) in rows:
            account, mbox_name, _url = self._mailbox_info(mailbox_rowid)
            results.append(
                {
                    "id": rowid,
                    "subject": subject or "",
                    "sender": sender or "",
                    "date_received": self._to_iso(date_received),
                    "date_sent": self._to_iso(date_sent) or self._to_iso(date_received),
                    "is_read": bool(read),
                    "is_flagged": bool(flagged),
                    "has_attachments": bool(has_att),
                    "mailbox_name": mbox_name,
                    "account_name": account,
                    "message_id": None,
                    "in_reply_to": None,
                    "size": size or 0,
                }
            )
        return results

    # ------------------------------------------------------------------
    # Attachments
    # ------------------------------------------------------------------

    def list_attachments(self, message_id: int) -> list[dict]:
        msg = self._read_message_file(message_id)
        if msg is None:
            return []
        result = []
        index = 0
        for part in msg.walk():
            filename = part.get_filename()
            if not filename and part.get_content_disposition() != "attachment":
                continue
            size = 0
            payload = part.get_payload(decode=True)
            if payload:
                size = len(payload)
            else:
                # .partial.emlx stubs carry the true size in a header
                apple_len = part.get("X-Apple-Content-Length")
                if apple_len and apple_len.strip().isdigit():
                    size = int(apple_len.strip())
            result.append(
                {
                    "index": index,
                    "name": filename or f"attachment_{index}",
                    "mime_type": part.get_content_type(),
                    "file_size": size,
                }
            )
            index += 1
        return result

    def get_attachment(
        self, message_id: int, attachment_index: int
    ) -> Optional[tuple[str, str, bytes]]:
        """Return (filename, mime_type, bytes) or None.

        None means the caller should fall back to the JXA bridge (message
        unknown, or the attachment body lives outside the .emlx and hasn't
        been located on disk).
        """
        msg = self._read_message_file(message_id)
        if msg is None:
            return None
        index = 0
        for part in msg.walk():
            filename = part.get_filename()
            if not filename and part.get_content_disposition() != "attachment":
                continue
            if index != attachment_index:
                index += 1
                continue
            name = filename or f"attachment_{attachment_index}"
            mime_type = part.get_content_type()
            payload = part.get_payload(decode=True)
            if payload:
                return name, mime_type, payload
            external = self._find_external_attachment(message_id, name)
            if external is not None:
                try:
                    return name, mime_type, external.read_bytes()
                except OSError as exc:
                    logger.warning("Failed reading %s: %s", external, exc)
            return None
        return None

    def _find_external_attachment(
        self, message_id: int, filename: str
    ) -> Optional[Path]:
        """Locate a stripped attachment body next to a .partial.emlx.

        Layout: .../Data/<n>/.../Messages/<id>.partial.emlx with bodies at
        .../Data/<n>/.../Attachments/<id>/<part>/<filename>.
        """
        emlx_path = self._find_emlx(message_id)
        if emlx_path is None:
            return None
        attachments_dir = emlx_path.parent.parent / "Attachments" / str(message_id)
        if not self._exists(attachments_dir):
            return None
        try:
            for candidate in attachments_dir.rglob("*"):
                if candidate.is_file() and candidate.name == filename:
                    return candidate
            # Fall back to any single file if names differ (Mail sometimes
            # renames on disk)
            files = [p for p in attachments_dir.rglob("*") if p.is_file()]
            if len(files) == 1:
                return files[0]
        except OSError:
            return None
        return None

    # ------------------------------------------------------------------
    # Flags (read side)
    # ------------------------------------------------------------------

    def get_flag(self, message_id: int) -> dict:
        """Flag state from the index. Color is only available when the
        schema exposes it; callers needing an authoritative color should
        use the JXA bridge when is_flagged is true and color_index < 0."""
        color_select = (
            f"m.{self._flag_color_col}" if self._flag_color_col else "-1"
        )
        rows = self._query(
            f"SELECT {self._flagged_expr}, {color_select} "
            "FROM messages m WHERE m.ROWID = ?",
            (message_id,),
        )
        if not rows:
            raise ValueError(f"Message {message_id} not found.")
        flagged, color_index = rows[0]
        if not flagged:
            color_index = -1
        return {
            "is_flagged": bool(flagged),
            "color_index": color_index if isinstance(color_index, int) else -1,
        }
