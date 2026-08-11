"""
MIME body extraction and .emlx file utilities.

Helpers for extracting plain-text and HTML bodies from parsed
email.message.Message objects, and for reading Mail.app's on-disk
.emlx message files.
"""

from __future__ import annotations

import email as email_lib
import html as html_lib
import re
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# .emlx file format
# ---------------------------------------------------------------------------
# An .emlx file is:  <byte count>\n<RFC 2822 message of that many bytes><XML plist>
# The plist carries Mail-internal metadata (flags etc.); we only need the
# message. ".partial.emlx" variants have attachment bodies stripped out and
# stored as separate files under a sibling Attachments/ directory.

def read_emlx_message_bytes(path: Path) -> bytes:
    """Return the raw RFC 2822 message bytes from an .emlx file."""
    data = path.read_bytes()
    newline = data.find(b"\n")
    if newline == -1:
        raise ValueError(f"Not an emlx file (no length line): {path}")
    try:
        length = int(data[:newline].strip())
    except ValueError:
        raise ValueError(f"Not an emlx file (bad length line): {path}")
    start = newline + 1
    return data[start : start + length]


def read_emlx(path: Path) -> email_lib.message.Message:
    """Parse an .emlx file into an email.message.Message."""
    return email_lib.message_from_bytes(read_emlx_message_bytes(path))


def read_emlx_headers(path: Path) -> email_lib.message.Message:
    """Parse only the header block of an .emlx file (cheap: reads the
    first few KB instead of the whole message)."""
    with path.open("rb") as fh:
        first_line = fh.readline()
        if not first_line.strip().isdigit():
            raise ValueError(f"Not an emlx file (bad length line): {path}")
        chunks = []
        while True:
            chunk = fh.read(8192)
            if not chunk:
                break
            chunks.append(chunk)
            blob = b"".join(chunks) if len(chunks) > 1 else chunk
            end = blob.find(b"\r\n\r\n")
            if end == -1:
                end = blob.find(b"\n\n")
            if end != -1:
                return email_lib.message_from_bytes(blob[:end])
    return email_lib.message_from_bytes(b"".join(chunks))


# ---------------------------------------------------------------------------
# HTML -> text (fallback when a message has no text/plain part)
# ---------------------------------------------------------------------------

_HTML_BLOCK_RE = re.compile(
    r"<(?:br|/p|/div|/tr|/li|/h[1-6]|/table)[^>]*>", re.IGNORECASE
)
_HTML_STRIP_RE = re.compile(
    r"<(?:script|style)\b.*?</(?:script|style)>", re.IGNORECASE | re.DOTALL
)
_HTML_TAG_RE = re.compile(r"<[^>]+>")


def html_to_text(html: str) -> str:
    """Very light HTML-to-text conversion for body previews."""
    text = _HTML_STRIP_RE.sub("", html)
    text = _HTML_BLOCK_RE.sub("\n", text)
    text = _HTML_TAG_RE.sub("", text)
    text = html_lib.unescape(text)
    # Collapse runs of blank lines / spaces
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n\s*\n+", "\n\n", text)
    return text.strip()


# ---------------------------------------------------------------------------
# Body extraction
# ---------------------------------------------------------------------------

def get_text_body(msg: email_lib.message.Message) -> str:
    """Return concatenated plain-text parts."""
    parts: list[str] = []
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                _append_part(part, parts)
    elif msg.get_content_type() == "text/plain":
        _append_part(msg, parts)
    return "\n\n".join(parts)


def get_html_body(msg: email_lib.message.Message) -> Optional[str]:
    """Return the first HTML part, or None."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                result = _decode_part(part)
                if result:
                    return result
    elif msg.get_content_type() == "text/html":
        return _decode_part(msg)
    return None


def _append_part(part: email_lib.message.Message, out: list[str]) -> None:
    text = _decode_part(part)
    if text:
        out.append(text)


def _decode_part(part: email_lib.message.Message) -> Optional[str]:
    payload = part.get_payload(decode=True)
    if not payload:
        return None
    charset = part.get_content_charset("utf-8") or "utf-8"
    return payload.decode(charset, errors="replace")
