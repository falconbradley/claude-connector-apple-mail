"""
MIME body extraction utilities.

Helpers for extracting plain-text and HTML bodies from parsed
email.message.Message objects.
"""

from __future__ import annotations

import email as email_lib
from typing import Optional


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
