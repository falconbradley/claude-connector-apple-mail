"""
Inline email preview card (MCP Apps, SEP-1865).

Chat hosts that implement the MCP Apps extension -- Claude Desktop and
claude.ai among them -- can render a tool's result with an HTML view the
server ships as a ``ui://`` resource. The host reads the resource once,
loads it in a sandboxed iframe next to the tool call in the transcript,
and forwards the tool's ``structuredContent`` to it over postMessage.

This module owns both halves of that contract, for both of the
connector's cards:

* the resources themselves -- ``ui/email_preview.html`` (one message) and
  ``ui/thread_preview.html`` (a whole conversation), each a single
  self-contained document (no external scripts, styles, or images, so
  they run under the host's restrictive default CSP);
* the payload builders that turn bridge message dicts into the shape each
  card renders, keeping the (possibly huge) HTML bodies out of the text
  the model sees.

Hosts without MCP Apps ignore the ``_meta.ui`` hint and see the text
content only, so the tool still degrades to a normal ``get_email`` answer.
"""

from __future__ import annotations

import email as email_lib
import json
from pathlib import Path
from typing import Any, Optional

from mcp.types import CallToolResult, TextContent

from .emlx import get_html_body

# Identifier hosts use to find the view for the tool; must match the tool's
# ``_meta.ui.resourceUri``. Stable across releases so cached templates stay
# valid.
UI_RESOURCE_URI = "ui://apple-mail/email-preview"
THREAD_UI_RESOURCE_URI = "ui://apple-mail/thread-preview"

# The only content type the MCP Apps MVP defines.
UI_MIME_TYPE = "text/html;profile=mcp-app"

# Resource-level rendering hints. No CSP block: the card needs no network
# access at all, so the host's restrictive default is exactly right. The
# card draws its own border, so ask the host not to add another.
UI_RESOURCE_META: dict[str, Any] = {"ui": {"prefersBorder": False}}

# Beyond this the HTML body is dropped from the card (plain text is shown
# instead). Newsletter-sized mail is a few hundred KB; this is a safety
# valve against pathological messages, not a typical-case limit.
MAX_HTML_BYTES = 1_500_000

# A thread card carries many bodies at once, so the per-message ceiling is
# not enough on its own. Budget applied newest-first: the messages the user
# is most likely to read keep their rich bodies, older ones fall back to
# plain text rather than the card shipping tens of megabytes to the host.
MAX_THREAD_HTML_BYTES = 2_500_000

# Conversations can be long (the bridge caps thread reads at 200). Past
# this the card shows the most recent slice and says how many it dropped --
# rendering 200 bodies helps nobody.
DEFAULT_THREAD_LIMIT = 25

_HTML_PATH = Path(__file__).parent / "ui" / "email_preview.html"
_THREAD_HTML_PATH = Path(__file__).parent / "ui" / "thread_preview.html"


def load_preview_html() -> str:
    """Return the email card's HTML document.

    Read on every call rather than cached at import: the host fetches the
    resource once per connection anyway, and a fresh read means an edit
    during development shows up without restarting the server.
    """
    return _HTML_PATH.read_text(encoding="utf-8")


def load_thread_preview_html() -> str:
    """Return the thread card's HTML document (see ``load_preview_html``)."""
    return _THREAD_HTML_PATH.read_text(encoding="utf-8")


def extract_html_body(source: Optional[str]) -> Optional[str]:
    """Pull the text/html part out of a raw RFC 2822 message, or None.

    Oversized bodies return None so the card falls back to plain text
    instead of shipping megabytes through the host.
    """
    if not source:
        return None
    try:
        msg = email_lib.message_from_string(source)
        html = get_html_body(msg)
    except Exception:
        return None
    if not html or not html.strip():
        return None
    if len(html.encode("utf-8", errors="replace")) > MAX_HTML_BYTES:
        return None
    return html


def build_preview_payload(
    message: dict[str, Any],
    *,
    attachments: list[dict[str, Any]],
    flag_color: Optional[str],
    body_html: Optional[str],
    mail_link: Optional[str],
    open_link: Optional[str],
) -> dict[str, Any]:
    """Shape a bridge message dict into what the card renders.

    Field names are the card's contract (see ``email_preview.html``), not
    the ``EmailDetail`` model's -- the card is the only consumer, and it
    wants recipient lists under ``to``/``cc`` and an ``attachments`` list
    rather than a count.
    """
    return {
        "id": message["id"],
        "subject": message.get("subject") or "(no subject)",
        "sender": message.get("sender") or "",
        "to": list(message.get("to_recipients") or []),
        "cc": list(message.get("cc_recipients") or []),
        "date_sent": message.get("date_sent"),
        "date_received": message.get("date_received"),
        "mailbox": message.get("mailbox_name") or "",
        "account": message.get("account_name") or "",
        "is_read": bool(message.get("is_read", True)),
        "is_flagged": bool(message.get("is_flagged", False)),
        "flag_color": flag_color,
        "has_attachments": bool(attachments) or bool(message.get("has_attachments")),
        "attachments": [
            {
                "filename": a.get("name") or "attachment",
                "content_type": a.get("mime_type") or "",
                "size": a.get("file_size") or 0,
            }
            for a in attachments
        ],
        "size": message.get("size") or 0,
        "message_id": message.get("message_id"),
        "body_text": message.get("body_text") or "",
        "body_html": body_html,
        "mail_link": mail_link,
        "open_link": open_link,
    }


def build_preview_result(payload: dict[str, Any], model_view: dict[str, Any]) -> CallToolResult:
    """Assemble the tool result: text for the model, structured data for the card.

    ``content`` is what the model reads and what non-Apps hosts display, so
    it carries the same JSON ``get_email`` returns -- no HTML body, which
    would only burn context. ``structuredContent`` goes to the iframe and
    carries everything, HTML included.
    """
    return CallToolResult(
        content=[TextContent(type="text", text=json.dumps(model_view, indent=2, default=str))],
        structuredContent=payload,
    )


def apply_thread_html_budget(
    entries: list[dict[str, Any]], budget: int = MAX_THREAD_HTML_BYTES
) -> int:
    """Drop ``body_html`` from oldest entries until the total fits ``budget``.

    ``entries`` is in chronological order, so the walk runs backwards: the
    newest messages -- the ones the user actually wants to read -- keep
    their rich bodies, and older ones fall back to the plain text the card
    already has. Returns how many bodies were dropped.
    """
    spent = 0
    dropped = 0
    for entry in reversed(entries):
        html = entry.get("body_html")
        if not html:
            continue
        cost = len(html.encode("utf-8", errors="replace"))
        if spent + cost > budget:
            entry["body_html"] = None
            dropped += 1
        else:
            spent += cost
    return dropped


def build_thread_payload(
    entries: list[dict[str, Any]],
    *,
    subject: str,
    total: int,
    thread_link: Optional[str],
    mail_link: Optional[str],
    open_link: Optional[str],
) -> dict[str, Any]:
    """Shape a conversation into what the thread card renders.

    ``entries`` are per-message payloads in the *same* shape the single-email
    card consumes (``build_preview_payload``), chronological, oldest first --
    one contract, two cards. Everything else here is thread-level: the base
    subject, who took part, the span, and the links that open the
    conversation in Mail.

    ``total`` is the conversation's real length; when it exceeds the entries
    shown, the card says how many are missing rather than silently lying
    about the thread's size.
    """
    participants: list[str] = []
    seen: set[str] = set()
    for entry in entries:
        sender = (entry.get("sender") or "").strip()
        if not sender:
            continue
        key = sender.lower()
        if key not in seen:
            seen.add(key)
            participants.append(sender)

    dates = [e.get("date_sent") or e.get("date_received") for e in entries]
    dates = [d for d in dates if d]

    return {
        "subject": subject or "(no subject)",
        "message_count": total,
        "shown_count": len(entries),
        "omitted_count": max(0, total - len(entries)),
        "participants": participants,
        "date_start": dates[0] if dates else None,
        "date_end": dates[-1] if dates else None,
        "thread_link": thread_link,
        "mail_link": mail_link,
        "open_link": open_link,
        "messages": entries,
    }
