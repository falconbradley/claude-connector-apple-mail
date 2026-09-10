"""
Inline email preview card (MCP Apps, SEP-1865).

Chat hosts that implement the MCP Apps extension -- Claude Desktop and
claude.ai among them -- can render a tool's result with an HTML view the
server ships as a ``ui://`` resource. The host reads the resource once,
loads it in a sandboxed iframe next to the tool call in the transcript,
and forwards the tool's ``structuredContent`` to it over postMessage.

This module owns the two halves of that contract for the connector:

* the resource itself -- ``ui/email_preview.html``, a single self-contained
  document (no external scripts, styles, or images, so it runs under the
  host's restrictive default CSP);
* the payload builder that turns a message dict from the bridge into the
  shape the card renders, keeping the (possibly huge) HTML body out of the
  text the model sees.

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

_HTML_PATH = Path(__file__).parent / "ui" / "email_preview.html"


def load_preview_html() -> str:
    """Return the card's HTML document.

    Read on every call rather than cached at import: the host fetches the
    resource once per connection anyway, and a fresh read means an edit
    during development shows up without restarting the server.
    """
    return _HTML_PATH.read_text(encoding="utf-8")


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
