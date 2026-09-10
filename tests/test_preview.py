"""
Tests for the inline email preview card (MCP Apps).

Two layers:

* Unit: the payload builder, the HTML-body extractor, and the shipped HTML
  document itself (self-contained, speaks the right protocol methods).
* Protocol: an in-memory MCP client session against the real server object
  with a fake bridge, asserting what a host actually sees on the wire --
  the tool's ``_meta.ui.resourceUri``, the ``ui://`` resource with the
  MCP Apps MIME type, and a tools/call result whose text carries no HTML
  while its structuredContent does.

No Mail.app involved; runs anywhere.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from apple_mail_mcp import preview, server  # noqa: E402
from apple_mail_mcp.preview import (  # noqa: E402
    UI_MIME_TYPE,
    UI_RESOURCE_URI,
    build_preview_payload,
    extract_html_body,
    load_preview_html,
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

_MESSAGE = {
    "id": 4242,
    "subject": "Q3 board deck — draft 2",
    "sender": "Alice Example <alice@example.com>",
    "date_received": "2026-09-08T17:02:11+00:00",
    "date_sent": "2026-09-08T17:01:58+00:00",
    "is_read": False,
    "is_flagged": True,
    "has_attachments": True,
    "mailbox_name": "INBOX",
    "account_name": "brad@impulselabs.com",
    "message_id": "abc123@example.com",
    "in_reply_to": None,
    "size": 48213,
    "body_text": "Hi Brad,\n\nDraft 2 attached. Notes inline.\n\n> older quoted text\n\n— Alice",
    "to_recipients": ["Brad <brad@impulselabs.com>"],
    "cc_recipients": ["Team <team@impulselabs.com>"],
}

_ATTACHMENTS = [
    {"index": 0, "name": "board-deck-v2.pdf", "mime_type": "application/pdf", "file_size": 1_204_331},
]

_SOURCE = (
    "From: Alice Example <alice@example.com>\r\n"
    "To: Brad <brad@impulselabs.com>\r\n"
    "Subject: Q3 board deck\r\n"
    "Message-ID: <abc123@example.com>\r\n"
    "MIME-Version: 1.0\r\n"
    'Content-Type: multipart/alternative; boundary="b1"\r\n'
    "\r\n"
    "--b1\r\n"
    "Content-Type: text/plain; charset=utf-8\r\n"
    "\r\n"
    "Hi Brad, draft 2 attached.\r\n"
    "--b1\r\n"
    "Content-Type: text/html; charset=utf-8\r\n"
    "\r\n"
    "<html><body><p>Hi <b>Brad</b>, draft 2 attached.</p>"
    "<script>alert(1)</script></body></html>\r\n"
    "--b1--\r\n"
)


class FakeBridge:
    """Just enough of HybridBridge for the preview tool."""

    def __init__(self, *, source: str | None = _SOURCE, flagged: bool = True) -> None:
        self.source = source
        self.message = dict(_MESSAGE, is_flagged=flagged)
        self.calls: list[str] = []

    def get_message(self, message_id: int):
        self.calls.append("get_message")
        return dict(self.message) if message_id == self.message["id"] else None

    def list_attachments(self, message_id: int):
        self.calls.append("list_attachments")
        return list(_ATTACHMENTS)

    def get_flag(self, message_id: int):
        self.calls.append("get_flag")
        return {"is_flagged": True, "color_index": 1, "flag_color": "orange"}

    def get_message_source(self, message_id: int):
        self.calls.append("get_message_source")
        return self.source

    def get_message_id_header(self, message_id: int):
        return self.message["message_id"]


@pytest.fixture
def fake_bridge(monkeypatch):
    bridge = FakeBridge()
    monkeypatch.setattr(server, "_bridge", bridge)
    # Keep the localhost redirector out of unit tests.
    monkeypatch.setattr(server, "_make_open_link", lambda mid: f"http://127.0.0.1:1/open/{mid}?t=x")
    return bridge


# ---------------------------------------------------------------------------
# Unit: payload + HTML extraction
# ---------------------------------------------------------------------------

def test_payload_shape_matches_card_contract():
    p = build_preview_payload(
        _MESSAGE,
        attachments=_ATTACHMENTS,
        flag_color="orange",
        body_html="<p>x</p>",
        mail_link="message://%3Cabc123%40example.com%3E",
        open_link="http://127.0.0.1:1/open/4242?t=x",
    )
    assert p["id"] == 4242
    assert p["to"] == ["Brad <brad@impulselabs.com>"]
    assert p["cc"] == ["Team <team@impulselabs.com>"]
    assert p["is_read"] is False and p["is_flagged"] is True
    assert p["flag_color"] == "orange"
    assert p["attachments"] == [
        {"filename": "board-deck-v2.pdf", "content_type": "application/pdf", "size": 1_204_331}
    ]
    assert p["has_attachments"] is True
    assert p["body_html"] == "<p>x</p>"
    assert p["body_text"].startswith("Hi Brad")
    assert p["open_link"].startswith("http://127.0.0.1")
    assert p["mail_link"].startswith("message://")


def test_payload_defaults_when_message_is_sparse():
    p = build_preview_payload(
        {"id": 1}, attachments=[], flag_color=None, body_html=None, mail_link=None, open_link=None
    )
    assert p["subject"] == "(no subject)"
    assert p["to"] == [] and p["cc"] == [] and p["attachments"] == []
    assert p["is_read"] is True and p["is_flagged"] is False
    assert p["has_attachments"] is False
    assert p["body_text"] == "" and p["body_html"] is None


def test_extract_html_body_finds_the_html_part():
    html = extract_html_body(_SOURCE)
    assert html is not None
    assert "<b>Brad</b>" in html


@pytest.mark.parametrize("source", [None, "", "Subject: x\r\n\r\nplain only\r\n"])
def test_extract_html_body_returns_none_without_html(source):
    assert extract_html_body(source) is None


def test_extract_html_body_drops_oversized_bodies(monkeypatch):
    monkeypatch.setattr(preview, "MAX_HTML_BYTES", 10)
    assert extract_html_body(_SOURCE) is None


# ---------------------------------------------------------------------------
# Unit: the shipped HTML document
# ---------------------------------------------------------------------------

def test_preview_html_is_a_self_contained_document():
    html = load_preview_html()
    assert html.lstrip().lower().startswith("<!doctype html>")
    # The host's default CSP allows no network: nothing may reference one.
    assert not re.search(r"""(src|href)\s*=\s*["']https?://""", html), "external asset reference"
    assert "<link" not in html.lower()
    # The stylesheet must not pull anything in either (the sanitizer's own
    # blocklist mentions @import in script, which is fine).
    for css in re.findall(r"<style[^>]*>(.*?)</style>", html, flags=re.S | re.I):
        assert "@import" not in css and "url(" not in css


def test_preview_html_speaks_the_mcp_apps_protocol():
    html = load_preview_html()
    for method in (
        "ui/initialize",
        "ui/notifications/initialized",
        "ui/notifications/tool-input",
        "ui/notifications/tool-result",
        "ui/notifications/tool-cancelled",
        "ui/notifications/host-context-changed",
        "ui/notifications/size-changed",
        "ui/resource-teardown",
        "ui/open-link",
        '"ping"',
    ):
        assert method in html, f"missing {method}"
    assert '"2026-01-26"' in html, "protocol version"


def test_preview_html_reads_every_payload_field_it_is_sent():
    """Guard against the Python payload and the JS card drifting apart."""
    html = load_preview_html()
    payload = build_preview_payload(
        _MESSAGE, attachments=_ATTACHMENTS, flag_color="orange",
        body_html="<p></p>", mail_link="m", open_link="o",
    )
    for key in payload:
        assert re.search(rf"\bm\.{key}\b|\bmsg\.{key}\b|\[\"{key}\"\]", html), f"card never reads {key!r}"


# ---------------------------------------------------------------------------
# Protocol: what a host sees over a real MCP session
# ---------------------------------------------------------------------------

@pytest.fixture
def session_factory():
    from mcp.shared.memory import create_connected_server_and_client_session

    lowlevel = getattr(server.mcp, "_mcp_server")
    return lambda: create_connected_server_and_client_session(lowlevel, raise_exceptions=True)


@pytest.mark.anyio
async def test_tool_advertises_its_ui_resource(session_factory):
    async with session_factory() as client:
        tools = (await client.list_tools()).tools
    by_name = {t.name: t for t in tools}
    assert "preview_email" in by_name
    meta = by_name["preview_email"].meta
    assert meta == {"ui": {"resourceUri": UI_RESOURCE_URI}}
    # get_email stays a plain text tool; only the preview renders a card.
    assert not (by_name["get_email"].meta or {}).get("ui")


@pytest.mark.anyio
async def test_ui_resource_is_listed_and_readable(session_factory):
    async with session_factory() as client:
        listed = (await client.list_resources()).resources
        found = [r for r in listed if str(r.uri) == UI_RESOURCE_URI]
        assert found, "ui:// resource not in resources/list"
        assert found[0].mimeType == UI_MIME_TYPE
        assert found[0].meta == {"ui": {"prefersBorder": False}}

        read = await client.read_resource(UI_RESOURCE_URI)
    contents = read.contents
    assert len(contents) == 1
    assert contents[0].mimeType == UI_MIME_TYPE
    assert contents[0].meta == {"ui": {"prefersBorder": False}}
    assert contents[0].text.lstrip().lower().startswith("<!doctype html>")


@pytest.mark.anyio
async def test_call_keeps_html_out_of_model_text_but_in_structured_content(session_factory, fake_bridge):
    async with session_factory() as client:
        result = await client.call_tool("preview_email", {"message_id": 4242})

    assert result.isError is False
    text = "".join(c.text for c in result.content if c.type == "text")
    model_view = json.loads(text)
    # The model sees exactly what get_email would have given it...
    assert model_view["id"] == 4242
    assert model_view["subject"] == _MESSAGE["subject"]
    assert model_view["body_text"] == _MESSAGE["body_text"]
    assert model_view["flag_color"] == "orange"
    assert model_view["attachment_count"] == 1
    assert "<b>Brad</b>" not in text and "<script>" not in text

    # ...while the card gets the full payload, HTML included.
    sc = result.structuredContent
    assert sc["id"] == 4242
    assert "<b>Brad</b>" in sc["body_html"]
    assert sc["attachments"][0]["filename"] == "board-deck-v2.pdf"
    assert sc["open_link"] == "http://127.0.0.1:1/open/4242?t=x"
    assert sc["mail_link"] == "message://%3Cabc123%40example.com%3E"

    # One source read for the HTML, on top of what get_email does.
    assert fake_bridge.calls.count("get_message_source") == 1
    assert fake_bridge.calls.count("get_message") == 1


@pytest.mark.anyio
async def test_call_falls_back_to_text_when_source_is_unavailable(session_factory, fake_bridge):
    fake_bridge.source = None
    async with session_factory() as client:
        result = await client.call_tool("preview_email", {"message_id": 4242})
    assert result.isError is False
    assert result.structuredContent["body_html"] is None
    assert result.structuredContent["body_text"].startswith("Hi Brad")


@pytest.mark.anyio
async def test_call_reports_missing_message_as_tool_error(session_factory, fake_bridge):
    async with session_factory() as client:
        result = await client.call_tool("preview_email", {"message_id": 99})
    assert result.isError is True
    assert "99" in result.content[0].text
