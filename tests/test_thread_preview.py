"""
Tests for the conversation preview card and thread-scoped links.

Same two layers as ``test_preview.py``:

* Unit: the thread payload builder, the HTML budget, and the shipped card
  document (self-contained, speaks the protocol, reads every field it is
  sent).
* Protocol: an in-memory MCP client session against the real server object
  with a fake bridge -- what a host actually sees for ``preview_thread``,
  and what ``get_email_link`` returns for each scope.

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
    THREAD_UI_RESOURCE_URI,
    UI_MIME_TYPE,
    apply_thread_html_budget,
    build_preview_payload,
    build_thread_payload,
    load_thread_preview_html,
)


# ---------------------------------------------------------------------------
# Fixtures: a three-message conversation across Inbox and Sent
# ---------------------------------------------------------------------------

def _msg(mid, sender, subject, when, mailbox="INBOX", **extra):
    base = {
        "id": mid,
        "subject": subject,
        "sender": sender,
        "date_received": when,
        "date_sent": when,
        "is_read": True,
        "is_flagged": False,
        "has_attachments": False,
        "mailbox_name": mailbox,
        "account_name": "iCloud",
        "message_id": f"msg{mid}@example.com",
        "in_reply_to": None,
        "size": 1000,
        "body_text": f"Body of message {mid}.",
        "to_recipients": ["Brad <brad@impulselabs.com>"],
        "cc_recipients": [],
    }
    base.update(extra)
    return base


_THREAD = [
    _msg(10, "Alice <alice@example.com>", "Contract renewal", "2026-09-01T10:00:00+00:00"),
    _msg(11, "Brad <brad@impulselabs.com>", "Re: Contract renewal",
         "2026-09-02T09:00:00+00:00", mailbox="Sent Messages"),
    _msg(12, "Alice <alice@example.com>", "Re: Contract renewal",
         "2026-09-03T11:30:00+00:00", is_flagged=True, has_attachments=True),
]

_SOURCES = {
    m["id"]: (
        "Subject: x\r\n"
        'Content-Type: text/html; charset=utf-8\r\n'
        "\r\n"
        f"<html><body><p>Rich body {m['id']}</p></body></html>\r\n"
    )
    for m in _THREAD
}


class FakeBridge:
    """Just enough of HybridBridge for the thread tools."""

    def __init__(self, thread=None) -> None:
        self.thread = [dict(m) for m in (thread if thread is not None else _THREAD)]
        self.calls: list[str] = []

    def _by_id(self, message_id):
        return next((dict(m) for m in self.thread if m["id"] == message_id), None)

    def get_thread_messages(self, message_id: int):
        self.calls.append("get_thread_messages")
        if self._by_id(message_id) is None:
            return []
        return [dict(m) for m in self.thread]

    def get_message(self, message_id: int):
        self.calls.append("get_message")
        return self._by_id(message_id)

    def list_attachments(self, message_id: int):
        self.calls.append("list_attachments")
        return [{"index": 0, "name": "terms.pdf", "mime_type": "application/pdf",
                 "file_size": 2048}]

    def get_flag(self, message_id: int):
        self.calls.append("get_flag")
        return {"is_flagged": True, "color_index": 1, "flag_color": "orange"}

    def get_message_source(self, message_id: int):
        self.calls.append("get_message_source")
        return _SOURCES.get(message_id)

    def get_message_id_header(self, message_id: int):
        m = self._by_id(message_id)
        return m["message_id"] if m else None


@pytest.fixture
def fake_bridge(monkeypatch):
    bridge = FakeBridge()
    monkeypatch.setattr(server, "_bridge", bridge)
    # Keep the localhost redirector out of unit tests.
    monkeypatch.setattr(server, "_make_open_link", lambda mid: f"http://127.0.0.1:1/open/{mid}?t=x")
    monkeypatch.setattr(server, "_make_thread_link", lambda mid: f"http://127.0.0.1:1/thread/{mid}?t=x")
    return bridge


def _entries(messages=_THREAD, body_html=None):
    return [
        build_preview_payload(
            m, attachments=[], flag_color=None,
            body_html=body_html, mail_link="m", open_link="o",
        )
        for m in messages
    ]


# ---------------------------------------------------------------------------
# Unit: thread payload
# ---------------------------------------------------------------------------

def test_thread_payload_summarises_the_conversation():
    p = build_thread_payload(
        _entries(), subject="Contract renewal", total=3,
        thread_link="t", mail_link="m", open_link="o",
    )
    assert p["subject"] == "Contract renewal"
    assert p["message_count"] == 3 and p["shown_count"] == 3
    assert p["omitted_count"] == 0
    # Participants are unique, in the order they first speak.
    assert p["participants"] == [
        "Alice <alice@example.com>", "Brad <brad@impulselabs.com>",
    ]
    assert p["date_start"] == "2026-09-01T10:00:00+00:00"
    assert p["date_end"] == "2026-09-03T11:30:00+00:00"
    assert p["thread_link"] == "t"
    assert [m["id"] for m in p["messages"]] == [10, 11, 12]


def test_thread_payload_counts_messages_it_is_not_showing():
    p = build_thread_payload(
        _entries(_THREAD[-2:]), subject="Contract renewal", total=40,
        thread_link=None, mail_link=None, open_link=None,
    )
    assert p["message_count"] == 40
    assert p["shown_count"] == 2
    assert p["omitted_count"] == 38


def test_thread_payload_survives_an_empty_conversation():
    p = build_thread_payload(
        [], subject="", total=0, thread_link=None, mail_link=None, open_link=None,
    )
    assert p["subject"] == "(no subject)"
    assert p["participants"] == [] and p["messages"] == []
    assert p["date_start"] is None and p["date_end"] is None


def test_html_budget_keeps_the_newest_bodies_and_drops_the_oldest():
    entries = _entries(body_html="<p>" + "x" * 100 + "</p>")
    dropped = apply_thread_html_budget(entries, budget=250)
    # Chronological in, so the last two survive and the oldest is dropped.
    assert dropped == 1
    assert entries[0]["body_html"] is None
    assert entries[1]["body_html"] and entries[2]["body_html"]


def test_html_budget_leaves_a_thread_that_fits_alone():
    entries = _entries(body_html="<p>small</p>")
    assert apply_thread_html_budget(entries, budget=1_000_000) == 0
    assert all(e["body_html"] for e in entries)


def test_html_budget_ignores_messages_without_html():
    entries = _entries()
    assert apply_thread_html_budget(entries, budget=0) == 0


# ---------------------------------------------------------------------------
# Unit: the shipped HTML document
# ---------------------------------------------------------------------------

def test_thread_html_is_a_self_contained_document():
    html = load_thread_preview_html()
    assert html.lstrip().lower().startswith("<!doctype html>")
    # The host's default CSP allows no network: nothing may reference one.
    assert not re.search(r"""(src|href)\s*=\s*["']https?://""", html), "external asset reference"
    assert "<link" not in html.lower()
    for css in re.findall(r"<style[^>]*>(.*?)</style>", html, flags=re.S | re.I):
        assert "@import" not in css and "url(" not in css


def test_thread_html_speaks_the_mcp_apps_protocol():
    html = load_thread_preview_html()
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


def test_thread_html_reads_every_thread_field_it_is_sent():
    """Guard against the Python payload and the JS card drifting apart."""
    html = load_thread_preview_html()
    payload = build_thread_payload(
        _entries(), subject="s", total=3, thread_link="t", mail_link="m", open_link="o",
    )
    for key in payload:
        assert re.search(rf"\bt\.{key}\b|\bthread\.{key}\b|\[\"{key}\"\]", html), \
            f"card never reads thread field {key!r}"


def test_thread_html_reads_every_per_message_field_it_is_sent():
    html = load_thread_preview_html()
    entry = _entries()[0]
    # Fields the row deliberately does not render: the thread header carries
    # the subject, and size/message_id belong to the single-message card.
    skip = {"subject", "size", "message_id"}
    for key in entry:
        if key in skip:
            continue
        assert re.search(rf"\bm\.{key}\b|\[\"{key}\"\]", html), \
            f"card never reads message field {key!r}"


# ---------------------------------------------------------------------------
# Protocol: what a host sees over a real MCP session
# ---------------------------------------------------------------------------

@pytest.fixture
def session_factory():
    from mcp.shared.memory import create_connected_server_and_client_session

    lowlevel = getattr(server.mcp, "_mcp_server")
    return lambda: create_connected_server_and_client_session(lowlevel, raise_exceptions=True)


@pytest.mark.anyio
async def test_preview_thread_advertises_its_own_ui_resource(session_factory):
    async with session_factory() as client:
        tools = (await client.list_tools()).tools
    by_name = {t.name: t for t in tools}
    assert "preview_thread" in by_name
    meta = by_name["preview_thread"].meta
    assert meta["ui"] == {"resourceUri": THREAD_UI_RESOURCE_URI}
    assert meta["ui/resourceUri"] == THREAD_UI_RESOURCE_URI  # legacy key
    # get_thread stays a plain text tool; only the preview renders a card.
    assert not (by_name["get_thread"].meta or {}).get("ui")


@pytest.mark.anyio
async def test_thread_ui_resource_is_listed_and_readable(session_factory):
    async with session_factory() as client:
        listed = (await client.list_resources()).resources
        found = [r for r in listed if str(r.uri) == THREAD_UI_RESOURCE_URI]
        assert found, "thread ui:// resource not in resources/list"
        assert found[0].mimeType == UI_MIME_TYPE

        read = await client.read_resource(THREAD_UI_RESOURCE_URI)
    contents = read.contents
    assert len(contents) == 1
    assert contents[0].mimeType == UI_MIME_TYPE
    assert contents[0].text.lstrip().lower().startswith("<!doctype html>")


@pytest.mark.anyio
async def test_preview_thread_gives_the_card_bodies_and_the_model_summaries(
    session_factory, fake_bridge
):
    async with session_factory() as client:
        result = await client.call_tool("preview_thread", {"message_id": 11})

    assert result.isError is False
    text = "".join(c.text for c in result.content if c.type == "text")
    model_view = json.loads(text)
    # The model gets the thread's shape, not its contents.
    assert model_view["subject"] == "Contract renewal"      # Re: stripped
    assert model_view["message_count"] == 3
    assert [m["id"] for m in model_view["messages"]] == [10, 11, 12]
    assert "Rich body" not in text and "Body of message" not in text

    # ...while the card gets every message in full.
    sc = result.structuredContent
    assert sc["message_count"] == 3 and sc["omitted_count"] == 0
    assert [m["id"] for m in sc["messages"]] == [10, 11, 12]
    assert "Rich body 12" in sc["messages"][2]["body_html"]
    assert sc["messages"][0]["body_text"] == "Body of message 10."
    assert sc["thread_link"] == "http://127.0.0.1:1/thread/11?t=x"
    # The thread-level links point at the newest message, not the one asked about.
    assert sc["open_link"] == "http://127.0.0.1:1/open/12?t=x"
    assert sc["mail_link"] == "message://%3Cmsg12%40example.com%3E"


@pytest.mark.anyio
async def test_preview_thread_titles_the_conversation_from_its_root(
    session_factory, monkeypatch
):
    """Gateways retitle replies ("[EXTERNAL]Re: ...") but never the original,
    so the thread's name comes from the message that started it."""
    thread = [dict(m) for m in _THREAD]
    for m in thread[1:]:
        m["subject"] = "[EXTERNAL]Re: Contract renewal"
    bridge = FakeBridge(thread=thread)
    monkeypatch.setattr(server, "_bridge", bridge)
    monkeypatch.setattr(server, "_make_open_link", lambda mid: None)
    monkeypatch.setattr(server, "_make_thread_link", lambda mid: None)

    async with session_factory() as client:
        result = await client.call_tool("preview_thread", {"message_id": 12})
    assert result.structuredContent["subject"] == "Contract renewal"

    # A truncated window must not rename the thread after its second message.
    async with session_factory() as client:
        result = await client.call_tool(
            "preview_thread", {"message_id": 12, "limit": 1}
        )
    assert result.structuredContent["subject"] == "Contract renewal"


@pytest.mark.anyio
async def test_preview_thread_only_chases_reads_the_summary_justifies(
    session_factory, fake_bridge
):
    async with session_factory() as client:
        await client.call_tool("preview_thread", {"message_id": 11})
    # Only message 12 is flagged and has attachments, so exactly one of each.
    assert fake_bridge.calls.count("list_attachments") == 1
    assert fake_bridge.calls.count("get_flag") == 1
    assert fake_bridge.calls.count("get_message_source") == 3


@pytest.mark.anyio
async def test_preview_thread_limit_shows_the_newest_and_counts_the_rest(
    session_factory, fake_bridge
):
    async with session_factory() as client:
        result = await client.call_tool("preview_thread", {"message_id": 10, "limit": 2})
    sc = result.structuredContent
    assert sc["message_count"] == 3
    assert sc["shown_count"] == 2 and sc["omitted_count"] == 1
    assert [m["id"] for m in sc["messages"]] == [11, 12]


@pytest.mark.anyio
async def test_preview_thread_renders_a_lone_message_as_a_thread_of_one(
    session_factory, monkeypatch
):
    bridge = FakeBridge(thread=[_THREAD[0]])
    # No conversation id: the bridge finds no thread at all.
    monkeypatch.setattr(bridge, "get_thread_messages", lambda mid: [])
    monkeypatch.setattr(server, "_bridge", bridge)
    monkeypatch.setattr(server, "_make_open_link", lambda mid: None)
    monkeypatch.setattr(server, "_make_thread_link", lambda mid: None)

    async with session_factory() as client:
        result = await client.call_tool("preview_thread", {"message_id": 10})
    sc = result.structuredContent
    assert sc["message_count"] == 1 and sc["omitted_count"] == 0
    assert [m["id"] for m in sc["messages"]] == [10]


@pytest.mark.anyio
async def test_preview_thread_reports_a_missing_message_as_a_tool_error(
    session_factory, fake_bridge
):
    async with session_factory() as client:
        result = await client.call_tool("preview_thread", {"message_id": 999})
    assert result.isError is True
    assert "999" in result.content[0].text


# ---------------------------------------------------------------------------
# Protocol: get_email_link scopes
# ---------------------------------------------------------------------------

@pytest.mark.anyio
async def test_link_defaults_to_the_single_message(session_factory, fake_bridge):
    async with session_factory() as client:
        result = await client.call_tool("get_email_link", {"message_id": 11})
    out = json.loads(result.content[0].text)
    assert out["scope"] == "message"
    assert out["message_id"] == 11
    assert out["mail_link"] == "message://%3Cmsg11%40example.com%3E"
    assert "thread_link" not in out


@pytest.mark.anyio
async def test_thread_scope_points_at_the_conversation(session_factory, fake_bridge):
    async with session_factory() as client:
        result = await client.call_tool(
            "get_email_link", {"message_id": 11, "scope": "thread"}
        )
    out = json.loads(result.content[0].text)
    assert out["scope"] == "thread"
    assert out["thread_size"] == 3
    assert out["newest_message_id"] == 12
    assert out["thread_link"] == "http://127.0.0.1:1/thread/11?t=x"
    # Falls back to the newest message for hosts that cannot reach the redirector.
    assert out["mail_link"] == "message://%3Cmsg12%40example.com%3E"
    # The response must not let the model claim Mail opened the whole thread.
    assert "cannot open a conversation" in out["note"]


@pytest.mark.anyio
async def test_thread_scope_skips_members_with_no_message_id(session_factory, monkeypatch):
    """An unsent draft at the end of a thread has no Message-ID yet."""
    thread = [dict(m) for m in _THREAD] + [
        _msg(13, "Brad <brad@impulselabs.com>", "Re: Contract renewal",
             "2026-09-04T08:00:00+00:00", mailbox="Drafts", message_id=""),
    ]
    bridge = FakeBridge(thread=thread)
    monkeypatch.setattr(server, "_bridge", bridge)
    monkeypatch.setattr(server, "_make_open_link", lambda mid: None)
    monkeypatch.setattr(server, "_make_thread_link", lambda mid: None)

    async with session_factory() as client:
        result = await client.call_tool(
            "get_email_link", {"message_id": 10, "scope": "thread"}
        )
    out = json.loads(result.content[0].text)
    assert out["thread_size"] == 4          # the draft still counts
    assert out["newest_message_id"] == 12   # but is not what opens


@pytest.mark.anyio
@pytest.mark.parametrize("raw", ["message", "MESSAGE", " Message ", "null", "", None])
async def test_scope_normalisation_accepts_what_hosts_actually_send(
    session_factory, fake_bridge, raw
):
    """Claude Desktop sends the string "null" for an unset parameter."""
    args = {"message_id": 11}
    if raw is not None:
        args["scope"] = raw
    async with session_factory() as client:
        result = await client.call_tool("get_email_link", args)
    assert json.loads(result.content[0].text)["scope"] == "message"


@pytest.mark.anyio
async def test_unknown_scope_is_rejected(session_factory, fake_bridge):
    async with session_factory() as client:
        result = await client.call_tool(
            "get_email_link", {"message_id": 11, "scope": "conversation"}
        )
    assert result.isError is True
    assert "conversation" in result.content[0].text
