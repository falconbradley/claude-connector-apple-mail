"""Tests for the localhost web-link redirector."""

from __future__ import annotations

import json
import urllib.error
import urllib.request

import pytest

from apple_mail_mcp.weblink import ThreadAnchor, WebLinkServer


@pytest.fixture
def server(tmp_path):
    opened: list[str] = []

    def resolve(mid: int):
        return {101: "invoice-101@example.com"}.get(mid)

    def resolve_thread(mid: int):
        # 101 and 102 are two messages of one three-message conversation;
        # 103 is a thread of one.
        if mid in (101, 102):
            return ThreadAnchor(message_id=102, rfc_id="reply-102@example.com", total=3)
        if mid == 103:
            return ThreadAnchor(message_id=103, rfc_id="alone-103@example.com", total=1)
        return None

    srv = WebLinkServer(
        resolve_rfc_id=resolve,
        state_path=tmp_path / "weblink.json",
        opener=lambda link: opened.append(link) or True,
        preferred_port=0,  # ephemeral for tests
        resolve_thread=resolve_thread,
    )
    srv.opened = opened
    yield srv
    srv.shutdown()


def _get(url: str) -> tuple[int, bytes]:
    try:
        with urllib.request.urlopen(url, timeout=5) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read()


def test_open_link_roundtrip(server):
    link = server.open_link(101)
    assert link.startswith("http://127.0.0.1:")
    status, body = _get(link)
    assert status == 200
    assert b"Opened in Mail" in body
    assert server.opened == [
        "message://%3Cinvoice-101%40example.com%3E"
    ]


def test_success_page_closes_its_own_tab(server):
    """Every click opens a browser tab; on success it must not linger."""
    _, body = _get(server.open_link(101))
    assert b"window.close()" in body
    _, body = _get(server.thread_link(101))
    assert b"window.close()" in body


def test_success_page_still_explains_itself_if_the_close_is_refused(server):
    """Safari and hand-navigated tabs refuse window.close(); the page a user
    is then left looking at has to make sense on its own."""
    _, body = _get(server.open_link(101))
    assert b"Opened in Mail" in body
    assert b"close this tab" in body


def test_error_pages_never_close_themselves(server):
    """An error page exists to be read."""
    for url in (
        server.open_link(999),                                   # 404
        server.thread_link(999),                                 # 404
        server.open_link(101).split("?t=")[0] + "?t=bad",         # 403
    ):
        status, body = _get(url)
        assert status in (403, 404)
        assert b"window.close()" not in body


def test_page_is_not_closed_when_mail_refuses_to_open(tmp_path):
    """A 500 tells the user to try the message:// link by hand -- closing the
    tab would take that link away with it."""
    srv = WebLinkServer(
        resolve_rfc_id=lambda mid: "x@example.com",
        state_path=tmp_path / "weblink.json",
        opener=lambda link: False,          # macOS `open` failed
        preferred_port=0,
    )
    try:
        status, body = _get(srv.open_link(101))
        assert status == 500
        assert b"window.close()" not in body
        assert b"message://" in body
    finally:
        srv.shutdown()


def test_thread_link_opens_the_newest_message(server):
    link = server.thread_link(101)
    assert "/thread/101" in link
    status, body = _get(link)
    assert status == 200
    assert b"Opened thread in Mail" in body
    # Asked about 101, opened 102 -- the conversation's newest message.
    assert server.opened == ["message://%3Creply-102%40example.com%3E"]


def test_thread_link_says_how_many_messages_it_did_not_open(server):
    status, body = _get(server.thread_link(101))
    assert status == 200
    assert b"3 messages" in body
    # The page must not overclaim: Mail only ever got one message.
    assert b"one message at" in body


def test_thread_link_on_a_lone_message_says_so(server):
    status, body = _get(server.thread_link(103))
    assert status == 200
    assert b"just the one message" in body
    assert server.opened == ["message://%3Calone-103%40example.com%3E"]


def test_thread_link_resolves_at_click_time_not_link_time(server):
    """Two members of one conversation produce different links that open the
    same (newest) message -- so an old link follows the thread as it grows."""
    a, b = server.thread_link(101), server.thread_link(102)
    assert a != b
    _get(a)
    _get(b)
    assert server.opened == ["message://%3Creply-102%40example.com%3E"] * 2


def test_unknown_thread_404s(server):
    status, _ = _get(server.thread_link(999))
    assert status == 404
    assert server.opened == []


def test_thread_link_requires_the_token(server):
    link = server.thread_link(101)
    status, _ = _get(link.split("?t=")[0] + "?t=wrong-token")
    assert status == 403
    assert server.opened == []


def test_thread_link_is_none_without_a_resolver(tmp_path):
    srv = WebLinkServer(
        resolve_rfc_id=lambda mid: "x@example.com",
        state_path=tmp_path / "weblink.json",
        opener=lambda link: True,
        preferred_port=0,
    )
    try:
        assert srv.thread_link(101) is None
        assert srv.open_link(101) is not None
    finally:
        srv.shutdown()


def test_unknown_message_404s(server):
    link = server.open_link(999)
    status, body = _get(link)
    assert status == 404
    assert server.opened == []


def test_bad_token_403s(server):
    link = server.open_link(101)
    status, _ = _get(link.split("?t=")[0] + "?t=wrong-token")
    assert status == 403
    status, _ = _get(link.split("?t=")[0])  # no token at all
    assert status == 403
    assert server.opened == []


def test_bad_path_404s(server):
    server.open_link(101)
    base = f"http://127.0.0.1:{server.port}"
    status, _ = _get(f"{base}/other?t={server.token}")
    assert status == 404
    status, _ = _get(f"{base}/open/notanumber?t={server.token}")
    assert status == 404


def test_ping(server):
    server.ensure_started()
    status, body = _get(
        f"http://127.0.0.1:{server.port}/ping?t={server.token}"
    )
    assert status == 200
    assert body.strip() == b"apple-mail-mcp-weblink"


def test_token_persisted_across_instances(tmp_path):
    state = tmp_path / "weblink.json"
    srv1 = WebLinkServer(
        resolve_rfc_id=lambda mid: None,
        state_path=state,
        opener=lambda link: True,
        preferred_port=0,
    )
    link1 = srv1.open_link(1)
    saved = json.loads(state.read_text())
    assert saved["token"] in link1
    assert saved["port"] == srv1.port
    srv1.shutdown()

    srv2 = WebLinkServer(
        resolve_rfc_id=lambda mid: None,
        state_path=state,
        opener=lambda link: True,
        preferred_port=0,
    )
    link2 = srv2.open_link(1)
    assert saved["token"] in link2  # same token — old links stay valid
    srv2.shutdown()


def test_sibling_port_reuse(tmp_path):
    """A second instance finding the port taken by a live sibling reuses
    the sibling's port in generated links instead of binding a new one."""
    state = tmp_path / "weblink.json"
    srv1 = WebLinkServer(
        resolve_rfc_id=lambda mid: "a@b.c",
        state_path=state,
        opener=lambda link: True,
        preferred_port=0,
    )
    assert srv1.ensure_started()

    srv2 = WebLinkServer(
        resolve_rfc_id=lambda mid: "a@b.c",
        state_path=state,
        opener=lambda link: True,
        preferred_port=srv1.port,  # state also points here
    )
    link = srv2.open_link(101)
    assert f":{srv1.port}/" in link
    assert srv2._httpd is None  # did not bind its own server
    srv1.shutdown()
