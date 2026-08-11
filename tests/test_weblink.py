"""Tests for the localhost web-link redirector."""

from __future__ import annotations

import json
import urllib.error
import urllib.request

import pytest

from apple_mail_mcp.weblink import WebLinkServer


@pytest.fixture
def server(tmp_path):
    opened: list[str] = []

    def resolve(mid: int):
        return {101: "invoice-101@example.com"}.get(mid)

    srv = WebLinkServer(
        resolve_rfc_id=resolve,
        state_path=tmp_path / "weblink.json",
        opener=lambda link: opened.append(link) or True,
        preferred_port=0,  # ephemeral for tests
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
