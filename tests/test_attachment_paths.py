"""
Attachment filenames are sender-controlled (Content-Disposition filename=)
and must never influence where bytes are written on disk. Regression tests
for the path traversal reported in issue #12.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from apple_mail_mcp.applescript import MailBridge
from apple_mail_mcp.emlx import safe_attachment_filename
from apple_mail_mcp.envelope import EnvelopeIndexBridge

from test_envelope import MULTIPART_MESSAGE, build_store, write_emlx

EVIL_NAME = "../../../../Library/LaunchAgents/evil.plist"


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("report.pdf", "report.pdf"),
        (EVIL_NAME, "evil.plist"),
        ("..\\..\\Windows\\evil.bat", "evil.bat"),
        ("/etc/passwd", "passwd"),
        ("..", "fallback"),
        ("../", "fallback"),
        (".bashrc", "bashrc"),
        ("a\x00b\nc:d", "a_b_c_d"),
        ("", "fallback"),
        (None, "fallback"),
    ],
)
def test_safe_attachment_filename(raw, expected):
    assert safe_attachment_filename(raw, "fallback") == expected


# ---------------------------------------------------------------------------
# Envelope Index bridge (read-only path)
# ---------------------------------------------------------------------------

def test_envelope_returns_sanitized_name(tmp_path):
    build_store(tmp_path)
    evil = MULTIPART_MESSAGE.replace(
        b"filename=report.pdf", f'filename="{EVIL_NAME}"'.encode()
    )
    messages_dir = (
        tmp_path / "V10" / "ACCT-UUID" / "INBOX.mbox" / "MBOX-UUID" / "Data"
        / "1" / "Messages"
    )
    write_emlx(messages_dir / "106.partial.emlx", evil)
    bridge = EnvelopeIndexBridge(mail_root=tmp_path)

    assert bridge.list_attachments(106)[0]["name"] == "evil.plist"
    name, _, data = bridge.get_attachment(106, 0)
    assert name == "evil.plist"
    assert data.startswith(b"%PDF")


# ---------------------------------------------------------------------------
# JXA bridge (the path that actually writes to disk)
# ---------------------------------------------------------------------------

def _jxa_bridge(monkeypatch, fake_run_jxa) -> MailBridge:
    bridge = MailBridge.__new__(MailBridge)  # skip the Mail.app probe
    bridge._message_cache = {}
    bridge._nonempty_mailboxes = set()
    monkeypatch.setattr(bridge, "_find_message", lambda _id: ("iCloud", "INBOX", None))
    monkeypatch.setattr(bridge, "_run_jxa", fake_run_jxa)
    return bridge


def _save_path(script: str) -> Path:
    match = re.search(r'mail\.save\(att, \{in: Path\("([^"]+)"\)\}\)', script)
    assert match, "save call not found in script"
    return Path(match.group(1))


def test_jxa_save_path_ignores_sender_filename(monkeypatch):
    seen = {}

    def fake_run_jxa(script, timeout=30):
        seen["script"] = script
        path = _save_path(script)
        path.write_bytes(b"%PDF-1.4 via jxa")
        seen["path"] = path
        return {"filename": EVIL_NAME, "mime_type": "application/pdf"}

    bridge = _jxa_bridge(monkeypatch, fake_run_jxa)
    name, mime, data = bridge.get_attachment(42, 0)

    # The save target is built without concatenating att.name()
    assert "+ fileName" not in seen["script"]
    assert seen["path"].name == "attachment"
    assert seen["path"].parent.name.startswith("apple_mail_att_")
    assert name == "evil.plist"
    assert mime == "application/pdf"
    assert data == b"%PDF-1.4 via jxa"


def test_jxa_refuses_to_follow_symlink_out_of_tmpdir(monkeypatch, tmp_path):
    secret = tmp_path / "secret.txt"
    secret.write_bytes(b"do not read")

    def fake_run_jxa(script, timeout=30):
        _save_path(script).symlink_to(secret)
        return {"filename": "x.pdf", "mime_type": "application/pdf"}

    bridge = _jxa_bridge(monkeypatch, fake_run_jxa)
    assert bridge.get_attachment(42, 0) is None


def test_jxa_list_attachments_sanitizes_names(monkeypatch):
    def fake_run_jxa(script, timeout=30):
        return [{"index": 0, "name": EVIL_NAME, "mime_type": "x", "file_size": 1}]

    bridge = _jxa_bridge(monkeypatch, fake_run_jxa)
    assert bridge.list_attachments(42)[0]["name"] == "evil.plist"
