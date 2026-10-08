"""Quick reply / quick forward: the send_email_reply and send_email_forward tools.

Nothing here talks to Mail.app. The server-layer tests swap in a recording
bridge; the script-layer tests inspect (and compile) the AppleScript that
MailBridge would run, without running it.
"""

from __future__ import annotations

import asyncio
import shutil
import subprocess
import pytest

from apple_mail_mcp import server
from apple_mail_mcp.applescript import (
    MailBridge,
    _as_escape,
    _parse_send_report,
    _strip_ws,
)

needs_osascript = pytest.mark.skipif(
    shutil.which("osascript") is None, reason="needs macOS osascript"
)


# ---------------------------------------------------------------------------
# Server layer
# ---------------------------------------------------------------------------


class _SendBridge:
    """Records send calls and replays a canned bridge report."""

    def __init__(self, report: dict | None = None, sent_copy: dict | None = None):
        self.calls: list[tuple[str, tuple, dict]] = []
        self.report = report or {
            "sent": True,
            "reason": None,
            "outbox": "left",
            "subject": "Re: Hello",
            "from": "Me <me@example.com>",
            "to_addresses": ["them@example.com"],
            "cc_addresses": [],
            "bcc_addresses": [],
        }
        self.sent_copy = sent_copy

    def send_reply(self, *args, **kwargs):
        self.calls.append(("reply", args, kwargs))
        return dict(self.report)

    def send_forward(self, *args, **kwargs):
        self.calls.append(("forward", args, kwargs))
        return dict(self.report)

    def find_sent_copy(self, subject, sent_after, **kwargs):
        return self.sent_copy


@pytest.fixture
def enabled(monkeypatch):
    monkeypatch.setenv(server.SENDING_ENV, "true")


@pytest.fixture
def bridge(monkeypatch):
    b = _SendBridge(sent_copy={"id": 42, "message_id": "abc@example.com"})
    monkeypatch.setattr(server, "_bridge", b)
    return b


@pytest.mark.parametrize("value", [None, "", "false", "0", "no", "${user_config.enable_sending}"])
def test_sending_is_off_unless_turned_on(monkeypatch, bridge, value):
    if value is None:
        monkeypatch.delenv(server.SENDING_ENV, raising=False)
    else:
        monkeypatch.setenv(server.SENDING_ENV, value)
    with pytest.raises(PermissionError, match="Allow sending email"):
        server.send_email_reply(message_id=1, body="hi")
    with pytest.raises(PermissionError, match="Allow sending email"):
        server.send_email_forward(message_id=1, to=["a@example.com"])
    assert bridge.calls == []


@pytest.mark.parametrize("value", ["true", "TRUE", "1", "yes", "on"])
def test_truthy_settings_turn_sending_on(monkeypatch, bridge, value):
    monkeypatch.setenv(server.SENDING_ENV, value)
    server.send_email_reply(message_id=1, body="hi")
    assert len(bridge.calls) == 1


def test_both_tools_are_marked_destructive_and_open_world():
    tools = {t.name: t for t in asyncio.run(server.mcp.list_tools())}
    for name in ("send_email_reply", "send_email_forward"):
        ann = tools[name].annotations
        assert ann.destructiveHint is True
        assert ann.openWorldHint is True
        assert ann.readOnlyHint is False
        assert ann.idempotentHint is False


def test_reply_requires_a_body(enabled, bridge):
    with pytest.raises(ValueError):
        server.send_email_reply(message_id=1, body="   ")
    assert bridge.calls == []


def test_forward_requires_a_recipient(enabled, bridge):
    with pytest.raises(ValueError):
        server.send_email_forward(message_id=1, to=[])
    with pytest.raises(ValueError):
        server.send_email_forward(message_id=1, to=["  "])
    assert bridge.calls == []


def test_forward_rejects_something_that_is_not_an_address(enabled, bridge):
    with pytest.raises(ValueError, match="not an email address"):
        server.send_email_forward(message_id=1, to=["Sam Smith"])
    assert bridge.calls == []


def test_forward_note_null_string_means_no_note(enabled, bridge):
    server.send_email_forward(message_id=7, to=["a@example.com"], body="null")
    kind, args, kwargs = bridge.calls[0]
    assert kind == "forward"
    assert args == (7, ["a@example.com"])
    assert kwargs["body"] is None


def test_reply_passes_its_options_through(enabled, bridge):
    server.send_email_reply(
        message_id=9, body="Thanks!", reply_all=True, cc=["c@example.com"],
        include_quoted=False,
    )
    kind, args, kwargs = bridge.calls[0]
    assert (kind, args) == ("reply", (9, "Thanks!"))
    assert kwargs["reply_all"] is True
    assert kwargs["cc_addresses"] == ["c@example.com"]
    assert kwargs["include_quoted"] is False


def test_confirmed_send_returns_the_sent_copy(enabled, bridge):
    r = server.send_email_reply(message_id=1, body="hi")
    assert r.status == "sent"
    assert r.sent_message_id == 42
    assert r.mail_link and "abc%40example.com" in r.mail_link
    assert r.to_addresses == ["them@example.com"]


def test_sent_but_not_yet_in_sent_is_reported_honestly(enabled, monkeypatch):
    monkeypatch.setattr(server, "_bridge", _SendBridge(sent_copy=None))
    r = server.send_email_reply(message_id=1, body="hi")
    assert r.status == "sent_unconfirmed"
    assert r.sent_message_id is None


def test_stuck_in_outbox_is_not_reported_as_sent(enabled, monkeypatch):
    b = _SendBridge()
    b.report["outbox"] = "still_queued"
    monkeypatch.setattr(server, "_bridge", b)
    r = server.send_email_reply(message_id=1, body="hi")
    assert r.status == "queued_in_outbox"
    assert "Outbox" in r.note


def test_failed_verification_raises_and_says_it_was_kept(enabled, monkeypatch):
    b = _SendBridge()
    b.report.update(sent=False, reason="body_not_in_message", outbox=None)
    monkeypatch.setattr(server, "_bridge", b)
    with pytest.raises(RuntimeError) as exc:
        server.send_email_reply(message_id=1, body="hi")
    assert "Not sent" in str(exc.value)
    assert "left open in Mail" in str(exc.value)


def test_unknown_outcome_warns_against_retrying(enabled, monkeypatch):
    b = _SendBridge()
    b.report.update(sent=None, reason="script_failed_or_timed_out")
    monkeypatch.setattr(server, "_bridge", b)
    with pytest.raises(RuntimeError, match="Do NOT retry"):
        server.send_email_forward(message_id=1, to=["a@example.com"])


# ---------------------------------------------------------------------------
# Report parsing and whitespace normalisation
# ---------------------------------------------------------------------------


def test_parse_sent_report():
    r = _parse_send_report(
        "SENT\n\nSUBJECT:Fwd: Hi\nFROM:Me <me@x.com>\nTO:a@x.com\nCC:b@x.com\n"
        "BCC:c@x.com\nOUTBOX:left"
    )
    assert r["sent"] is True
    assert r["outbox"] == "left"
    assert r["subject"] == "Fwd: Hi"
    assert (r["to_addresses"], r["cc_addresses"], r["bcc_addresses"]) == (
        ["a@x.com"], ["b@x.com"], ["c@x.com"]
    )


def test_parse_not_sent_report():
    r = _parse_send_report("NOTSENT\nREASON:to_recipients_mismatch\nSUBJECT:Fwd: Hi")
    assert r["sent"] is False
    assert r["reason"] == "to_recipients_mismatch"


def test_parse_error_report():
    r = _parse_send_report("ERROR:message_not_found")
    assert r["sent"] is False
    assert r["reason"] == "message_not_found"


def test_parse_garbage_is_not_sent():
    r = _parse_send_report("something unexpected")
    assert r["sent"] is False
    assert r["reason"].startswith("unexpected_output")


def test_strip_ws_removes_the_whitespace_mail_rewrites():
    assert _strip_ws(" a\tb\r\nc\u00a0d\u2028e\u2029f ") == "abcdef"


# ---------------------------------------------------------------------------
# Script layer
# ---------------------------------------------------------------------------


def _script_for(kind: str, **kwargs) -> str:
    b = MailBridge.__new__(MailBridge)
    b._find_message = lambda mid: ("iCloud", "INBOX", None)
    seen: dict[str, str] = {}

    def capture(script, timeout=30):
        seen["script"] = script
        return "NOTSENT\nREASON:test"

    b._run_applescript = capture
    if kind == "reply":
        b.send_reply(123, kwargs.pop("body", "Thanks"), **kwargs)
    else:
        b.send_forward(123, kwargs.pop("to", ["a@example.com"]), **kwargs)
    return seen["script"]


def test_every_check_comes_before_the_send():
    script = _script_for("forward", body="FYI")
    send_at = script.index("set ok to send o")
    for check in (
        'leaveOpen(o, "body_not_in_message")',
        'leaveOpen(o, "to_recipients_mismatch")',
        'leaveOpen(o, "no_recipients")',
        "my focusBody(outSubj)",
        "save o",
    ):
        assert script.index(check) < send_at, check
    assert script.count("send o\n") == 1


def test_forward_checks_recipients_exactly_and_reply_does_not():
    fwd = _script_for("forward", to=["Sam <sam@example.com>"])
    assert '{"sam@example.com"}, true)' in fwd
    reply = _script_for("reply")
    assert "{}, false)" in reply


def test_paste_only_happens_into_the_compose_window():
    script = _script_for("reply")
    assert "if (name of w) is not expectedTitle" in script


def test_drafts_are_refused_before_any_script_runs():
    b = MailBridge.__new__(MailBridge)
    b._find_message = lambda mid: ("iCloud", "Drafts", None)
    b._run_applescript = lambda *a, **k: pytest.fail("must not run")
    with pytest.raises(ValueError, match="draft"):
        b.send_reply(1, "hi")


def test_timeout_is_reported_as_unknown_not_as_failure():
    b = MailBridge.__new__(MailBridge)
    b._find_message = lambda mid: ("iCloud", "INBOX", None)
    b._run_applescript = lambda *a, **k: None
    assert b.send_forward(1, ["a@example.com"])["sent"] is None


@needs_osascript
@pytest.mark.parametrize("kind", ["reply", "forward"])
def test_generated_scripts_compile(tmp_path, kind):
    body = 'He said "hi" \\ then\nleft'
    script = _script_for(kind, body=body, cc_addresses=['O"Neil <o@example.com>'])
    src = tmp_path / "s.applescript"
    src.write_text(script)
    r = subprocess.run(
        ["osacompile", "-o", str(tmp_path / "s.scpt"), str(src)],
        capture_output=True, text=True,
    )
    assert r.returncode == 0, r.stderr


@needs_osascript
def test_as_escape_round_trips_quotes_and_backslashes():
    value = 'He said "hi" \\ C:\\x ""'
    r = subprocess.run(
        ["osascript", "-e", f'return "{_as_escape(value)}"'],
        capture_output=True, text=True,
    )
    assert r.returncode == 0, r.stderr
    assert r.stdout.rstrip("\n") == value
