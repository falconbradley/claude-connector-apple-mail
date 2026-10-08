"""
Regression tests for the two flag-writing defects found on 2026-09-20 while a
scheduled inbox-review run was setting flags on live mail.

Bug 1 -- set_email_flag(flag="green") reported green but get_email_flag read
back orange. The write path (JXA `flagIndex`) was correct all along; the fast
read path decoded the colour from the Envelope Index `messages.flag_color`
column, which is NOT a colour index -- it only ever holds 0 or 1, so every
flagged message came back as "red" or "orange". The colour actually lives in
bits 39-41 of `messages.flags`, using the same 0-6 encoding as `flagIndex`.
See docs/flag-index.md for the ground-truth probe.

Bug 2 -- set_email_flag(message_id=..., flag_color="green") had its unknown
kwarg silently dropped by the MCP layer, `flag` fell back to its None default,
and the call cleared a live orange flag while reporting success.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from apple_mail_mcp import server
from apple_mail_mcp.applescript import _FLAG_COLOR_MAP, _FLAG_COLOR_ORDER
from apple_mail_mcp.envelope import (
    _FLAG_COLOR_MASK,
    _FLAG_COLOR_SHIFT,
    EnvelopeIndexBridge,
)

from test_envelope import build_store

COLORS = ["red", "orange", "yellow", "green", "blue", "purple", "gray"]

# ---------------------------------------------------------------------------
# Bug 1: the fast read path must decode the colour from the flags bitfield
# ---------------------------------------------------------------------------


def _set_raw(root: Path, rowid: int, *, flags: int, flag_color=None) -> None:
    """Write a raw flags bitfield (and legacy flag_color) for one message."""
    db = sqlite3.connect(root / "V10" / "MailData" / "Envelope Index")
    cols = {r[1] for r in db.execute("PRAGMA table_info(messages)")}
    db.execute("UPDATE messages SET flags = ? WHERE ROWID = ?", (flags, rowid))
    if "flagged" in cols:
        db.execute(
            "UPDATE messages SET flagged = ? WHERE ROWID = ?",
            (1 if flags & (1 << 4) else 0, rowid),
        )
    if flag_color is not None and "flag_color" in cols:
        db.execute(
            "UPDATE messages SET flag_color = ? WHERE ROWID = ?", (flag_color, rowid)
        )
    db.commit()
    db.close()


def _flags_for(color_index: int, *, flagged: bool = True) -> int:
    value = 1  # read bit, so the row looks like ordinary mail
    if flagged:
        value |= 1 << 4
    return value | (color_index << _FLAG_COLOR_SHIFT)


@pytest.fixture(params=["modern", "legacy"])
def store(request, tmp_path):
    build_store(tmp_path, modern=(request.param == "modern"))
    return tmp_path


@pytest.mark.parametrize("color_index,name", list(enumerate(COLORS)))
def test_every_flag_colour_round_trips_through_the_index(store, color_index, name):
    """Bits 39-41 decode to the same 0-6 order Mail.app's flagIndex uses."""
    _set_raw(store, 103, flags=_flags_for(color_index))
    result = EnvelopeIndexBridge(mail_root=store).get_flag(103)
    assert result["is_flagged"] is True
    assert result["color_index"] == color_index
    assert _FLAG_COLOR_ORDER[result["color_index"]] == name


def test_green_does_not_read_back_as_orange(store):
    """The exact reported symptom: write green, read green -- not orange."""
    _set_raw(store, 103, flags=_flags_for(_FLAG_COLOR_MAP["green"]))
    result = EnvelopeIndexBridge(mail_root=store).get_flag(103)
    assert _FLAG_COLOR_ORDER[result["color_index"]] == "green"


def test_legacy_flag_color_column_is_ignored(store):
    """A flag_color column disagreeing with the bitfield must not win.

    On the real store flag_color was 1 ("orange") for messages that were
    genuinely green, red and gray -- 17 of 22 flagged messages were misread.
    """
    db = sqlite3.connect(store / "V10" / "MailData" / "Envelope Index")
    cols = {r[1] for r in db.execute("PRAGMA table_info(messages)")}
    if "flag_color" not in cols:
        db.execute("ALTER TABLE messages ADD COLUMN flag_color INTEGER")
    db.commit()
    db.close()

    _set_raw(store, 103, flags=_flags_for(_FLAG_COLOR_MAP["green"]), flag_color=1)
    result = EnvelopeIndexBridge(mail_root=store).get_flag(103)
    assert _FLAG_COLOR_ORDER[result["color_index"]] == "green"


def test_colour_bits_are_ignored_when_the_message_is_unflagged(store):
    """Mail leaves the colour bits set after unflagging; they are stale."""
    _set_raw(store, 103, flags=_flags_for(_FLAG_COLOR_MAP["purple"], flagged=False))
    result = EnvelopeIndexBridge(mail_root=store).get_flag(103)
    assert result["is_flagged"] is False
    assert result["color_index"] == -1


def test_colour_mask_covers_exactly_three_bits():
    assert _FLAG_COLOR_MASK == 0x7 << _FLAG_COLOR_SHIFT
    assert (_FLAG_COLOR_MASK >> _FLAG_COLOR_SHIFT) == 0b111


def test_hybrid_read_path_reports_the_bitfield_colour(store, monkeypatch):
    """End-to-end through HybridBridge: the sqlite fast path, no JXA."""
    from apple_mail_mcp.hybrid import HybridBridge

    _set_raw(store, 103, flags=_flags_for(_FLAG_COLOR_MAP["green"]))
    h = HybridBridge()
    monkeypatch.setattr(h, "_envelope", lambda: EnvelopeIndexBridge(mail_root=store))

    def _no_jxa():
        raise AssertionError("fast path must answer without falling back to JXA")

    monkeypatch.setattr(h, "_jxa", _no_jxa)
    assert h.get_flag(103)["flag_color"] == "green"


# ---------------------------------------------------------------------------
# Bug 2: an unknown parameter must not silently remove the flag
# ---------------------------------------------------------------------------


class _RecordingBridge:
    """Stands in for the real bridge; echoes back what Mail.app would store."""

    def __init__(self, *, lies: bool = False) -> None:
        self.calls: list[tuple[int, object]] = []
        self.lies = lies

    def set_flag(self, message_id, flag=None):
        self.calls.append((message_id, flag))
        if self.lies:
            # A write that did not take: the script ran, nothing changed.
            return {"success": True, "is_flagged": False, "color_index": -1}
        if flag is None:
            return {"success": True, "is_flagged": False, "color_index": -1}
        return {
            "success": True,
            "is_flagged": True,
            "color_index": _FLAG_COLOR_MAP[flag],
        }


@pytest.fixture
def bridge(monkeypatch):
    b = _RecordingBridge()
    monkeypatch.setattr(server, "_bridge", b)
    return b


def test_flag_color_alias_sets_the_colour(bridge):
    """The reported call. It must set green, not clear the flag."""
    result = server.set_email_flag(279676, flag_color="green")
    assert bridge.calls == [(279676, "green")]
    assert result.flag_color == "green" and result.success is True


def test_omitting_the_flag_entirely_raises_and_writes_nothing(bridge):
    """A mistyped parameter name lands here -- it must not unflag anything."""
    with pytest.raises(ValueError, match="requires a flag"):
        server.set_email_flag(279676)
    assert bridge.calls == []


def test_explicit_none_still_removes_the_flag(bridge):
    result = server.set_email_flag(279676, None)
    assert bridge.calls == [(279676, None)]
    assert result.flag_color is None and result.success is True


def test_explicit_none_via_the_alias_also_removes(bridge):
    result = server.set_email_flag(279676, flag_color=None)
    assert bridge.calls == [(279676, None)]
    assert result.flag_color is None


def test_conflicting_aliases_raise(bridge):
    with pytest.raises(ValueError, match="Conflicting arguments"):
        server.set_email_flag(279676, flag="red", flag_color="blue")
    assert bridge.calls == []


def test_matching_aliases_are_accepted(bridge):
    result = server.set_email_flag(279676, flag="blue", flag_color="blue")
    assert bridge.calls == [(279676, "blue")]
    assert result.flag_color == "blue"


def test_alias_rejects_unknown_colours(bridge):
    with pytest.raises(ValueError, match="chartreuse"):
        server.set_email_flag(279676, flag_color="chartreuse")
    assert bridge.calls == []


@pytest.mark.parametrize("spelling", ["null", "NULL", "none", ""])
def test_unset_spellings_still_remove_the_flag(bridge, spelling):
    """Pinned by tests/test_param_normalization.py -- keep it working."""
    result = server.set_email_flag(279676, spelling)
    assert bridge.calls == [(279676, None)]
    assert result.flag_color is None


def test_a_write_that_did_nothing_is_not_reported_as_success(monkeypatch):
    """set_flag's JXA returns success whenever the *script* ran."""
    b = _RecordingBridge(lies=True)
    monkeypatch.setattr(server, "_bridge", b)
    with pytest.raises(RuntimeError, match="Failed to set the 'green' flag"):
        server.set_email_flag(279676, "green")


def test_a_write_that_stored_the_wrong_colour_is_reported(monkeypatch):
    """The JXA fallback sets flaggedStatus=true, which is always red."""

    class _FallsBackToRed(_RecordingBridge):
        def set_flag(self, message_id, flag=None):
            self.calls.append((message_id, flag))
            return {"success": True, "is_flagged": True, "color_index": 0}

    b = _FallsBackToRed()
    monkeypatch.setattr(server, "_bridge", b)
    with pytest.raises(RuntimeError, match="stored red instead"):
        server.set_email_flag(279676, "green")


def test_a_failed_removal_is_reported(monkeypatch):
    class _KeepsTheFlag(_RecordingBridge):
        def set_flag(self, message_id, flag=None):
            self.calls.append((message_id, flag))
            return {"success": True, "is_flagged": True, "color_index": 1}

    b = _KeepsTheFlag()
    monkeypatch.setattr(server, "_bridge", b)
    with pytest.raises(RuntimeError, match="still reports it flagged"):
        server.set_email_flag(279676, None)
