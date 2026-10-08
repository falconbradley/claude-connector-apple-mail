"""
Tests for MailBridge construction cost and message-location resolution.

Constructing MailBridge used to prescan every mailbox, counting messages with
`mb.messages.length`. Measured on a 4-account store that was ~20s warm and
over 40s cold -- 97% of construction, against a 40s timeout, spent before the
actual tool call had started. The count was then discarded; only "is it
non-empty" was kept.

The prescan is now lazy, and HybridBridge seeds the message location from the
Envelope Index so the common paths (flag read/write) never trigger it. The
per-message lookup also uses byId (~17ms) instead of whose({id}) (5.2s on a
46k-message INBOX).
"""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from apple_mail_mcp import applescript
from apple_mail_mcp.applescript import MailBridge
from apple_mail_mcp.envelope import EnvelopeIndexBridge
from apple_mail_mcp.hybrid import HybridBridge

from test_envelope import build_store


class _ScriptRecordingBridge(MailBridge):
    """A MailBridge whose JXA calls are recorded instead of executed."""

    def __init__(self, responses=None):
        self.scripts: list[str] = []
        self.responses = list(responses or [])
        super().__init__()

    def _run_jxa(self, script, timeout=30, deadline=None):
        self.scripts.append(script)
        if self.responses:
            return self.responses.pop(0)
        return {"running": True}


# ---------------------------------------------------------------------------
# Construction no longer prescans
# ---------------------------------------------------------------------------


def test_construction_runs_exactly_one_cheap_script():
    b = _ScriptRecordingBridge()
    assert len(b.scripts) == 1
    assert "running" in b.scripts[0]


def test_construction_does_not_count_messages():
    """`mb.messages.length` is the expensive part -- it must not run here."""
    b = _ScriptRecordingBridge()
    assert "messages.length" not in b.scripts[0]
    assert "mailboxes()" not in b.scripts[0]


def test_construction_leaves_the_prescan_unevaluated():
    b = _ScriptRecordingBridge()
    assert b._nonempty_cache is None


def test_construction_raises_when_mail_is_not_running():
    class _NotRunning(_ScriptRecordingBridge):
        def _run_jxa(self, script, timeout=30, deadline=None):
            self.scripts.append(script)
            return {"running": False}

    with pytest.raises(RuntimeError, match="Mail.app is not running"):
        _NotRunning()


# ---------------------------------------------------------------------------
# The prescan is lazy and computed once
# ---------------------------------------------------------------------------


def test_prescan_runs_on_first_use_and_is_cached():
    b = _ScriptRecordingBridge(
        responses=[
            {"running": True},
            {"nonempty": [{"account": "iCloud", "mailbox": "INBOX"}]},
        ]
    )
    assert len(b.scripts) == 1
    first = b._nonempty_mailboxes()
    assert first == {("iCloud", "INBOX")}
    assert len(b.scripts) == 2
    assert "messages.length" in b.scripts[1]

    second = b._nonempty_mailboxes()
    assert second == first
    assert len(b.scripts) == 2, "second call must not re-run the scan"


def test_prescan_timeout_degrades_instead_of_raising():
    class _Timeout(_ScriptRecordingBridge):
        def _run_jxa(self, script, timeout=30, deadline=None):
            self.scripts.append(script)
            if len(self.scripts) == 1:
                return {"running": True}
            raise RuntimeError("Mail.app is not responding")

    b = _Timeout()
    assert b._nonempty_mailboxes() == set()


# ---------------------------------------------------------------------------
# Seeded locations skip the scan entirely
# ---------------------------------------------------------------------------


def test_remember_location_lets_find_message_skip_jxa():
    b = _ScriptRecordingBridge()
    b.remember_location(279676, "iCloud", "INBOX")
    before = len(b.scripts)
    assert b._find_message(279676) == ("iCloud", "INBOX", None)
    assert len(b.scripts) == before, "a seeded location must cost no JXA call"


def test_remember_location_ignores_blanks():
    b = _ScriptRecordingBridge()
    b.remember_location(1, "", "INBOX")
    b.remember_location(2, "iCloud", "")
    assert b._message_cache == {}


def test_remember_location_does_not_clobber_a_discovered_location():
    b = _ScriptRecordingBridge()
    b._message_cache[7] = ("Google", "All Mail", None)
    b.remember_location(7, "iCloud", "INBOX")
    assert b._message_cache[7] == ("Google", "All Mail", None)


# ---------------------------------------------------------------------------
# The index supplies the location
# ---------------------------------------------------------------------------


@pytest.fixture(params=["modern", "legacy"])
def store(request, tmp_path):
    build_store(tmp_path, modern=(request.param == "modern"))
    return tmp_path


def test_index_resolves_a_message_location(store):
    """The fixture's urls are the older user@host form, so the account
    display name is the address rather than a friendly name."""
    env = EnvelopeIndexBridge(mail_root=store)
    assert env.get_message_location(101) == ("brad@icloud.com", "INBOX")


def test_index_returns_none_for_an_unknown_message(store):
    env = EnvelopeIndexBridge(mail_root=store)
    assert env.get_message_location(999999) is None


def test_hybrid_seeds_the_jxa_bridge_from_the_index(store, monkeypatch):
    h = HybridBridge()
    monkeypatch.setattr(h, "_envelope", lambda: EnvelopeIndexBridge(mail_root=store))
    fake = _ScriptRecordingBridge()
    h._mail = fake

    h._jxa(for_message=101)
    assert fake._message_cache[101] == ("brad@icloud.com", "INBOX", None)

    before = len(fake.scripts)
    assert fake._find_message(101) == ("brad@icloud.com", "INBOX", None)
    assert len(fake.scripts) == before


def test_hybrid_survives_an_index_that_cannot_answer(monkeypatch):
    """No index (no Full Disk Access) must not break the JXA path."""
    h = HybridBridge()
    monkeypatch.setattr(h, "_envelope", lambda: None)
    fake = _ScriptRecordingBridge()
    h._mail = fake
    h._jxa(for_message=101)  # must not raise
    assert fake._message_cache == {}


# ---------------------------------------------------------------------------
# byId for targeted lookups, whose() for the membership scan
# ---------------------------------------------------------------------------


def test_targeted_lookups_use_byid():
    """byId is ~300x faster than whose({id}) on a large mailbox."""
    b = _ScriptRecordingBridge()
    b.remember_location(7, "iCloud", "INBOX")
    b.set_flag(7, "green")
    script = b.scripts[-1]
    assert "messages.byId(7)" in script
    assert "messages.whose({id: 7})" not in script


def test_the_membership_scan_still_uses_whose():
    """byId resolves app-wide, so it cannot answer 'is it in this mailbox'."""
    b = _ScriptRecordingBridge(
        responses=[
            {"running": True},
            {"nonempty": [{"account": "iCloud", "mailbox": "INBOX"}]},
            None,
        ]
    )
    b._find_message(4242)
    scan = b.scripts[-1]
    assert "messages.whose({id: 4242})" in scan
    # the word appears in the comment explaining why; the call must not
    assert "messages.byId(" not in scan


def test_search_shares_its_time_budget_with_the_prescan():
    """The prescan now runs inside the tool call, so it must draw on the same
    budget rather than spending its own ceiling on top of the search's."""
    seen: list[object] = []

    class _Budgeted(_ScriptRecordingBridge):
        def _run_jxa(self, script, timeout=30, deadline=None):
            self.scripts.append(script)
            if len(self.scripts) == 1:
                return {"running": True}
            if "messages.length" in script:
                seen.append(deadline)
                return {"nonempty": [{"account": "iCloud", "mailbox": "INBOX"}]}
            return {"total": 0, "messages": []}

    b = _Budgeted()
    b.search_messages(subject_contains="invoice")
    assert seen, "the prescan did not run"
    assert seen[0] is not None, "the prescan ran without the search's deadline"
    assert isinstance(seen[0], applescript._Deadline)
