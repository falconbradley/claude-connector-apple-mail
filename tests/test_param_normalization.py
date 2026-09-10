"""
Claude Desktop sometimes sends the *string* "null" for an optional parameter
it means to leave unset. Seen in the server log on 2026-09-10:

    tools/call search_emails {'query': 'null', 'from_address': 'parentsquare.com', ...}

Before the fix that filtered subjects for the literal word "null". These
tests pin the normalisation for every optional string filter, and for the
flag argument where "null" is the documented way to say "remove".
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from apple_mail_mcp import server  # noqa: E402
from apple_mail_mcp.server import _optional  # noqa: E402


class _RecordingBridge:
    """Captures the kwargs the tool hands the bridge; returns nothing."""

    def __init__(self) -> None:
        self.search_kwargs: dict | None = None
        self.flag_args: tuple | None = None
        self.last_engine = "sqlite"

    def search_messages(self, **kwargs):
        self.search_kwargs = kwargs
        return 0, []

    def set_flag(self, message_id, flag=None):
        self.flag_args = (message_id, flag)
        return {"success": True, "color_index": -1}


@pytest.fixture
def bridge(monkeypatch):
    b = _RecordingBridge()
    monkeypatch.setattr(server, "_bridge", b)
    return b


@pytest.mark.parametrize("raw", [None, "", "   ", "null", "NULL", "Null", "None", "none"])
def test_optional_treats_unset_spellings_as_none(raw):
    assert _optional(raw) is None


@pytest.mark.parametrize("raw", ["nullable", "none of the above", "0", "INBOX", " x "])
def test_optional_keeps_real_values(raw):
    assert _optional(raw) == raw


def test_search_with_string_null_query_applies_no_text_filter(bridge):
    server.search_emails(query="null", from_address="parentsquare.com")
    kw = bridge.search_kwargs
    # from_address still wins for the sender; the bogus query must not leak
    # into the subject filter.
    assert kw["subject_contains"] is None
    assert kw["sender_contains"] == "parentsquare.com"


def test_search_normalises_every_optional_filter(bridge):
    server.search_emails(
        query="None", mailbox="null", account="", from_address="  ",
        to_address="NULL", subject="none", since="null", before="None",
    )
    kw = bridge.search_kwargs
    for key in (
        "mailbox_name", "account_name", "subject_contains", "sender_contains",
        "to_address_contains", "since", "before",
    ):
        assert kw[key] is None, key


def test_search_real_values_pass_through_unchanged(bridge):
    server.search_emails(query="invoice", mailbox="INBOX", since="2026-09-01")
    kw = bridge.search_kwargs
    assert kw["subject_contains"] == "invoice"
    assert kw["sender_contains"] == "invoice"
    assert kw["mailbox_name"] == "INBOX"
    assert kw["since"] is not None


def test_set_email_flag_with_string_null_removes_the_flag(bridge):
    result = server.set_email_flag(7, "null")
    assert bridge.flag_args == (7, None)
    assert result.flag_color is None and result.success is True


def test_set_email_flag_still_rejects_unknown_colours(bridge):
    with pytest.raises(ValueError, match="chartreuse"):
        server.set_email_flag(7, "chartreuse")
