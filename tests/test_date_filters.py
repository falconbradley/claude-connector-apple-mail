"""
Regression tests for the `since`/`before` filters on search_emails.

Results carry UTC timestamps (EnvelopeEngine._to_iso renders with
tz=timezone.utc), so a filter value with no offset has to mean UTC too.
It used to mean *local* midnight -- datetime.timestamp() resolves a naive
datetime against the host zone -- so since="2026-08-24T00:00:00" skipped a
message stamped 2026-08-24T03:33Z and the search silently returned nothing.

Run:  uv run --with pytest pytest tests/test_date_filters.py -q
"""

from __future__ import annotations

import os
import time
from datetime import datetime, timezone

import pytest

from apple_mail_mcp.server import _parse_iso_utc


@pytest.fixture(autouse=True)
def ahead_of_utc():
    """Pin the host zone well away from UTC so a local/UTC mixup can't pass."""
    previous = os.environ.get("TZ")
    os.environ["TZ"] = "America/Los_Angeles"
    time.tzset()
    yield
    if previous is None:
        del os.environ["TZ"]
    else:
        os.environ["TZ"] = previous
    time.tzset()


def test_naive_input_is_utc_not_local():
    parsed = _parse_iso_utc("2026-08-24T00:00:00", "since")
    assert parsed == datetime(2026, 8, 24, 0, 0, tzinfo=timezone.utc)
    # The bug: local midnight would land at 07:00Z under PDT.
    assert parsed.timestamp() == datetime(
        2026, 8, 24, 0, 0, tzinfo=timezone.utc
    ).timestamp()


def test_the_reported_case_now_matches():
    """since=2026-08-24T00:00:00 must not exclude a 2026-08-24T03:33Z message."""
    since = _parse_iso_utc("2026-08-24T00:00:00", "since")
    message = datetime(2026, 8, 24, 3, 33, tzinfo=timezone.utc)
    assert message >= since


def test_result_timestamp_round_trips():
    """A date_sent copied out of a search result reproduces the same instant."""
    emitted = "2026-08-24T03:33:00+00:00"
    assert _parse_iso_utc(emitted, "since") == datetime.fromisoformat(emitted)


def test_bare_date_is_utc_midnight():
    assert _parse_iso_utc("2026-08-24", "since") == datetime(
        2026, 8, 24, tzinfo=timezone.utc
    )


@pytest.mark.parametrize(
    "value,expected_hour_utc",
    [
        ("2026-08-24T03:33:00Z", 3),          # Z suffix
        ("2026-08-24T03:33:00+00:00", 3),     # explicit UTC offset
        ("2026-08-23T20:33:00-07:00", 3),     # local wall clock, same instant
        ("2026-08-24T05:33:00+02:00", 3),     # a third zone, same instant
    ],
)
def test_explicit_offsets_are_honoured(value, expected_hour_utc):
    parsed = _parse_iso_utc(value, "since")
    assert parsed.tzinfo is not None
    assert parsed.utcoffset().total_seconds() == 0  # normalised to UTC
    assert parsed.hour == expected_hour_utc


def test_invalid_value_names_the_field_and_shows_examples():
    with pytest.raises(ValueError) as excinfo:
        _parse_iso_utc("last Tuesday", "before")
    message = str(excinfo.value)
    assert "'before'" in message
    assert "last Tuesday" in message
    assert "2026-08-24" in message  # a usable example, not just "ISO-8601"
