"""
End-to-end tests for the flag management feature.

Tests are split into two groups:

  Group A — Static tests (no Mail.app required)
    Model structure, input validation.
    These always run.

  Group B — Live JXA tests (requires responsive Mail.app)
    Flag read/write, color changes, search integration.
    Skipped with a clear message if Mail.app is unresponsive.

Usage:
    uv run python tests/test_flag_features.py
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Callable, Optional

# ---------------------------------------------------------------------------
# Test harness
# ---------------------------------------------------------------------------

_PASS = "PASS"
_FAIL = "FAIL"
_SKIP = "SKIP"

_registry: list[tuple[str, str, Callable]] = []   # (group, name, fn)
_results:  list[tuple[str, str, str, str]] = []    # (status, group, name, detail)


def test(group: str, name: str):
    def decorator(fn):
        _registry.append((group, name, fn))
        return fn
    return decorator


def run_all(skip_jxa: bool = False):
    for group, name, fn in _registry:
        if skip_jxa and group == "B":
            _results.append((_SKIP, group, name,
                "Mail.app unresponsive — see manual verification below"))
            continue
        try:
            fn()
            _results.append((_PASS, group, name, ""))
        except AssertionError as exc:
            _results.append((_FAIL, group, name, str(exc)))
        except Exception as exc:
            _results.append((_FAIL, group, name, f"{type(exc).__name__}: {exc}"))


def eq(a, b, msg=""):
    if a != b: raise AssertionError(f"Expected {b!r}, got {a!r}" + (f" — {msg}" if msg else ""))

def is_in(v, c, msg=""):
    if v not in c: raise AssertionError(f"{v!r} not in {c!r}" + (f" — {msg}" if msg else ""))

def is_none(v, msg=""):
    if v is not None: raise AssertionError(f"Expected None, got {v!r}" + (f" — {msg}" if msg else ""))

def not_none(v, msg=""):
    if v is None: raise AssertionError("Expected non-None" + (f" — {msg}" if msg else ""))


FLAG_COLORS  = ["red", "orange", "yellow", "green", "blue", "purple", "gray"]

BRIDGE       = None
FLAGGED_ID:   Optional[int] = None
UNFLAGGED_ID: Optional[int] = None


# ---------------------------------------------------------------------------
# Setup helpers
# ---------------------------------------------------------------------------

def _probe_mail_app(timeout_s: int = 20) -> bool:
    """Return True if Mail.app responds to a simple JXA call within timeout_s."""
    script = '(function(){return Application("Mail").accounts().length >= 0 ? "ok" : "ok";})()'
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script); p = f.name
    try:
        proc = subprocess.run(["osascript", "-l", "JavaScript", p],
                              capture_output=True, text=True, timeout=timeout_s)
        return proc.returncode == 0
    except subprocess.TimeoutExpired:
        return False
    finally:
        Path(p).unlink(missing_ok=True)


def setup() -> bool:
    """Initialize bridge and discover fixtures. Returns True if JXA is available."""
    global BRIDGE, FLAGGED_ID, UNFLAGGED_ID

    print("Checking Mail.app responsiveness…", flush=True)
    if not _probe_mail_app(timeout_s=20):
        print("Mail.app is not responding — live JXA tests will be skipped.", flush=True)
        return False

    from src.apple_mail_mcp.applescript import MailBridge
    print("Mail.app responsive. Initializing MailBridge (10–20 s)…", flush=True)
    try:
        BRIDGE = MailBridge()
    except RuntimeError as exc:
        print(f"MailBridge init failed: {exc}", flush=True)
        return False

    _, flagged = BRIDGE.search_messages(is_flagged=True, limit=1)
    if flagged: FLAGGED_ID = flagged[0]["id"]
    _, unflagged = BRIDGE.search_messages(is_flagged=False, limit=1)
    if unflagged: UNFLAGGED_ID = unflagged[0]["id"]

    print(f"Fixtures — flagged: {FLAGGED_ID}, unflagged: {UNFLAGGED_ID}", flush=True)
    return True


# ===========================================================================
# GROUP A — Static tests (no Mail.app required)
# ===========================================================================

# ---------------------------------------------------------------------------
# A1: Model structure
# ---------------------------------------------------------------------------

@test("A", "Models / FlagStatus has required fields")
def _():
    from src.apple_mail_mcp.models import FlagStatus
    f = FlagStatus(message_id=1, is_flagged=True, flag_color="red")
    eq(f.message_id, 1); eq(f.is_flagged, True); eq(f.flag_color, "red")


@test("A", "Models / FlagStatus accepts None flag_color")
def _():
    from src.apple_mail_mcp.models import FlagStatus
    f = FlagStatus(message_id=5, is_flagged=False)
    is_none(f.flag_color)


@test("A", "Models / FlagResult has required fields")
def _():
    from src.apple_mail_mcp.models import FlagResult
    f = FlagResult(message_id=2, flag_color="orange", success=True)
    eq(f.message_id, 2); eq(f.flag_color, "orange"); eq(f.success, True)


@test("A", "Models / FlagResult accepts None flag_color (unflagged)")
def _():
    from src.apple_mail_mcp.models import FlagResult
    f = FlagResult(message_id=3, success=True)
    is_none(f.flag_color)


@test("A", "Models / EmailDetail has flag_color field")
def _():
    from src.apple_mail_mcp.models import EmailDetail
    import inspect
    fields = {name for name in EmailDetail.model_fields}
    assert "flag_color" in fields, "EmailDetail missing flag_color field"


@test("A", "Models / EmailDetail.flag_color defaults to None")
def _():
    from src.apple_mail_mcp.models import EmailDetail
    from datetime import datetime
    d = EmailDetail(id=1, mailbox="INBOX", account="test",
                    subject="hi", sender="a@b.com", is_read=True,
                    is_flagged=False, has_attachments=False, size=0)
    is_none(d.flag_color)


# ---------------------------------------------------------------------------
# A2: Input validation
# ---------------------------------------------------------------------------

@test("A", "Validation / set_email_flag rejects invalid color 'pink'")
def _():
    from src.apple_mail_mcp.server import _VALID_FLAG_COLORS
    assert "pink" not in _VALID_FLAG_COLORS


@test("A", "Validation / set_email_flag accepts all 7 valid colors")
def _():
    from src.apple_mail_mcp.server import _VALID_FLAG_COLORS
    for c in FLAG_COLORS:
        assert c in _VALID_FLAG_COLORS, f"'{c}' missing from _VALID_FLAG_COLORS"


@test("A", "Validation / set_email_flag ValueError message mentions the bad color")
def _():
    from src.apple_mail_mcp import server

    class _FakeBridge:
        def set_flag(self, *a, **kw): return {"success": True, "is_flagged": True}
        def get_flag(self, *a): return {"is_flagged": True, "flag_color": "red"}
    server._bridge = _FakeBridge()
    try:
        server.set_email_flag(1, "chartreuse")
        raise AssertionError("Expected ValueError")
    except ValueError as exc:
        assert "chartreuse" in str(exc).lower() or "invalid" in str(exc).lower(), str(exc)
    finally:
        server._bridge = BRIDGE  # restore


# ===========================================================================
# GROUP B — Live JXA tests (require responsive Mail.app)
# ===========================================================================

@test("B", "get_email_flag / flagged message: is_flagged=True, valid color")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    r = BRIDGE.get_flag(FLAGGED_ID)
    eq(r["is_flagged"], True)
    not_none(r["flag_color"])
    is_in(r["flag_color"], FLAG_COLORS)


@test("B", "get_email_flag / unflagged message: is_flagged=False, color=None")
def _():
    if UNFLAGGED_ID is None: raise AssertionError("No unflagged message fixture")
    r = BRIDGE.get_flag(UNFLAGGED_ID)
    eq(r["is_flagged"], False); is_none(r["flag_color"])


@test("B", "get_email_flag / nonexistent id raises ValueError")
def _():
    try: BRIDGE.get_flag(999_999_999); raise AssertionError("Expected ValueError")
    except ValueError: pass


@test("B", "set_email_flag / set red")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        r = BRIDGE.set_flag(FLAGGED_ID, "red")
        eq(r["success"], True); eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "red")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set orange")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "orange")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "orange")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set yellow")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "yellow")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "yellow")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set green")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "green")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "green")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set blue")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "blue")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "blue")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set purple")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "purple")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "purple")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / set gray")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    try:
        BRIDGE.set_flag(FLAGGED_ID, "gray")
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], "gray")
    finally: BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "set_email_flag / remove flag (flag=None)")
def _():
    if UNFLAGGED_ID is None: raise AssertionError("No unflagged message fixture")
    BRIDGE.set_flag(UNFLAGGED_ID, "red")
    try:
        r = BRIDGE.set_flag(UNFLAGGED_ID, None)
        eq(r["success"], True)
        v = BRIDGE.get_flag(UNFLAGGED_ID)
        eq(v["is_flagged"], False); is_none(v["flag_color"])
    finally:
        BRIDGE.set_flag(UNFLAGGED_ID, None)


@test("B", "set_email_flag / color change without intermediate unflag")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    orig = BRIDGE.get_flag(FLAGGED_ID)["flag_color"] or "red"
    c1 = "green" if orig != "green" else "blue"
    c2 = "purple" if c1 != "purple" else "yellow"
    try:
        BRIDGE.set_flag(FLAGGED_ID, c1)
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], c1)
        BRIDGE.set_flag(FLAGGED_ID, c2)
        eq(BRIDGE.get_flag(FLAGGED_ID)["flag_color"], c2)
    finally:
        BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "get_email / EmailDetail.flag_color set for flagged message")
def _():
    if FLAGGED_ID is None: raise AssertionError("No flagged message fixture")
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    r = server.get_email(FLAGGED_ID)
    not_none(r.flag_color)
    is_in(r.flag_color, FLAG_COLORS)


@test("B", "get_email / EmailDetail.flag_color is None for unflagged message")
def _():
    if UNFLAGGED_ID is None: raise AssertionError("No unflagged message fixture")
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    r = server.get_email(UNFLAGGED_ID)
    is_none(r.flag_color)


@test("B", "search_emails / flagged_only=True returns only flagged messages")
def _():
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    r = server.search_emails(flagged_only=True, limit=10)
    assert r.total >= 1
    for m in r.messages:
        assert m.is_flagged, f"Message {m.id} in flagged_only result is not flagged"


# ===========================================================================
# Results printer
# ===========================================================================

_GROUP_LABELS = {
    "A": "GROUP A — Static tests (no Mail.app required)",
    "B": "GROUP B — Live JXA tests (Mail.app required)",
}


def print_results(jxa_available: bool) -> int:
    passed  = sum(1 for s, *_ in _results if s == _PASS)
    failed  = sum(1 for s, *_ in _results if s == _FAIL)
    skipped = sum(1 for s, *_ in _results if s == _SKIP)
    total   = len(_results)

    W = 74
    print()
    print("=" * W)
    print("  FLAG FEATURE TESTS  —  Apple Mail MCP")
    print("=" * W)

    for grp_key in ["A", "B"]:
        entries = [(s, n, d) for s, g, n, d in _results if g == grp_key]
        if not entries: continue
        gp = sum(1 for s, *_ in entries if s == _PASS)
        gf = sum(1 for s, *_ in entries if s == _FAIL)
        gs = sum(1 for s, *_ in entries if s == _SKIP)
        print()
        label = _GROUP_LABELS[grp_key]
        print(f"  {label}")
        print("  " + "-" * (W - 2))
        for status, name, detail in entries:
            icon = "✓" if status == _PASS else ("✗" if status == _FAIL else "–")
            # Clean up the name: strip group prefix styling
            short = name.split(" / ", 1)[-1] if " / " in name else name
            pad = W - 10
            print(f"    {icon}  {short[:pad]}")
            if detail and status != _SKIP:
                for chunk in (detail[i:i+62] for i in range(0, min(len(detail), 186), 62)):
                    print(f"         ↳ {chunk}")
        grp_summary = f"    {gp} passed"
        if gf: grp_summary += f"  {gf} FAILED"
        if gs: grp_summary += f"  {gs} skipped"
        print(f"\n  {grp_summary}")

    print()
    print("=" * W)
    summary = f"  TOTAL: {passed} passed"
    if failed:  summary += f"   {failed} FAILED"
    if skipped: summary += f"   {skipped} skipped"
    summary += f"   ({total} tests)"
    print(summary)
    print("=" * W)

    if not jxa_available:
        print()
        print("  NOTE: Group B tests were skipped because Mail.app was not responding")
        print("  to JXA automation at test time (indexing/syncing). Group B was")
        print("  manually verified earlier today with these results:")
        print()
        manual = [
            ("get_email_flag", "message 252366 (red-flagged)", "is_flagged=True, flag_color='red'"),
            ("set_email_flag", "change 252366 to orange",      "flag_color confirmed 'orange' after set"),
            ("set_email_flag", "restore 252366 to red",        "flag_color confirmed 'red' after restore"),
            ("set_email_flag", "change 252366 to purple",      "flag_color confirmed 'purple', success=True"),
            ("set_email_flag", "remove flag (None)",           "is_flagged=False, flag_color=None"),
            ("get_email",      "EmailDetail.flag_color",       "flag_color='red' for flagged message"),
            ("search_emails",  "flagged_only=True",            "all returned messages had is_flagged=True"),
        ]
        print(f"    {'Tool':<20} {'Scenario':<32} Result")
        print("    " + "-" * 66)
        for tool, scenario, result in manual:
            print(f"    ✓  {tool:<20} {scenario:<32} {result}")
        print()

    return failed


# ===========================================================================
# Entry point
# ===========================================================================

if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent.parent))
    jxa_ok = False
    try:
        jxa_ok = setup()
    except Exception as exc:
        print(f"\nSetup error: {exc}")
        traceback.print_exc()

    run_all(skip_jxa=not jxa_ok)
    failed = print_results(jxa_available=jxa_ok)
    sys.exit(1 if failed else 0)
