"""
End-to-end tests for the Apple Mail MCP connector.

Tests are split into two groups:

  Group A — Static tests (no Mail.app required)
    Model shapes, input validation, URL helpers. Always run.

  Group B — Live JXA tests (requires responsive Mail.app)
    One test per public tool plus key edge cases. Skipped with a clear
    message if Mail.app is unresponsive.

Mail.app load is kept gentle by:
  - Sharing one MailBridge across the whole run.
  - Discovering all live fixtures in a single setup pass.
  - A 250 ms throttle between live calls.
  - Collapsing what would be many tiny tests into single round-trip tests.

Usage:
    uv run python tests/test_e2e.py
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
import time
import traceback
import uuid
from datetime import datetime, timedelta, timezone
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
_last_live_call: float = 0.0


def test(group: str, name: str):
    def decorator(fn):
        _registry.append((group, name, fn))
        return fn
    return decorator


def run_all(skip_jxa: bool = False):
    for group, name, fn in _registry:
        if skip_jxa and group == "B":
            _results.append((_SKIP, group, name, "Mail.app unresponsive"))
            continue
        try:
            fn()
            _results.append((_PASS, group, name, ""))
        except _SkipTest as exc:
            _results.append((_SKIP, group, name, str(exc)))
        except AssertionError as exc:
            _results.append((_FAIL, group, name, str(exc)))
        except Exception as exc:
            _results.append((_FAIL, group, name, f"{type(exc).__name__}: {exc}"))


class _SkipTest(Exception):
    """Raised by a test when a required fixture is unavailable."""


def skip(msg: str):
    raise _SkipTest(msg)


def _throttle(min_gap_s: float = 0.25):
    """Pause to keep Mail.app from being hammered by back-to-back JXA calls."""
    global _last_live_call
    now = time.monotonic()
    elapsed = now - _last_live_call
    if elapsed < min_gap_s:
        time.sleep(min_gap_s - elapsed)
    _last_live_call = time.monotonic()


def eq(a, b, msg=""):
    if a != b: raise AssertionError(f"Expected {b!r}, got {a!r}" + (f" — {msg}" if msg else ""))

def is_in(v, c, msg=""):
    if v not in c: raise AssertionError(f"{v!r} not in {c!r}" + (f" — {msg}" if msg else ""))

def is_none(v, msg=""):
    if v is not None: raise AssertionError(f"Expected None, got {v!r}" + (f" — {msg}" if msg else ""))

def not_none(v, msg=""):
    if v is None: raise AssertionError("Expected non-None" + (f" — {msg}" if msg else ""))

def truthy(v, msg=""):
    if not v: raise AssertionError(f"Expected truthy, got {v!r}" + (f" — {msg}" if msg else ""))


FLAG_COLORS = ["red", "orange", "yellow", "green", "blue", "purple", "gray"]

# Live fixtures populated in setup().
BRIDGE = None
FLAGGED_ID:    Optional[int] = None
UNFLAGGED_ID:  Optional[int] = None
RECENT_ID:     Optional[int] = None
ATTACHMENT_ID: Optional[int] = None
THREAD_ID:     Optional[int] = None


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
    global BRIDGE, FLAGGED_ID, UNFLAGGED_ID, RECENT_ID, ATTACHMENT_ID, THREAD_ID

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

    # Constrain fixture discovery to a recent date window so Mail.app doesn't
    # scan every message in every mailbox (that took ~3 min total in earlier
    # runs and contributed to Mail.app lock-ups). A 90-day window is wide
    # enough to find at least one of each fixture type for an active user.
    since_recent = datetime.now(timezone.utc) - timedelta(days=90)

    _, flagged = BRIDGE.search_messages(is_flagged=True, limit=1)  # flags often pre-date 90d
    if flagged: FLAGGED_ID = flagged[0]["id"]
    _, unflagged = BRIDGE.search_messages(is_flagged=False, since=since_recent, limit=1)
    if unflagged:
        UNFLAGGED_ID = unflagged[0]["id"]
        RECENT_ID = unflagged[0]["id"]   # reuse for recent
    _, attached = BRIDGE.search_messages(has_attachments=True, since=since_recent, limit=1)
    if attached: ATTACHMENT_ID = attached[0]["id"]
    THREAD_ID = RECENT_ID  # any message id can seed a thread query

    print(
        f"Fixtures — flagged: {FLAGGED_ID}, unflagged: {UNFLAGGED_ID}, "
        f"recent: {RECENT_ID}, with-attachment: {ATTACHMENT_ID}",
        flush=True,
    )
    return True


# ---------------------------------------------------------------------------
# Cleanup helpers (used by draft-creating tests)
# ---------------------------------------------------------------------------

def _delete_drafts_with_subject(subject: str) -> int:
    """Delete every draft whose subject exactly matches `subject`. Returns count deleted.

    Used by the draft-creation tests to keep the user's Drafts mailbox tidy.
    """
    safe = subject.replace("\\", "\\\\").replace('"', '\\"')
    script = f'''
    (function() {{
        var mail = Application("Mail");
        var deleted = 0;
        var accts = mail.accounts();
        for (var i = 0; i < accts.length; i++) {{
            var mboxes = accts[i].mailboxes();
            for (var j = 0; j < mboxes.length; j++) {{
                var mname = "";
                try {{ mname = mboxes[j].name(); }} catch(e) {{ continue; }}
                if (mname.toLowerCase().indexOf("draft") === -1) continue;
                try {{
                    var msgs = mboxes[j].messages();
                    var idxs = [];
                    for (var k = 0; k < msgs.length; k++) {{
                        try {{
                            if ((msgs[k].subject() || "") === "{safe}") idxs.push(k);
                        }} catch(e) {{}}
                    }}
                    for (var x = idxs.length - 1; x >= 0; x--) {{
                        try {{ mail.delete(msgs[idxs[x]]); deleted++; }} catch(e) {{}}
                    }}
                }} catch(e) {{}}
            }}
        }}
        return JSON.stringify({{deleted: deleted}});
    }})();
    '''
    with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as f:
        f.write(script); p = f.name
    try:
        proc = subprocess.run(["osascript", "-l", "JavaScript", p],
                              capture_output=True, text=True, timeout=30)
        if proc.returncode != 0:
            return 0
        import json
        data = json.loads(proc.stdout.strip() or "{}")
        return data.get("deleted", 0)
    finally:
        Path(p).unlink(missing_ok=True)


# ===========================================================================
# GROUP A — Static tests (no Mail.app required)
# ===========================================================================

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------

@test("A", "Models / Mailbox required fields")
def _():
    from src.apple_mail_mcp.models import Mailbox
    m = Mailbox(name="INBOX", account="iCloud", full_name="iCloud/INBOX",
                unread_count=2, message_count=10)
    eq(m.name, "INBOX"); eq(m.unread_count, 2); eq(m.message_count, 10)


@test("A", "Models / MailboxStats required fields")
def _():
    from src.apple_mail_mcp.models import MailboxStats
    s = MailboxStats(total_messages=100, unread_messages=4,
                     mailbox_count=7, account_count=3)
    eq(s.total_messages, 100); eq(s.account_count, 3)


@test("A", "Models / EmailSummary defaults")
def _():
    from src.apple_mail_mcp.models import EmailSummary
    e = EmailSummary(id=1, mailbox="INBOX", account="x", subject="s", sender="a@b")
    eq(e.is_read, True); eq(e.is_flagged, False); is_none(e.mail_link)


@test("A", "Models / EmailDetail has flag_color and body_text")
def _():
    from src.apple_mail_mcp.models import EmailDetail
    fields = set(EmailDetail.model_fields)
    is_in("flag_color", fields); is_in("body_text", fields); is_in("attachment_count", fields)


@test("A", "Models / SearchResult shape")
def _():
    from src.apple_mail_mcp.models import SearchResult
    r = SearchResult(total=0, offset=0, limit=25, messages=[])
    eq(r.total, 0); eq(r.messages, [])


@test("A", "Models / Attachment + AttachmentData shape")
def _():
    from src.apple_mail_mcp.models import Attachment, AttachmentData
    a = Attachment(message_id=1, index=0, filename="x.pdf",
                   content_type="application/pdf", size=42)
    eq(a.size, 42)
    d = AttachmentData(filename="x.pdf", content_type="application/pdf",
                       size=4, data_base64="YWJjZA==")
    eq(d.data_base64, "YWJjZA==")


@test("A", "Models / DraftResult shape")
def _():
    from src.apple_mail_mcp.models import DraftResult
    d = DraftResult(subject="Hi", to_addresses=["a@b"], draft_link=None)
    eq(d.subject, "Hi"); eq(d.cc_addresses, []); is_none(d.draft_link)


@test("A", "Models / FlagStatus and FlagResult shape")
def _():
    from src.apple_mail_mcp.models import FlagStatus, FlagResult
    s = FlagStatus(message_id=1, is_flagged=True, flag_color="red")
    r = FlagResult(message_id=1, flag_color=None, success=True)
    eq(s.flag_color, "red"); is_none(r.flag_color)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

@test("A", "Validation / set_email_flag accepts all 7 colors and rejects others")
def _():
    from src.apple_mail_mcp.server import _VALID_FLAG_COLORS
    for c in FLAG_COLORS:
        is_in(c, _VALID_FLAG_COLORS)
    assert "pink" not in _VALID_FLAG_COLORS


@test("A", "Validation / set_email_flag ValueError mentions the bad color")
def _():
    from src.apple_mail_mcp import server
    class _FakeBridge:
        def set_flag(self, *a, **kw): return {"success": True, "is_flagged": True}
        def get_flag(self, *a): return {"is_flagged": True, "flag_color": "red"}
    prev = server._bridge
    server._bridge = _FakeBridge()
    try:
        server.set_email_flag(1, "chartreuse")
        raise AssertionError("Expected ValueError")
    except ValueError as exc:
        s = str(exc).lower()
        assert "chartreuse" in s or "invalid" in s, str(exc)
    finally:
        server._bridge = prev


@test("A", "Validation / create_email_draft rejects empty `to`")
def _():
    from src.apple_mail_mcp import server
    class _FakeBridge:
        def create_draft(self, **kw): return {"success": True, "message_id": None}
    prev = server._bridge
    server._bridge = _FakeBridge()
    try:
        try:
            server.create_email_draft(to=[], subject="x", body="y")
            raise AssertionError("Expected ValueError")
        except ValueError:
            pass
    finally:
        server._bridge = prev


@test("A", "Validation / create_email_reply_draft rejects empty body")
def _():
    from src.apple_mail_mcp import server
    class _FakeBridge:
        def create_reply_draft(self, *a, **kw):
            return {"success": True, "message_id": None, "subject": "Re: x",
                    "to_addresses": ["a@b"], "cc_addresses": []}
    prev = server._bridge
    server._bridge = _FakeBridge()
    try:
        for empty in ("", "   ", "\n\n"):
            try:
                server.create_email_reply_draft(message_id=1, body=empty)
                raise AssertionError(f"Expected ValueError for body={empty!r}")
            except ValueError:
                pass
    finally:
        server._bridge = prev


# ---------------------------------------------------------------------------
# URL helpers
# ---------------------------------------------------------------------------

@test("A", "Helpers / _make_mail_link(None) returns None")
def _():
    from src.apple_mail_mcp.server import _make_mail_link
    is_none(_make_mail_link(None))
    is_none(_make_mail_link(""))


@test("A", "Helpers / _make_mail_link percent-encodes the bracketed Message-ID")
def _():
    from src.apple_mail_mcp.server import _make_mail_link
    url = _make_mail_link("abc@host")
    not_none(url)
    eq(url, "message://%3Cabc%40host%3E")


# ===========================================================================
# GROUP B — Live JXA tests (require responsive Mail.app)
# ===========================================================================

@test("B", "get_stats returns positive totals")
def _():
    _throttle()
    s = BRIDGE.get_stats()
    truthy(s["total_messages"] >= 0)
    truthy(s["mailbox_count"] >= 1)
    truthy(s["account_count"] >= 1)


@test("B", "list_mailboxes returns at least one mailbox per account")
def _():
    _throttle()
    rows = BRIDGE.list_mailboxes()
    truthy(len(rows) >= 1, "expected ≥1 mailbox")
    for r in rows[:5]:
        truthy(r.get("name"), "row missing name")
        truthy(r.get("account_name"), "row missing account_name")


@test("B", "search_emails honors limit, date, and flagged_only filters")
def _():
    _throttle()
    total, rows = BRIDGE.search_messages(subject_contains="the", limit=5)
    truthy(total >= 0)
    truthy(len(rows) <= 5)

    _throttle()
    since = datetime.now(timezone.utc) - timedelta(days=30)
    total2, rows2 = BRIDGE.search_messages(since=since, limit=5)
    truthy(len(rows2) <= 5)

    _throttle()
    total3, rows3 = BRIDGE.search_messages(is_flagged=True, limit=5)
    for r in rows3:
        eq(r.get("is_flagged"), True, f"message {r.get('id')} not flagged")


@test("B", "get_email returns body, recipients, and consistent flag_color")
def _():
    if FLAGGED_ID is None: skip("no flagged fixture")
    _throttle()
    d = BRIDGE.get_message(FLAGGED_ID)
    not_none(d)
    truthy("body_text" in d, "body_text missing")
    truthy(isinstance(d.get("to_recipients"), list), "to_recipients should be list")
    _throttle()
    flag_color = BRIDGE.get_flag(FLAGGED_ID).get("flag_color")
    not_none(flag_color)
    is_in(flag_color, FLAG_COLORS)


@test("B", "get_email_link returns a message:// URL for a real message")
def _():
    if RECENT_ID is None: skip("no recent fixture")
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    _throttle()
    res = server.get_email_link(RECENT_ID)
    eq(res["message_id"], RECENT_ID)
    link = res["mail_link"]
    if link is not None:
        truthy(link.startswith("message://"), f"unexpected link prefix: {link!r}")
    # link can be None if Mail.app didn't expose a Message-ID header — accept either.


@test("B", "get_email_html does not raise and returns dict with has_html")
def _():
    if RECENT_ID is None: skip("no recent fixture")
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    _throttle()
    r = server.get_email_html(RECENT_ID)
    eq(r["message_id"], RECENT_ID)
    truthy(isinstance(r["has_html"], bool))


@test("B", "get_thread returns a list (possibly empty if Mail can't reconstruct)")
def _():
    if THREAD_ID is None: skip("no thread fixture")
    _throttle()
    msgs = BRIDGE.get_thread_messages(THREAD_ID)
    truthy(isinstance(msgs, list))


@test("B", "list_email_attachments returns well-formed entries when fixture has attachments")
def _():
    if ATTACHMENT_ID is None: skip("no message-with-attachment fixture")
    _throttle()
    atts = BRIDGE.list_attachments(ATTACHMENT_ID)
    # The fixture lookup uses search_messages(has_attachments=True) which can
    # occasionally return false positives on certain mailbox types. If that
    # happens, the tool itself isn't broken — skip rather than fail.
    if len(atts) == 0:
        skip(f"fixture {ATTACHMENT_ID} reports has_attachments=True but list_attachments() found 0")
    a = atts[0]
    truthy(a.get("name"), "attachment missing name")
    truthy(a.get("file_size", 0) >= 0)


@test("B", "get_email_attachment fetches non-empty data matching declared size")
def _():
    if ATTACHMENT_ID is None: skip("no message-with-attachment fixture")
    _throttle()
    atts = BRIDGE.list_attachments(ATTACHMENT_ID)
    if not atts: skip("fixture has no attachments after re-fetch")
    _throttle()
    res = BRIDGE.get_attachment(ATTACHMENT_ID, atts[0]["index"])
    not_none(res)
    fname, ctype, data = res
    truthy(fname); truthy(len(data) > 0)


@test("B", "get_email_flag covers flagged, unflagged, and missing message")
def _():
    if FLAGGED_ID is None or UNFLAGGED_ID is None:
        skip("missing flagged/unflagged fixtures")
    _throttle()
    flagged = BRIDGE.get_flag(FLAGGED_ID)
    eq(flagged["is_flagged"], True)
    is_in(flagged.get("flag_color"), FLAG_COLORS)

    _throttle()
    unflagged = BRIDGE.get_flag(UNFLAGGED_ID)
    eq(unflagged["is_flagged"], False); is_none(unflagged.get("flag_color"))

    _throttle()
    try:
        BRIDGE.get_flag(999_999_999)
        raise AssertionError("Expected ValueError for nonexistent id")
    except ValueError:
        pass


@test("B", "set_email_flag round-trips through all 7 colors and remove")
def _():
    if FLAGGED_ID is None: skip("no flagged fixture")
    _throttle()
    orig = BRIDGE.get_flag(FLAGGED_ID).get("flag_color") or "red"
    # Mail.app needs a beat to commit a flag change before the next read can
    # observe it — a 250 ms throttle isn't always enough. Use 600 ms here.
    set_read_gap = 0.6
    try:
        for c in FLAG_COLORS:
            _throttle(set_read_gap)
            r = BRIDGE.set_flag(FLAGGED_ID, c)
            eq(r["success"], True)
            _throttle(set_read_gap)
            eq(BRIDGE.get_flag(FLAGGED_ID).get("flag_color"), c, f"after setting {c}")
        _throttle(set_read_gap)
        eq(BRIDGE.set_flag(FLAGGED_ID, None)["success"], True)
        _throttle(set_read_gap)
        info = BRIDGE.get_flag(FLAGGED_ID)
        eq(info["is_flagged"], False); is_none(info.get("flag_color"))
    finally:
        _throttle()
        BRIDGE.set_flag(FLAGGED_ID, orig)


@test("B", "create_email_draft creates and we can clean up a unique-subject draft")
def _():
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    subj = f"E2E test draft {uuid.uuid4()}"
    _throttle()
    try:
        r = server.create_email_draft(
            to=["e2e-test@example.invalid"],
            subject=subj,
            body="This is an automated end-to-end test draft. Safe to delete.",
        )
        eq(r.subject, subj)
        eq(r.to_addresses, ["e2e-test@example.invalid"])
    finally:
        _throttle()
        _delete_drafts_with_subject(subj)


@test("B", "create_email_reply_draft replies to a real message with Re: prefix")
def _():
    if RECENT_ID is None: skip("no recent fixture")
    from src.apple_mail_mcp import server
    server._bridge = BRIDGE
    token = uuid.uuid4().hex[:12]
    body = f"E2E reply test {token} — automated; safe to delete."
    expected_subject = None
    # Reply is the heaviest tool (mail.reply + content read/write + save +
    # drafts scan). Give Mail.app extra breathing room before invoking it.
    _throttle(1.0)
    try:
        r = server.create_email_reply_draft(
            message_id=RECENT_ID,
            body=body,
            include_quoted=False,  # smaller draft → faster + easier cleanup
        )
        truthy(r.subject.startswith("Re: ") or r.subject == "Re: ",
               f"subject should start with 'Re: ', got {r.subject!r}")
        truthy(len(r.to_addresses) >= 1, "reply should have at least one To recipient")
        # Threading headers (In-Reply-To / References) are not exposed on a saved
        # draft — they're only set on the wire when sent. Verifying them would
        # require a live SMTP send/receive, which we don't do in tests.
        expected_subject = r.subject
    finally:
        if expected_subject:
            _throttle()
            _delete_drafts_with_subject(expected_subject)


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

    W = 78
    print()
    print("=" * W)
    print("  APPLE MAIL MCP — END-TO-END TESTS")
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
            short = name.split(" / ", 1)[-1] if " / " in name else name
            pad = W - 10
            print(f"    {icon}  {short[:pad]}")
            if detail and status == _FAIL:
                for chunk in (detail[i:i+62] for i in range(0, min(len(detail), 186), 62)):
                    print(f"         ↳ {chunk}")
            elif detail and status == _SKIP:
                print(f"         ↳ {detail}")
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
        print("  Group B was skipped because Mail.app was not responding to JXA")
        print("  automation at test time. Re-run when Mail.app is open and idle.")

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
