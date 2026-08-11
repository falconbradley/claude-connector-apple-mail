"""
Self-test for the fast Envelope Index read path.

Run from a terminal that has Full Disk Access:

    uv run python -m apple_mail_mcp.selftest

What it checks:
  1. Full Disk Access / Envelope Index reachability
  2. Actual schema on this machine (columns the adaptive SQL will use)
  3. Timing of stats / mailbox / search queries
  4. Full message fetch (body via .emlx) for the newest message
  5. ID equivalence between the Envelope Index ROWID and Mail.app's
     scripting id (required for flag/draft writes to target the right
     message). Skipped with --no-jxa or when Mail.app isn't running.

Exit code 0 = all run checks passed.
"""

from __future__ import annotations

import json
import subprocess
import sys
import time

from .envelope import EnvelopeIndexBridge, EnvelopeUnavailable

_OK = "\033[32mPASS\033[0m"
_FAIL = "\033[31mFAIL\033[0m"
_WARN = "\033[33mWARN\033[0m"


def _jxa_probe_first_inbox_message() -> dict | None:
    """Ask Mail.app for one INBOX message's scripting id + RFC Message-ID."""
    script = """
    (function() {
        var mail = Application("Mail");
        var accounts = mail.accounts();
        for (var i = 0; i < accounts.length; i++) {
            if (!accounts[i].enabled()) continue;
            var mboxes = accounts[i].mailboxes.whose({name: "INBOX"});
            for (var j = 0; j < mboxes.length; j++) {
                if (mboxes[j].messages.length === 0) continue;
                var msg = mboxes[j].messages[0];
                return JSON.stringify({
                    id: msg.id(),
                    messageId: msg.messageId(),
                    subject: msg.subject(),
                    account: accounts[i].name()
                });
            }
        }
        return JSON.stringify(null);
    })();
    """
    try:
        proc = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", script],
            capture_output=True,
            text=True,
            timeout=60,
        )
    except subprocess.TimeoutExpired:
        return None
    if proc.returncode != 0 or not proc.stdout.strip():
        print(f"  JXA probe stderr: {proc.stderr.strip()[:200]}")
        return None
    try:
        return json.loads(proc.stdout.strip())
    except json.JSONDecodeError:
        return None


def main() -> int:
    check_jxa = "--no-jxa" not in sys.argv
    failures = 0

    # -- 1. Open the store ------------------------------------------------
    print("== 1. Envelope Index availability ==")
    try:
        t0 = time.monotonic()
        env = EnvelopeIndexBridge()
        print(f"  {_OK}  opened {env.db_path}")
        print(f"        in {(time.monotonic() - t0) * 1000:.0f}ms")
    except EnvelopeUnavailable as exc:
        print(f"  {_FAIL}  {exc}")
        print(
            "\n  To fix: System Settings -> Privacy & Security -> "
            "Full Disk Access -> enable your terminal app (and Claude "
            "Desktop for the extension itself), then rerun."
        )
        return 1

    # -- 2. Schema ---------------------------------------------------------
    print("\n== 2. Schema on this machine ==")
    for table, cols in sorted(env._cols.items()):
        print(f"  {table}: {', '.join(sorted(cols))}")
    print(f"  read expr:      {env._read_expr}")
    print(f"  flagged expr:   {env._flagged_expr}")
    print(f"  deleted filter: {env._not_deleted_expr}")
    print(f"  attachments:    {env._has_attachments_expr}")
    print(f"  epoch:          {'unix' if env._epoch_offset == 0 else 'mac-absolute'}")

    # -- 3. Query timings ---------------------------------------------------
    print("\n== 3. Query timings ==")
    t0 = time.monotonic()
    stats = env.get_stats()
    print(
        f"  get_stats:       {(time.monotonic() - t0) * 1000:7.0f}ms  "
        f"({stats['total_messages']} messages, {stats['unread_messages']} unread, "
        f"{stats['account_count']} accounts)"
    )
    t0 = time.monotonic()
    mailboxes = env.list_mailboxes()
    print(
        f"  list_mailboxes:  {(time.monotonic() - t0) * 1000:7.0f}ms  "
        f"({len(mailboxes)} mailboxes)"
    )
    t0 = time.monotonic()
    total, rows = env.search_messages(limit=25)
    search_ms = (time.monotonic() - t0) * 1000
    print(f"  search (recent): {search_ms:7.0f}ms  ({total} total, {len(rows)} returned)")
    t0 = time.monotonic()
    total_u, _ = env.search_messages(is_unread=True, limit=25)
    print(
        f"  search (unread): {(time.monotonic() - t0) * 1000:7.0f}ms  ({total_u} total)"
    )
    t0 = time.monotonic()
    total_q, _ = env.search_messages(
        subject_contains="invoice", sender_contains="invoice", limit=25
    )
    print(
        f"  search (query):  {(time.monotonic() - t0) * 1000:7.0f}ms  ({total_q} matching 'invoice')"
    )
    if not rows:
        print(f"  {_FAIL}  search returned no messages at all")
        return 1
    status = _OK if search_ms < 2000 else _WARN
    print(f"  {status}  queries executed")

    # -- 4. Full message fetch ----------------------------------------------
    print("\n== 4. Newest message via fast path ==")
    newest = rows[0]
    print(f"  id={newest['id']}  [{newest['account_name']}/{newest['mailbox_name']}]")
    print(f"  subject: {newest['subject'][:70]!r}")
    t0 = time.monotonic()
    detail = env.get_message(newest["id"])
    fetch_ms = (time.monotonic() - t0) * 1000
    if detail is None:
        print(f"  {_FAIL}  get_message returned None")
        failures += 1
    else:
        body = (detail.get("body_text") or "").strip()
        rfc_id = detail.get("message_id")
        print(f"  get_message: {fetch_ms:.0f}ms  body={len(body)} chars  "
              f"to={len(detail['to_recipients'])} cc={len(detail['cc_recipients'])}")
        print(f"  Message-ID header: {rfc_id!r}")
        if body and rfc_id:
            print(f"  {_OK}  body + headers extracted from .emlx")
        elif not body:
            # The first .emlx index build happens inside get_message; a
            # missing body usually means the emlx file wasn't found.
            print(f"  {_WARN}  no body text — check .emlx mapping "
                  "(may be an HTML-only or not-yet-downloaded message; "
                  "try rerunning)")

    # -- 5. Scripting-id equivalence -----------------------------------------
    print("\n== 5. ROWID vs Mail.app scripting id ==")
    if not check_jxa:
        print("  skipped (--no-jxa)")
    else:
        probe = _jxa_probe_first_inbox_message()
        if probe is None:
            print(f"  {_WARN}  could not probe Mail.app via JXA "
                  "(is Mail.app running? does this terminal have Automation "
                  "permission?) — writes can't be verified, reads unaffected")
        else:
            jxa_id = probe["id"]
            jxa_rfc = (probe.get("messageId") or "").strip().strip("<>")
            fast_rfc = env.get_message_id_header(jxa_id)
            print(f"  JXA message: id={jxa_id} account={probe.get('account')!r}")
            print(f"  JXA  Message-ID: {jxa_rfc!r}")
            print(f"  fast Message-ID: {fast_rfc!r}")
            if fast_rfc and jxa_rfc and fast_rfc == jxa_rfc:
                print(f"  {_OK}  ids are equivalent — writes will target "
                      "the right messages")
            elif fast_rfc is None:
                print(f"  {_FAIL}  ROWID {jxa_id} has no .emlx / not in index — "
                      "id equivalence NOT confirmed; do not trust set_email_flag "
                      "on fast-path ids until this passes")
                failures += 1
            else:
                print(f"  {_FAIL}  Message-ID mismatch — ids are NOT equivalent; "
                      "writes on fast-path ids would hit the wrong message")
                failures += 1

    print(f"\n{'All checks passed.' if failures == 0 else f'{failures} check(s) failed.'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
