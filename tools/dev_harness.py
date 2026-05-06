"""Dev harness — run create_reply_draft scenarios against the local Mail.app
without rebuilding the .mcpb. Read-only verification (NEVER deletes).

Usage:
    python tools/dev_harness.py                    # default: all scenarios
    python tools/dev_harness.py --scenarios quoted_simple
    python tools/dev_harness.py --source-id 257008
    python tools/dev_harness.py --inspect 257010   # dump one draft and exit
    python tools/dev_harness.py --probe            # run AppleScript probes only
    python tools/dev_harness.py --watchdog 600     # max wall-clock seconds (default 600)
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import sys
import textwrap
import time
import uuid
from typing import Any, Optional

_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_REPO, "src"))


def _install_watchdog(seconds: int) -> None:
    """Hard SIGALRM after `seconds` so a hung Mail.app can never leave this
    script running indefinitely. Print and exit non-zero. No bare-shell
    `until ! pgrep` waiters anywhere — those self-match and never exit."""
    def _bail(_signum, _frame):
        print(f"\n!! WATCHDOG: forced exit after {seconds}s (Mail.app likely hung).",
              file=sys.stderr, flush=True)
        os._exit(124)
    signal.signal(signal.SIGALRM, _bail)
    signal.alarm(seconds)


from apple_mail_mcp.applescript import MailBridge  # noqa: E402

sys.path.insert(0, _HERE)
from draft_inspect import (  # noqa: E402
    find_drafts_with_subject,
    dump,
)


# ----------------------------------------------------------------------
# Scenario registry
# ----------------------------------------------------------------------

def make_scenarios(unique_token: str) -> dict[str, dict]:
    """Build scenario kwargs. Each body embeds a unique token so we can
    detect it later in the saved draft body."""
    return {
        "quoted_reply": {
            "label": "Quoted, reply (single recipient)",
            "kwargs": {"include_quoted": True, "reply_all": False},
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=quoted_reply\n"
                "This is the test body. The original message should appear quoted below."
            ),
        },
        "quoted_reply_all": {
            "label": "Quoted, reply-all",
            "kwargs": {"include_quoted": True, "reply_all": True},
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=quoted_reply_all\n"
                "Reply-all test. Original quoted below."
            ),
        },
        "no_quote_reply": {
            "label": "No quote, reply",
            "kwargs": {"include_quoted": False, "reply_all": False},
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=no_quote_reply\n"
                "No quoted block expected."
            ),
        },
        "no_quote_reply_all": {
            "label": "No quote, reply-all",
            "kwargs": {"include_quoted": False, "reply_all": True},
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=no_quote_reply_all\n"
                "Reply-all, no quote."
            ),
        },
        "with_cc": {
            "label": "Quoted, reply + extra cc address",
            "kwargs": {
                "include_quoted": True,
                "reply_all": False,
                "cc_addresses": ["Test CC <harness-cc@example.invalid>"],
            },
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=with_cc\n"
                "This reply has an extra cc."
            ),
        },
        "with_bcc": {
            "label": "Quoted, reply + extra bcc address",
            "kwargs": {
                "include_quoted": True,
                "reply_all": False,
                "bcc_addresses": ["Test BCC <harness-bcc@example.invalid>"],
            },
            "body": (
                f"DEV_HARNESS_TOKEN={unique_token} scenario=with_bcc\n"
                "This reply has an extra bcc."
            ),
        },
    }


# ----------------------------------------------------------------------
# Source picking
# ----------------------------------------------------------------------

def pick_source_message(bridge: MailBridge, override: Optional[int]) -> dict:
    if override is not None:
        msg = bridge.get_message(override)
        if not msg:
            raise SystemExit(f"--source-id {override} not found.")
        return msg
    # Pick a recent inbox message that isn't a reply, our test draft,
    # or an artificial probe draft.
    total, msgs = bridge.search_messages(limit=30)
    for m in msgs:
        subj = (m.get("subject") or "").strip()
        if not subj:
            continue
        low = subj.lower()
        if low.startswith("re:") or low.startswith("probe_") or low.startswith("dev_harness"):
            continue
        # Avoid our own outgoing test drafts (sender = us)
        sender = (m.get("sender") or "").lower()
        if "brad" in sender and "tallon" in sender:
            continue
        return m
    if msgs:
        return msgs[0]
    raise SystemExit("No messages found in inbox.")


# ----------------------------------------------------------------------
# Verification
# ----------------------------------------------------------------------

def verify(scenario_name: str, expectations: dict, draft: dict, body_token: str) -> tuple[bool, list[str]]:
    """Compare a draft dict against expectations. Returns (passed, issues)."""
    issues: list[str] = []
    if not draft.get("exists"):
        return False, ["draft does not exist anymore"]

    # 1. Body must contain the unique token (proves our body wasn't lost)
    body = draft.get("body") or ""
    if body_token not in body:
        issues.append(f"body missing token {body_token!r} (got {len(body)} chars; first 80: {body[:80]!r})")

    # 2. Quoted block expectation — look for an attribution line
    expects_quote = expectations["kwargs"]["include_quoted"]
    has_attribution = ("wrote:" in body.lower()) and ("on " in body.lower())
    if expects_quote and not has_attribution:
        issues.append("expected quoted block ('On <date>, <sender> wrote:') — attribution not found")
    if not expects_quote and has_attribution:
        issues.append("expected NO quote, but found attribution line")

    # 3. Threading headers must always be present (proves reply linkage)
    if not draft.get("in_reply_to"):
        issues.append("In-Reply-To header missing — threading broken")
    if not draft.get("references"):
        issues.append("References header missing — threading broken")

    # 4. Reply-all expectation: at least one cc OR multiple to recipients
    if expectations["kwargs"]["reply_all"]:
        n_cc = len(draft.get("cc") or [])
        n_to = len(draft.get("to") or [])
        if n_cc + n_to < 1:
            issues.append("reply_all=True but no cc and no to recipients")

    # 5. Extra cc / bcc additions are present
    extra_cc = expectations["kwargs"].get("cc_addresses") or []
    cc_addrs = " ; ".join(draft.get("cc") or []).lower()
    for cc_addr in extra_cc:
        # Match by email address part (ignore display name capitalization)
        addr_part = cc_addr.split("<")[-1].rstrip(">").strip().lower()
        if addr_part not in cc_addrs:
            issues.append(f"expected cc address {addr_part!r} not found in cc {draft.get('cc')!r}")

    extra_bcc = expectations["kwargs"].get("bcc_addresses") or []
    bcc_addrs = " ; ".join(draft.get("bcc") or []).lower()
    for bcc_addr in extra_bcc:
        addr_part = bcc_addr.split("<")[-1].rstrip(">").strip().lower()
        if addr_part not in bcc_addrs:
            issues.append(f"expected bcc address {addr_part!r} not found in bcc {draft.get('bcc')!r}")

    return (len(issues) == 0), issues


# ----------------------------------------------------------------------
# Pretty printer
# ----------------------------------------------------------------------

def print_draft(d: dict) -> None:
    print(f"  id:          {d['id']}")
    print(f"  exists:      {d.get('exists')}")
    if not d.get("exists"):
        return
    print(f"  account:     {d.get('account')}")
    print(f"  mailbox:     {d.get('mailbox')}")
    print(f"  subject:     {d.get('subject')!r}")
    print(f"  message_id:  {d.get('message_id_header')!r}")
    print(f"  to:          {d.get('to')}")
    print(f"  cc:          {d.get('cc')}")
    print(f"  bcc:         {d.get('bcc')}")
    print(f"  in_reply_to: {d.get('in_reply_to')!r}")
    print(f"  references:  {d.get('references')!r}")
    body = d.get("body") or ""
    print(f"  body length: {len(body)}")
    print(f"  body preview (first 400 chars):")
    preview = body[:400]
    for line in preview.splitlines() or [""]:
        print(f"    │ {line}")
    if len(body) > 400:
        print(f"    │ … (truncated, {len(body) - 400} more chars)")
    headers = d.get("headers") or []
    print(f"  headers ({len(headers)}):")
    for n, v in headers:
        truncv = v if len(v) <= 120 else v[:120] + "…"
        print(f"    {n}: {truncv}")


# ----------------------------------------------------------------------
# Scenario runner
# ----------------------------------------------------------------------

def run_scenario(
    bridge: MailBridge,
    name: str,
    spec: dict,
    source: dict,
    created_ids_log: list[int],
) -> bool:
    print()
    print("=" * 78)
    print(f"Scenario: {name}  —  {spec['label']}")
    print("=" * 78)
    body = spec["body"]
    kwargs = spec["kwargs"]
    expected_subj = "Re: " + (source.get("subject") or "")

    # Pre-state: count drafts matching expected subject
    pre_ids = set(find_drafts_with_subject(bridge, expected_subj))
    print(f"Pre-call drafts matching {expected_subj!r}: {len(pre_ids)}")

    print(f"Calling create_reply_draft(message_id={source['id']}, body=<{len(body)}c>, **{kwargs})")
    t0 = time.monotonic()
    result = bridge.create_reply_draft(source["id"], body, **kwargs)
    elapsed = time.monotonic() - t0
    print(f"  elapsed: {elapsed:.1f}s")
    print(f"  result:  {json.dumps(result, default=str)}")

    # Allow a moment for any sync settling.
    time.sleep(1.0)

    # Post-state
    post_ids = set(find_drafts_with_subject(bridge, expected_subj))
    new_ids = sorted(post_ids - pre_ids)
    created_ids_log.extend(new_ids)
    print(f"Post-call drafts matching subject: {len(post_ids)}  (new this run: {new_ids})")

    if not new_ids:
        print("FAIL: no new draft was created.")
        return False

    # Inspect each new draft
    all_pass = True
    for did in new_ids:
        print(f"\nDraft {did} diagnostics:")
        d = dump(bridge, did)
        print_draft(d)
        body_token_match = body.split("\n", 1)[0]  # the DEV_HARNESS_TOKEN line
        passed, issues = verify(name, spec, d, body_token_match)
        if passed:
            print(f"VERDICT for draft {did}: ✅ PASS")
        else:
            all_pass = False
            print(f"VERDICT for draft {did}: ❌ FAIL")
            for issue in issues:
                print(f"  - {issue}")

    if len(new_ids) != 1:
        all_pass = False
        print(f"VERDICT: ❌ FAIL (expected exactly 1 new draft, got {len(new_ids)})")

    return all_pass


# ----------------------------------------------------------------------
# AppleScript probes (read-only — for ground-truth investigation)
# ----------------------------------------------------------------------

def probe_with_window(bridge: MailBridge, source_id: int) -> None:
    """Probe `reply with opening window`. Read content & Drafts count.
    NEVER closes the window or saves — leaves state intact for inspection."""
    print()
    print("=" * 78)
    print("PROBE WW-1: After `reply ... with opening window`, content & Drafts count")
    print("=" * 78)
    p = bridge._run_applescript(f'''tell application "Mail"
    set srcMsg to missing value
    repeat with acc in (every account)
        if srcMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if srcMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {source_id})
                        if (count of tgts) > 0 then
                            set srcMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if srcMsg is missing value then return "ERROR:source_not_found"
    set expectedSubj to "Re: " & ((subject of srcMsg) as string)

    -- count Drafts BEFORE reply
    set preCount to 0
    repeat with acc in (every account)
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    try
                        if name of mb contains "raft" then
                            set preCount to preCount + (count of (every message of mb whose subject is expectedSubj))
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat

    reply srcMsg with opening window without reply to all
    delay 2

    -- inspect OutgoingMessage
    set foundMsg to missing value
    repeat with m in (get outgoing messages)
        try
            if (subject of m) as string is expectedSubj then
                set foundMsg to m
                exit repeat
            end if
        end try
    end repeat
    set body1 to ""
    if foundMsg is not missing value then
        try
            set body1 to (content of foundMsg) as string
        end try
    end if
    set len1 to (length of body1)
    if len1 > 400 then
        set sample to text 1 thru 400 of body1
    else
        set sample to body1
    end if

    -- count Drafts AFTER reply
    set postCount to 0
    repeat with acc in (every account)
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    try
                        if name of mb contains "raft" then
                            set postCount to postCount + (count of (every message of mb whose subject is expectedSubj))
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat

    return "PRE=" & preCount & linefeed & "POST=" & postCount & linefeed & "OUT_LEN=" & len1 & linefeed & "SAMPLE=" & sample
end tell''', timeout=30)
    print(p or "(no output)")


def probe(bridge: MailBridge, source_id: int) -> None:
    """Run small AppleScript probes to learn what the actual Mail.app
    behavior is. Read-only investigation — does not call create_reply_draft.

    NB: Probes still trigger Mail's `reply` command which auto-saves a draft.
    They never delete it. They observe the auto-save state and then exit.
    """
    print()
    print("=" * 78)
    print("PROBE 1: After `reply ... without opening window`, what does")
    print("         `content of foundMsg` (the OutgoingMessage) return?")
    print("=" * 78)
    p1 = bridge._run_applescript(f'''tell application "Mail"
    -- find the source
    set srcMsg to missing value
    repeat with acc in (every account)
        if srcMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if srcMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {source_id})
                        if (count of tgts) > 0 then
                            set srcMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if srcMsg is missing value then return "ERROR:source_not_found"
    set expectedSubj to "Re: " & ((subject of srcMsg) as string)
    reply srcMsg without opening window without reply to all
    delay 1
    set foundMsg to missing value
    repeat with m in (get outgoing messages)
        try
            if (subject of m) as string is expectedSubj then
                set foundMsg to m
                exit repeat
            end if
        end try
    end repeat
    if foundMsg is missing value then return "ERROR:no_outgoing"
    set body1 to ""
    try
        set body1 to (content of foundMsg) as string
    end try
    set len1 to (length of body1)
    -- prefix only first 300 chars
    if len1 > 300 then
        set sample to text 1 thru 300 of body1
    else
        set sample to body1
    end if
    return "LEN=" & len1 & linefeed & "SAMPLE=" & sample
end tell''', timeout=30)
    print(p1 or "(no output)")

    print()
    print("=" * 78)
    print("PROBE 2: Find the just-created auto-saved Drafts entry; what is")
    print("         its `content`?")
    print("=" * 78)
    # Use a fresh reply from a different (or same) source. We DON'T do another
    # reply here to avoid stacking outgoing messages — instead, find the most
    # recent Drafts entry whose subject matches and read its content.
    p2 = bridge._run_applescript(f'''tell application "Mail"
    set srcMsg to missing value
    repeat with acc in (every account)
        if srcMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if srcMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {source_id})
                        if (count of tgts) > 0 then
                            set srcMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if srcMsg is missing value then return "ERROR:source_not_found"
    set expectedSubj to "Re: " & ((subject of srcMsg) as string)
    set foundDraft to missing value
    set foundId to -1
    repeat with acc in (every account)
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    try
                        if name of mb contains "raft" then
                            repeat with d in (every message of mb whose subject is expectedSubj)
                                try
                                    set dId to (id of d) as integer
                                    if dId > foundId then
                                        set foundDraft to d
                                        set foundId to dId
                                    end if
                                end try
                            end repeat
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if foundDraft is missing value then return "ERROR:no_draft"
    set body2 to ""
    try
        set body2 to (content of foundDraft) as string
    end try
    set len2 to (length of body2)
    if len2 > 300 then
        set sample2 to text 1 thru 300 of body2
    else
        set sample2 to body2
    end if
    return "DRAFT_ID=" & foundId & linefeed & "LEN=" & len2 & linefeed & "SAMPLE=" & sample2
end tell''', timeout=30)
    print(p2 or "(no output)")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--source-id", type=int, default=None)
    p.add_argument("--scenarios",
                   default="quoted_reply,no_quote_reply,quoted_reply_all,no_quote_reply_all,with_cc,with_bcc")
    p.add_argument("--inspect", type=int, default=None,
                   help="Just dump diagnostics for a specific draft id, then exit.")
    p.add_argument("--probe", action="store_true",
                   help="Run read-only AppleScript probes only.")
    p.add_argument("--watchdog", type=int, default=600,
                   help="Max wall-clock seconds before SIGALRM force-exit (default 600).")
    args = p.parse_args()

    _install_watchdog(args.watchdog)

    import logging
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")

    bridge = MailBridge()

    if args.inspect:
        d = dump(bridge, args.inspect)
        print(json.dumps(d, indent=2, default=str))
        return 0

    source = pick_source_message(bridge, args.source_id)
    print(f"Source message:")
    print(f"  id={source.get('id')}  account={source.get('account')}  mailbox={source.get('mailbox')}")
    print(f"  subject={source.get('subject')!r}")
    print(f"  sender={source.get('sender')!r}")

    if args.probe:
        probe(bridge, source["id"])
        probe_with_window(bridge, source["id"])
        return 0

    token = uuid.uuid4().hex[:8]
    scenarios = make_scenarios(token)

    requested = [s.strip() for s in args.scenarios.split(",") if s.strip()]
    unknown = [s for s in requested if s not in scenarios]
    if unknown:
        print(f"Unknown scenario(s): {unknown}. Available: {list(scenarios)}")
        return 2

    created_ids: list[int] = []
    results: dict[str, bool] = {}

    for name in requested:
        passed = run_scenario(bridge, name, scenarios[name], source, created_ids)
        results[name] = passed
        # gentle throttle between scenarios
        time.sleep(1.0)

    print()
    print("=" * 78)
    print("SUMMARY")
    print("=" * 78)
    for name, ok in results.items():
        print(f"  {'✅' if ok else '❌'} {name}")
    print()
    print(f"Drafts created during this run (NOT deleted — clean up via Mail UI):")
    for did in created_ids:
        print(f"  - {did}")

    return 0 if all(results.values()) else 1


if __name__ == "__main__":
    sys.exit(main())
