"""Test 3 candidate strategies for create_reply_draft, then report which
saves a draft with our exact body content.

Read-only on existing drafts. Creates 3 test drafts (one per strategy).
Never deletes anything.
"""
from __future__ import annotations

import os
import signal
import sys
import time


def _watchdog(seconds: int = 300) -> None:
    def _bail(_s, _f):
        print(f"!! WATCHDOG: forced exit after {seconds}s.", file=sys.stderr, flush=True)
        os._exit(124)
    signal.signal(signal.SIGALRM, _bail)
    signal.alarm(seconds)


_watchdog(300)

_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_REPO, "src"))
sys.path.insert(0, _HERE)

from apple_mail_mcp.applescript import MailBridge  # noqa: E402
from draft_inspect import find_drafts_with_subject, dump  # noqa: E402


def pick_source(bridge: MailBridge) -> dict:
    total, msgs = bridge.search_messages(limit=20)
    for m in msgs:
        s = (m.get("subject") or "").strip()
        if s and not s.lower().startswith("re:"):
            return m
    return msgs[0]


def run_strategy(bridge: MailBridge, label: str, script: str) -> dict:
    print()
    print("=" * 78)
    print(f"STRATEGY: {label}")
    print("=" * 78)
    t0 = time.monotonic()
    out = bridge._run_applescript(script, timeout=45)
    elapsed = time.monotonic() - t0
    print(f"  elapsed: {elapsed:.1f}s")
    print(f"  AppleScript output:")
    for line in (out or "(none)").splitlines():
        print(f"    | {line}")
    return {"output": out, "elapsed": elapsed}


def main() -> int:
    bridge = MailBridge()
    src = pick_source(bridge)
    print(f"Source: id={src['id']}  subject={src.get('subject')!r}")
    src_id = src["id"]
    expected_subj = "Re: " + (src.get("subject") or "")
    pre_ids = set(find_drafts_with_subject(bridge, expected_subj))
    print(f"Pre-existing drafts matching expected subject: {len(pre_ids)}")

    # Strategy A: reply WITHOUT window, set content, save
    marker_a = f"PROBE_A_{int(time.time())}"
    script_a = f'''tell application "Mail"
    set srcMsg to missing value
    repeat with acc in (every account)
        if srcMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if srcMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {src_id})
                        if (count of tgts) > 0 then
                            set srcMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if srcMsg is missing value then return "ERROR:src"

    -- Snapshot OutgoingMessage count
    set preCount to count of (get outgoing messages)

    reply srcMsg without opening window without reply to all
    delay 1

    set postList to get outgoing messages
    set postCount to count of postList
    if postCount is preCount then return "ERROR:no_new_outgoing"
    set foundMsg to last item of postList

    set content of foundMsg to "{marker_a} body"
    save foundMsg

    return "SUCCESS:new_outgoing_count_delta=" & (postCount - preCount)
end tell'''
    res_a = run_strategy(bridge, "A: reply WITHOUT window + set content + save", script_a)
    time.sleep(2.0)

    # Strategy B: reply WITH window + set visible false + set content + save
    marker_b = f"PROBE_B_{int(time.time())}"
    script_b = f'''tell application "Mail"
    set srcMsg to missing value
    repeat with acc in (every account)
        if srcMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if srcMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {src_id})
                        if (count of tgts) > 0 then
                            set srcMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if srcMsg is missing value then return "ERROR:src"

    set preCount to count of (get outgoing messages)
    reply srcMsg with opening window without reply to all
    delay 0.4

    set postList to get outgoing messages
    if (count of postList) is preCount then return "ERROR:no_new_outgoing"
    set foundMsg to last item of postList

    -- Hide ASAP
    try
        set visible of foundMsg to false
    end try

    -- Capture Mail's auto-generated quoted block
    set qb to ""
    try
        set qb to (content of foundMsg) as string
    end try
    set qbLen to length of qb

    set content of foundMsg to "{marker_b} body"
    save foundMsg

    return "SUCCESS:auto_quote_len=" & qbLen
end tell'''
    res_b = run_strategy(bridge, "B: reply WITH window + visible=false + set + save", script_b)
    time.sleep(2.0)

    # Strategy C: make new outgoing message with properties (no reply — no threading)
    marker_c = f"PROBE_C_{int(time.time())}"
    script_c = f'''tell application "Mail"
    set newMsg to make new outgoing message with properties {{visible:false, subject:"PROBE_C_test", content:"{marker_c} body"}}
    save newMsg
    return "SUCCESS:made_outgoing_message"
end tell'''
    res_c = run_strategy(bridge, "C: make new outgoing message + save (NO threading)", script_c)
    time.sleep(2.0)

    # Now identify the new drafts and inspect them
    print()
    print("=" * 78)
    print("INSPECTING NEW DRAFTS")
    print("=" * 78)
    post_ids = set(find_drafts_with_subject(bridge, expected_subj))
    new_for_subject = sorted(post_ids - pre_ids)
    # Also find the strategy-C draft (different subject)
    c_ids = find_drafts_with_subject(bridge, "PROBE_C_test")

    findings = {}
    for did in new_for_subject:
        d = dump(bridge, did)
        body = d.get("body") or ""
        contains_a = marker_a in body
        contains_b = marker_b in body
        findings[did] = {
            "exists": d.get("exists"),
            "body_len": len(body),
            "body_first_300": body[:300],
            "contains_marker_a": contains_a,
            "contains_marker_b": contains_b,
            "in_reply_to": d.get("in_reply_to"),
            "has_quote_attribution": ("wrote:" in body.lower() and "on " in body.lower()),
        }
        print(f"\nDraft id={did}  subject={d.get('subject')!r}")
        print(f"  body_len:               {findings[did]['body_len']}")
        print(f"  contains marker A?      {contains_a}")
        print(f"  contains marker B?      {contains_b}")
        print(f"  in_reply_to set?        {bool(findings[did]['in_reply_to'])}")
        print(f"  has quote attribution?  {findings[did]['has_quote_attribution']}")
        print(f"  body first 300:")
        for line in body[:300].splitlines() or [""]:
            print(f"    │ {line}")

    for did in c_ids:
        d = dump(bridge, did)
        body = d.get("body") or ""
        contains_c = marker_c in body
        print(f"\nDraft id={did}  subject={d.get('subject')!r}")
        print(f"  body_len:               {len(body)}")
        print(f"  contains marker C?      {contains_c}")
        print(f"  body first 200:")
        for line in body[:200].splitlines() or [""]:
            print(f"    │ {line}")

    print()
    print("=" * 78)
    print("VERDICT")
    print("=" * 78)
    a_works = any(f["contains_marker_a"] for f in findings.values())
    b_works = any(f["contains_marker_b"] for f in findings.values())
    print(f"  Strategy A (without window):           {'✅ WORKS' if a_works else '❌ FAILS'}")
    print(f"  Strategy B (with window, hidden):      {'✅ WORKS' if b_works else '❌ FAILS'}")
    print(f"  Strategy C (make new outgoing):        ran (check threading manually)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
