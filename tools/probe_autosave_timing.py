"""Probe: how long does Mail.app's auto-save take after `reply` to materialize
in Drafts? Single AppleScript invocation polls every 0.5s up to 15s, returns
the time-to-appear and the content length."""
from __future__ import annotations

import os
import signal
import sys
import time


def _watchdog(seconds: int = 180) -> None:
    def _bail(_s, _f):
        print(f"!! WATCHDOG: forced exit after {seconds}s.", file=sys.stderr, flush=True)
        os._exit(124)
    signal.signal(signal.SIGALRM, _bail)
    signal.alarm(seconds)


_watchdog(180)

_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_REPO, "src"))
sys.path.insert(0, _HERE)

from apple_mail_mcp.applescript import MailBridge  # noqa: E402
from draft_inspect import find_drafts_with_subject  # noqa: E402


def main() -> int:
    bridge = MailBridge()
    total, msgs = bridge.search_messages(limit=20)
    src = next(
        (m for m in msgs if (m.get("subject") or "").strip() and not (m.get("subject") or "").lower().startswith("re:")),
        msgs[0],
    )
    src_id = src["id"]
    print(f"Source: id={src_id}  subject={src.get('subject')!r}")
    expected_subj = "Re: " + (src.get("subject") or "")
    pre_ids = sorted(find_drafts_with_subject(bridge, expected_subj))
    print(f"Pre-existing drafts matching subject: {len(pre_ids)}")

    # Single AppleScript: reply, then poll Drafts every 0.5s for up to 15s,
    # report when a NEW draft (id not in pre snapshot) first appears AND
    # has non-empty content.
    pre_ids_list = ", ".join(str(i) for i in pre_ids) or "0"
    script = f'''tell application "Mail"
    set existingIds to {{{pre_ids_list}}}
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
    if srcMsg is missing value then return "ERROR:src_not_found"
    set expectedSubj to "Re: " & ((subject of srcMsg) as string)
    set t0 to (current date)

    reply srcMsg without opening window without reply to all

    set log to ""
    repeat 30 times
        set elapsed to ((current date) - t0)
        set foundId to -1
        set foundLen to 0
        repeat with acc in (every account)
            if foundId is not -1 then exit repeat
            try
                if enabled of acc is true then
                    repeat with mb in (every mailbox of acc)
                        if foundId is not -1 then exit repeat
                        try
                            if name of mb contains "raft" then
                                repeat with d in (every message of mb whose subject is expectedSubj)
                                    try
                                        set dId to (id of d) as integer
                                        set isOld to false
                                        repeat with eid in existingIds
                                            if (eid as integer) is dId then
                                                set isOld to true
                                                exit repeat
                                            end if
                                        end repeat
                                        if not isOld then
                                            set foundId to dId
                                            try
                                                set foundLen to length of ((content of d) as string)
                                            end try
                                            exit repeat
                                        end if
                                    end try
                                end repeat
                            end if
                        end try
                    end repeat
                end if
            end try
        end repeat
        set log to log & "t=" & elapsed & "s id=" & foundId & " len=" & foundLen & linefeed
        if foundId > -1 and foundLen > 100 then
            return log & "DONE:t=" & elapsed & "s id=" & foundId & " len=" & foundLen
        end if
        delay 0.5
    end repeat
    return log & "TIMEOUT_AFTER_15s"
end tell'''
    print()
    print("Polling Drafts (every 0.5s for up to 15s) for the auto-save to appear with content >100 chars...")
    out = bridge._run_applescript(script, timeout=30)
    print(out or "(no output)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
