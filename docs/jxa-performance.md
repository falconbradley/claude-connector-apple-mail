# Why the JXA bridge was slow, and what it costs now

Measured 2026-09-20 on macOS 27.0 (build 26A428), Mail 16.0, against a live
store: 4 accounts (one disabled), 36 mailboxes, 268,679 messages.

## 1. Construction spent all its time on a count it threw away

`MailBridge.__init__` prescanned every mailbox to learn which ones were
non-empty, then kept only `(account, mailbox)` pairs — the counts were
discarded. Timing each phase of that init script:

| phase | time |
|---|---|
| `System Events` running check | 349 ms |
| `mail.accounts()` | 39 ms |
| enumerating mailboxes | 129 ms |
| **`mb.messages.length` over 36 mailboxes** | **20,227 ms** |
| total | 20,752 ms |

97% of construction was counting messages. Per account: Google 13.6s (20
mailboxes), iCloud 6.4s (8), USC 0.2s (8). Worse, that 20s is the *warm*
number — Mail serves the first query cold, which is how this blew the 40s
timeout in the first place and why a retry moments later succeeded at 30.5s.

No cheaper phrasing of the question exists. Probing only the first element
(`mb.messages[0]`) and bulk-fetching ids were both measured and are the same
speed or slower, because Mail materialises the collection either way:

| strategy | time (warm) |
|---|---|
| `mb.messages.length > 0` | 4,274 ms |
| `mb.messages[0]` probe | 4,737 ms |
| `mb.messages.id().length` | 5,085 ms |

**Fix:** the prescan is now lazy (`MailBridge._nonempty_mailboxes()`),
computed on first use and cached. Construction runs only the running check.
Because it now executes *inside* a tool call rather than before one, it takes
the caller's `_Deadline` so it draws on the same budget instead of adding its
own ceiling on top.

## 2. The prescan is usually not needed at all

It exists to serve two things: the JXA search fallback, and `_find_message`,
which scans non-empty mailboxes to locate a message. But the Envelope Index
already knows every message's mailbox, as a single indexed lookup.

`EnvelopeIndexBridge.get_message_location()` returns it, and
`HybridBridge._jxa(for_message=...)` seeds it into the JXA bridge via
`MailBridge.remember_location()` before any flag read/write. `_find_message`
then returns from cache and neither the prescan nor the scan ever runs.

Validated by sampling one message per mailbox (14) and asking JXA to confirm
each claimed location: **11/14 confirmed**. All three misses are the
*disabled* "Impulse" account, whose messages JXA skips and cannot reach at
all — so the index is reliable for every account JXA can actually serve.

## 3. `byId` instead of `whose({id: …})`

Once the mailbox is known, the message still has to be reached. `whose({id})`
filters the collection; `byId` resolves directly:

| lookup on iCloud/INBOX (46,321 messages) | time |
|---|---|
| `mb.messages.whose({id: N})` | 5,209 ms |
| `mb.messages.byId(N)` | 17 ms |

**~300x faster.** Writes work through it too (`msg.flagIndex = 4` via a `byId`
specifier: 228 ms, confirmed from a separate process).

One sharp edge, verified explicitly: **`byId` is app-global.** It ignores the
mailbox it is rooted at, and `sent.messages.byId(n)` happily returns a message
that lives in INBOX — across accounts, too. It always returns the *correct*
message, so it is safe for reaching a message whose location is already known,
but it can never be used to test whether a mailbox *contains* one. The
membership scan inside `_find_message` therefore still uses `whose({id})`, and
carries a comment saying why. `mail.messages.byId(n)` at application level does
not exist; the specifier must be rooted at a mailbox.

## Results

End-to-end against real Mail, via the MCP tools:

| operation | before | after |
|---|---|---|
| `MailBridge()` construction | 20–40 s | 0.21 s |
| `_find_message` (index-seeded) | 29.6 s | ~0 s |
| JXA `get_flag` | 15.9 s | 0.20 s |
| `set_email_flag` (whole call) | 30–50 s | 0.25 s |
| `search_emails` | — | 1.73 s |

The no-index fallback (`APPLE_MAIL_MCP_DISABLE_FAST=1`, or no Full Disk
Access) still works: `_find_message` pays the prescan and scan, 27.8 s. That
cost is unchanged and unavoidable without the index — but it is now paid only
when something actually needs it, rather than added to every construction.
