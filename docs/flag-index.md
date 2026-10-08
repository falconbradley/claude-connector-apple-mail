# Ground truth: what `flagIndex` and the Envelope Index flag bits actually mean

Probed 2026-09-20 on **macOS 27.0 (build 26A428), Mail 16.0**, against a real
iCloud IMAP message (`279676`, mailbox `INBOX`).

This note exists because `set_email_flag(flag="green")` reported green while
`get_email_flag` read back orange, stably and repeatedly. Two candidate
explanations were ruled out by measurement before the real one was found.

## Method

For each index `0..6`: write `msg.flagIndex = i` from one `osascript -l
JavaScript` process, let that process **exit**, then read the value back from a
**separate** process, and separately read the raw row out of the Envelope Index
SQLite database. The separate-process read matters — the in-script read-back
was itself under suspicion.

## Result 1 — `flagIndex` is exactly what the comment claimed

| index | 0 | 1 | 2 | 3 | 4 | 5 | 6 |
|---|---|---|---|---|---|---|---|
| colour | red | orange | yellow | green | blue | purple | gray |

`flagIndex` round-tripped **perfectly** for all seven values, including from the
unflagged state, read back from a separate process, and stable at +0s, +2s,
+5s and +10s. There is no IMAP/iCloud normalisation, no snapping to a nearby
value, and no commit delay.

**The comment at `applescript.py:67-68` is CONFIRMED, not corrected.** So is the
in-script read-back in `MailBridge.set_flag` — the value it returns is real.
Hypotheses 1 and 2 in the bug report are both disproven.

## Result 2 — the write path was never the problem; the *read* path was

`MailBridge.set_flag` and `MailBridge.get_flag` do share one mapping table, and
an off-by-N there could not produce write-green/read-orange. But the MCP server
does not read through `MailBridge`. `HybridBridge.get_flag` answers from the
**Envelope Index** (SQLite) and only falls back to JXA when the index cannot
supply a colour. The two directions were therefore decoding *different sources*.

`EnvelopeIndexBridge.get_flag` read the colour from `messages.flag_color`.
That column is **not a colour index**:

```
flag_color distribution over all 267,188 messages:  0 -> 202,895   1 -> 32   NULL -> 64,293
over the 84 flagged messages:                       0 -> 52        1 -> 32
```

It only ever holds 0 or 1, and it does not track the colour: writing each of
the seven `flagIndex` values in turn left `flag_color` pinned at `1` every
time. Because `_FLAG_COLOR_ORDER[0] == "red"` and `_FLAG_COLOR_ORDER[1] ==
"orange"`, every flagged message read back as red or orange no matter what
colour it really was. Green (3) reading as orange (1) is exactly that.

## Result 3 — the colour lives in bits 39-41 of `messages.flags`

XOR-ing the `flags` bitfield across the seven writes isolates three bits:

```
bits that vary across colours 0-6: 0x38000000000  ->  bit positions 39, 40, 41
```

| written `flagIndex` | `flags` (hex) | `(flags >> 39) & 0b111` |
|---|---|---|
| unflagged | `0x3020003fc01` | 6 *(stale)* |
| 0 red | `0x00020003fc11` | 0 |
| 1 orange | `0x00820003fc11` | 1 |
| 2 yellow | `0x01020003fc11` | 2 |
| 3 green | `0x01820003fc11` | 3 |
| 4 blue | `0x02020003fc11` | 4 |
| 5 purple | `0x02820003fc11` | 5 |
| 6 gray | `0x03020003fc11` | 6 |

So `(flags >> 39) & 0b111` **is** `flagIndex`, same 0-6 encoding.

Two caveats, both now handled in `EnvelopeIndexBridge.get_flag`:

1. **The colour bits are stale residue once a message is unflagged.** The
   unflagged row above still decodes to 6 ("gray"); only bit 4 (`0x10`,
   flagged) changed. A colour is meaningful *only* when the flagged bit is set.
2. Fall back to JXA when there is no `flags` column, rather than guessing.

## Validation

Every flagged message in the live store was dumped from Mail.app in one JXA
pass (`id`, `flagIndex`) and cross-checked against both decodings:

```
flagged messages cross-checked: 22
  OLD decode (flag_color column):  correct=5   WRONG=17
  NEW decode (flags bits 39-41):   correct=22  WRONG=0
```

17 of 22 flagged messages were being misreported before the fix — including 15
green messages reported as orange and 11 red ones reported as orange.

The three inbox-review carry-forward items (`278773`, `278913`, `279678`) are
genuinely orange and read as orange under **both** decodings, so the carry-forward
behaviour is unchanged by this fix.

## Reproducing

`tests/test_flag_color.py` pins all of the above against a synthetic Envelope
Index. To re-probe a future macOS/Mail version against live mail, write each
index from one process and read it back from another — never trust a read-back
taken inside the writing script without confirming it separately first.
