# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] — 2026-10-08

### Breaking

- **`set_email_flag` no longer removes a flag by default.** Omitting the flag argument used to clear the flag and report success, so a misspelled parameter name (dropped by the MCP layer) silently unflagged messages. It now raises; pass an explicit `null` to remove a flag. `flag_color` is accepted as an alias for `flag`, and a write that Mail did not actually store raises instead of reporting success.

### Added

- **Quick reply and quick forward, off by default.** `send_email_reply` replies and sends; `send_email_forward` forwards with attachments and an optional note, and sends. Both use Mail's native reply/forward, so threading, quoting and attachments behave as in Mail. They refuse unless the new **Allow sending email** extension setting is on, and carry destructive/open-world tool annotations so hosts can require approval per send. Nothing is sent until the saved message has been read back and checked: the reply text or note must be at the top of the body and a forward's recipients must match exactly; otherwise it is left open in Mail as a draft. A timeout is reported as an unknown outcome with an explicit instruction not to retry. Results report `sent`, `sent_unconfirmed` or `queued_in_outbox`, with a link to the Sent copy when found.
- **`preview_thread`** renders a whole conversation inline as an MCP Apps card: one collapsible row per message, newest open.
- **`get_email_link(scope="thread")`** returns a link resolved at click time that follows the conversation as replies arrive. Mail.app cannot open a conversation, so it fronts the newest message and says so.

### Fixed

- **Flag colours read back correctly.** The colour now comes from bits 39–41 of the Envelope Index's flags field; the `flag_color` column it used before only ever holds 0 or 1, and 17 of 22 flagged messages read back as the wrong colour.
- **Reply drafts whose text contained a double quote failed.** AppleScript strings were escaped by doubling quotes, which is a syntax error; they now use backslash escapes.
- **Flag and attachment calls no longer stall for tens of seconds.** The AppleScript bridge stopped counting every mailbox at start-up (20–40s on a large store), finds messages by id directly (~17ms instead of ~5s on a large inbox), and is told each message's mailbox by the Envelope Index.
- Conversations are titled from their root message, so a reply retitled by a mail gateway does not rename the thread.
- The open-in-Mail redirector's success page closes its own browser tab.

## [1.2.5] — 2026-10-06

### Changed

- **Releases ship only the versioned bundle**, `apple-mail-X.Y.Z.mcpb`. The unversioned `apple-mail.mcpb` is no longer built or published.
- **`build.sh` is shared verbatim across the Apple connectors**, reading names from `manifest.json`. Every repo's build now runs `./test.sh`, validates the manifest, checks the version across `manifest.json`, `pyproject.toml`, `__init__.py`, and `uv.lock`, checks the manifest's tool list against the running server (`.github/check_tools.py`, also used by CI), refuses to overwrite an already-built version without `--force`, and records the bundle's checksum in `dist/SHA1SUMS`. `--skip-tests` skips `./test.sh`.

## [1.2.4] — 2026-10-06

### Changed

- **Permission errors name the exact binary to grant.** Claude Desktop launches extensions through a helper that makes the spawned `uv` — not Claude — responsible for their privacy grants. Errors now print that `uv`'s path (e.g. `~/Library/Application Support/Claude/uv-runtime/uv-0.9.7-darwin-arm64/uv`) ready to paste into System Settings, and explain that enabling Claude is not enough, that one grant covers every Apple connector, and that it must be redone when Claude Desktop updates its `uv`.
- The slow-fallback hints (returned when a search times out or uses a filter the AppleScript path cannot serve) now include those steps instead of pointing at the README.
- **README:** permissions now consistently point at `uv`, the restart step is consistently "quit Claude (⌘Q) and reopen it", and a new *Using several Apple connectors* section covers installing them one at a time, the shared `uv` grant, re-granting after updates, and verifying each connector.
- **Release flow aligned with the other Apple connectors.** Tag-triggered [release workflow](.github/workflows/release.yml) and [CI](.github/workflows/ci.yml), shared verbatim across the family, replace manual releases; this changelog was added and supplies each release's notes.
- CI checks that `manifest.json`, `pyproject.toml`, `__init__.py`, and `uv.lock` agree on the version, and compares the manifest's tool list against the running server rather than a regex over its source.

## Earlier releases

Releases before this one are described on the [GitHub Releases](../../releases) page.
