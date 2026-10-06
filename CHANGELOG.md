# Changelog

All notable changes to this project are documented here.
This project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
