#!/usr/bin/env bash
# build.sh — validate and pack the Apple Mail MCP desktop extension
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT="${SCRIPT_DIR}/dist"

echo "=== Apple Mail MCP — build ==="

# Check for mcpb CLI
if ! command -v mcpb &>/dev/null; then
    echo "Installing mcpb CLI…"
    npm install -g @anthropic-ai/mcpb
fi

# Validate
echo ""
echo "Validating manifest…"
mcpb validate "${SCRIPT_DIR}/manifest.json"
echo "✓ Manifest valid."

# Pack
echo ""
mkdir -p "${OUT}"
VERSION=$(grep '^version' "${SCRIPT_DIR}/pyproject.toml" | head -1 | sed 's/version = "\(.*\)"/\1/')
STABLE="${OUT}/apple-mail.mcpb"
VERSIONED="${OUT}/apple-mail-${VERSION}.mcpb"
mcpb pack "${SCRIPT_DIR}" "${STABLE}"

# Make the version visible in Finder. We do three things, because for an
# arbitrary `.mcpb` file (not an app bundle) Finder's Get Info pane does
# NOT read kMDItemVersion — it only reads CFBundleShortVersionString from
# Info.plist for bundles. So:
#
#   1. Write a Spotlight comment via Finder, which shows in Get Info →
#      Comments. AppleScript here drives Finder so the comment lands in
#      both Spotlight and the parent folder's .DS_Store.
#   2. Stamp kMDItemVersion as well, for tools / scripts that DO read it
#      (e.g. `mdls`, Spotlight-based file managers).
#   3. Produce a versioned filename copy alongside the stable name. The
#      version is then visible in the filename itself — the bulletproof
#      display that survives copy/upload/quarantine stripping the xattrs.
osascript -e "tell application \"Finder\" to set comment of (POSIX file \"${STABLE}\" as alias) to \"Apple Mail MCP — v${VERSION}\"" >/dev/null 2>&1 || true
xattr -w "com.apple.metadata:kMDItemVersion" "${VERSION}" "${STABLE}" 2>/dev/null || true
mdimport "${STABLE}" 2>/dev/null || true
cp -f "${STABLE}" "${VERSIONED}"
osascript -e "tell application \"Finder\" to set comment of (POSIX file \"${VERSIONED}\" as alias) to \"Apple Mail MCP — v${VERSION}\"" >/dev/null 2>&1 || true
xattr -w "com.apple.metadata:kMDItemVersion" "${VERSION}" "${VERSIONED}" 2>/dev/null || true
mdimport "${VERSIONED}" 2>/dev/null || true

echo ""
echo "✓ Built:"
echo "    ${STABLE}        (stable name — drag-install target)"
echo "    ${VERSIONED}    (versioned name — visible version in filename)"
echo ""
echo "Both files have the version stamped into Finder Comments (Get Info → Comments)."
echo ""
echo "To install: double-click the .mcpb file, or drag it into Claude Desktop."
echo ""
echo "ℹ  Mail.app must be running. macOS will prompt for Automation permission on first use."
