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
mcpb pack "${SCRIPT_DIR}" "${OUT}/apple-mail.mcpb"

# Stamp the version into the file's Spotlight metadata so it shows in
# Finder → Get Info → More Info: Version.
VERSION=$(grep '^version' "${SCRIPT_DIR}/pyproject.toml" | head -1 | sed 's/version = "\(.*\)"/\1/')
xattr -w "com.apple.metadata:kMDItemVersion" "${VERSION}" "${OUT}/apple-mail.mcpb"
mdimport "${OUT}/apple-mail.mcpb" 2>/dev/null || true

echo ""
echo "✓ Built: ${OUT}/apple-mail.mcpb  (v${VERSION})"
echo ""
echo "To install: double-click the .mcpb file, or drag it into Claude Desktop."
echo ""
echo "ℹ  Mail.app must be running. macOS will prompt for Automation permission on first use."
