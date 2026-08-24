#!/usr/bin/env bash
# extract_icon.sh — refresh the connector's icons from Mail.app.
#
# Dev-only (tools/ is excluded from the .mcpb); the PNGs are committed
# artifacts, so this only needs re-running when Apple restyles the app icon.
# Mirrors the Apple Messages and Apple Reminders connectors.
#
# Small sizes are downscaled from a 1024 master rather than rendered directly,
# so the icon's soft outer shadow is resolved at full resolution first.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
APP="${1:-/System/Applications/Mail.app}"
MASTER="$(mktemp -t icon-master).png"

[[ -d "${APP}" ]] || { echo "No app at ${APP}" >&2; exit 1; }
command -v swift   >/dev/null || { echo "swift not found (install Xcode command line tools)" >&2; exit 1; }
command -v magick  >/dev/null || { echo "ImageMagick (magick) not found" >&2; exit 1; }

swift "${ROOT}/tools/extract_icon.swift" "${APP}" "${MASTER}" 1024

mkdir -p "${ROOT}/icons"
for size in 128 256 512; do
    magick "${MASTER}" -filter Lanczos -resize "${size}x${size}" \
        -strip "${ROOT}/icons/icon-${size}.png"
done
cp -f "${ROOT}/icons/icon-512.png" "${ROOT}/icon.png"
rm -f "${MASTER}"

echo "✓ Extracted icon.png and icons/icon-{128,256,512}.png from ${APP}"
