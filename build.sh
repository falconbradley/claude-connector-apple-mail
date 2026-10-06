#!/usr/bin/env bash
# build.sh — test, validate, and pack this MCP desktop extension into
# dist/<name>-<version>.mcpb. Shared verbatim by every
# claude-connector-apple-* repo; names come from manifest.json and the
# repo's tests from ./test.sh.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "${SCRIPT_DIR}"
OUT="${SCRIPT_DIR}/dist"

FORCE=0
SKIP_TESTS=0
for arg in "$@"; do
    case "${arg}" in
        -f|--force) FORCE=1 ;;
        --skip-tests) SKIP_TESTS=1 ;;
        -h|--help)
            cat <<'USAGE'
usage: build.sh [--force] [--skip-tests]

Runs ./test.sh, validates the manifest, checks the version agrees across
every file that carries it and the manifest's tools match the server, and
packs the extension into dist/<name>-<version>.mcpb.

  -f, --force     Rebuild a version that has already been built, overwriting
                  the existing artifact. Only safe when that artifact never
                  left this machine (see the refusal message for why).
  --skip-tests    Skip ./test.sh (e.g. when they have just run).
USAGE
            exit 0
            ;;
        *)
            echo "✗ Unknown argument: ${arg}" >&2
            echo "  Try: build.sh --help" >&2
            exit 2
            ;;
    esac
done

NAME=$(python3 -c "import json;print(json.load(open('manifest.json'))['name'])")
TITLE="$(python3 -c "import json;print(json.load(open('manifest.json'))['display_name'])") MCP"
VERSION=$(python3 -c "import json;print(json.load(open('manifest.json'))['version'])")
ARTIFACT="${OUT}/${NAME}-${VERSION}.mcpb"
ARTIFACT_NAME="$(basename "${ARTIFACT}")"

echo "=== ${TITLE} ${VERSION} — build ==="

if ! command -v mcpb &>/dev/null; then
    echo "Installing mcpb CLI…"
    npm install -g @anthropic-ai/mcpb
fi

# Test
if [[ "${SKIP_TESTS}" -ne 1 ]]; then
    echo ""
    echo "Running tests…"
    ./test.sh
fi

# Validate
echo ""
echo "Validating manifest…"
mcpb validate manifest.json
echo "✓ Manifest valid."

# The version lives in several files and the packer reads only manifest.json,
# while the server reports __init__.py's over the wire — so any drift ships a
# mislabelled bundle. Fail instead.
python3 .github/check_versions.py

# Every tool the manifest advertises must exist on the server, and vice versa,
# or the installed extension promises a surface it does not have.
uv run --quiet python .github/check_tools.py

# A built version is spent. The bundle carries no build number, so its
# version string is the only identity an artifact has — and `mcpb pack`
# stamps every zip entry with the wall-clock time of the build, so two
# builds of identical source are never byte-identical. A rebuild cannot be
# checked against the original by re-packing and comparing: the second file
# simply overwrites the first and the difference becomes unrecoverable.
# Refuse, and make the bump explicit.
if [[ -e "${ARTIFACT}" && "${FORCE}" -ne 1 ]]; then
    EXISTING_SHA=$(shasum -a 1 "${ARTIFACT}" | cut -d' ' -f1)
    EXISTING_WHEN=$(date -r "${ARTIFACT}" "+%Y-%m-%d %H:%M")
    echo "" >&2
    echo "✗ Version ${VERSION} has already been built." >&2
    echo "    ${ARTIFACT_NAME}  ${EXISTING_SHA}  (${EXISTING_WHEN})" >&2
    echo "" >&2
    echo "  Rebuilding would overwrite it with a different file carrying the" >&2
    echo "  same version — and builds are not reproducible, so the two could" >&2
    echo "  never be told apart afterwards." >&2
    echo "" >&2
    echo "  Bump the version (see the README's Releasing section), then re-run." >&2
    echo "  If that build never left this machine, overwrite it with:" >&2
    echo "    ./build.sh --force" >&2
    exit 1
fi
if [[ -e "${ARTIFACT}" ]]; then
    echo "!  --force: overwriting the existing ${VERSION} artifact."
fi

# Pack
echo ""
mkdir -p "${OUT}"
mcpb pack "${SCRIPT_DIR}" "${ARTIFACT}"

# Make the version visible in Finder beyond the filename. For an arbitrary
# `.mcpb` (not an app bundle) Get Info does NOT read kMDItemVersion, so write
# a Spotlight comment via Finder (shows in Get Info → Comments) and stamp
# kMDItemVersion too, for tools that do read it (`mdls`, Spotlight).
osascript -e "tell application \"Finder\" to set comment of (POSIX file \"${ARTIFACT}\" as alias) to \"${TITLE} — v${VERSION}\"" >/dev/null 2>&1 || true
xattr -w "com.apple.metadata:kMDItemVersion" "${VERSION}" "${ARTIFACT}" 2>/dev/null || true
mdimport "${ARTIFACT}" 2>/dev/null || true

# Record the artifact's checksum. Builds are not reproducible, so a version's
# hash cannot be recovered by rebuilding it later; this ledger is the only
# place it survives.
LEDGER="${OUT}/SHA1SUMS"
if [[ ! -e "${LEDGER}" ]]; then
    {
        echo "# SHA-1 checksums of every ${TITLE} bundle built from this checkout."
        echo "#"
        echo "# Verify the archive:      cd dist && shasum -c SHA1SUMS"
        echo "# Ignoring pruned files:   cd dist && shasum -c --ignore-missing SHA1SUMS"
        echo "#"
        echo "# Appended by build.sh, one line per version, in build order. A --force"
        echo "# rebuild replaces that version's line. Keep comment lines starting with"
        echo "# '#'; a blank line makes shasum report a formatting warning."
    } > "${LEDGER}"
fi
# --force rebuilds a version that is already listed: replace its line rather
# than leaving two entries claiming different hashes for one version.
LEDGER_RE="  ${ARTIFACT_NAME//./\\.}$"        # dots are filename literals
if grep -q "${LEDGER_RE}" "${LEDGER}" 2>/dev/null; then
    grep -v "${LEDGER_RE}" "${LEDGER}" > "${LEDGER}.tmp"
    mv "${LEDGER}.tmp" "${LEDGER}"
fi
( cd "${OUT}" && shasum -a 1 "${ARTIFACT_NAME}" ) >> "${LEDGER}"

echo ""
echo "✓ Built ${ARTIFACT}"
echo "  Checksum recorded in dist/SHA1SUMS — verify with: cd dist && shasum -c SHA1SUMS"
echo ""
echo "To install: double-click the .mcpb file, or drag it into Claude Desktop."
echo "ℹ  macOS attributes this extension's permissions to Claude Desktop's"
echo "   bundled uv, not to Claude — see the README's Permissions section."
