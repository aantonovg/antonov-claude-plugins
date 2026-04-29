#!/bin/bash
# LibreOffice MCP plugin — one-shot installer for macOS.
#
# Idempotent. Safe to re-run after `brew upgrade --cask libreoffice`.
# What it does:
#   1. Verifies LibreOffice 24.2+ and `unopkg` are on PATH.
#   2. (macOS Sequoia+) Re-signs LibreOffice ad-hoc to remove launch
#      constraints that block its embedded Python framework.
#   3. Installs the .oxt extension into the user profile via `unopkg add`.
#   4. Tells the user how to verify (open LibreOffice, see HTTP on :8765).
#
# Linux users: re-sign step is skipped. Just runs unopkg add.

set -euo pipefail

PLUGIN_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OXT="$PLUGIN_ROOT/extension/libreoffice-mcp-extension.oxt"

red()    { printf "\033[0;31m%s\033[0m\n" "$*"; }
green()  { printf "\033[0;32m%s\033[0m\n" "$*"; }
yellow() { printf "\033[1;33m%s\033[0m\n" "$*"; }
blue()   { printf "\033[0;34m%s\033[0m\n" "$*"; }

blue "==> LibreOffice MCP plugin installer"
echo  "    plugin root: $PLUGIN_ROOT"
echo  "    .oxt file:   $OXT"
echo

# 1. Prerequisites
if ! command -v soffice >/dev/null 2>&1 && ! command -v libreoffice >/dev/null 2>&1; then
  red   "✗ LibreOffice is not on PATH."
  echo  "  Install it first:"
  echo  "    macOS:        brew install --cask libreoffice"
  echo  "    Debian/Ubuntu: sudo apt install libreoffice"
  echo  "    Fedora:        sudo dnf install libreoffice"
  echo  "    Windows:       https://www.libreoffice.org/download/"
  exit 1
fi
LO_VERSION=$( (soffice --version 2>/dev/null || libreoffice --version 2>/dev/null) | head -1 )
green "✓ LibreOffice found: $LO_VERSION"

if ! command -v unopkg >/dev/null 2>&1; then
  red   "✗ \`unopkg\` is missing — LibreOffice is installed but its tools are not on PATH."
  echo  "  On macOS:   /Applications/LibreOffice.app/Contents/MacOS/unopkg"
  echo  "  Add that directory to PATH, or symlink: ln -s /Applications/LibreOffice.app/Contents/MacOS/unopkg /usr/local/bin/unopkg"
  exit 1
fi
green "✓ unopkg found: $(command -v unopkg)"

if ! command -v uv >/dev/null 2>&1; then
  red   "✗ \`uv\` is missing — needed to run the MCP bridge."
  echo  "  Install: curl -LsSf https://astral.sh/uv/install.sh | sh"
  echo  "       or: brew install uv"
  exit 1
fi
green "✓ uv found: $(uv --version)"

if [[ ! -f "$OXT" ]]; then
  red "✗ Pre-built extension not found at $OXT"
  echo "  Try rebuilding from source: cd extension/source && ./build.sh"
  exit 1
fi
green "✓ pre-built .oxt found"

# 2. macOS-specific re-sign
if [[ "$(uname -s)" == "Darwin" ]]; then
  echo
  yellow "==> macOS detected — checking LibreOffice code signature"
  LO_APP="/Applications/LibreOffice.app"
  if [[ ! -d "$LO_APP" ]]; then
    red "✗ Expected $LO_APP, not found. If LibreOffice is installed in a non-standard location, set LO_APP env var:"
    echo  "  LO_APP=/path/to/LibreOffice.app $0"
    LO_APP="${LO_APP_OVERRIDE:-${LO_APP}}"
    [[ ! -d "$LO_APP" ]] && exit 1
  fi
  # Quick signature probe: if the embedded LibreOfficePython has launch
  # constraints (TDF-issued signature), running it crashes immediately with
  # SIGKILL/CODESIGNING. We check by inspecting codesign flags.
  RESIGN_NEEDED=1
  if codesign -dvvv "$LO_APP" 2>&1 | grep -q 'flags=0x10000(runtime)' \
     && codesign -d --entitlements - "$LO_APP" 2>&1 | grep -q 'com.apple.security.cs.disable-library-validation'; then
    # That's the TDF signature; the only reliable test is a real launch but
    # that's expensive — assume re-sign is needed. ad-hoc re-sign is idempotent.
    RESIGN_NEEDED=1
  fi
  if [[ "$RESIGN_NEEDED" == "1" ]]; then
    yellow "    ad-hoc re-sign needed to allow embedded Python (sudo prompt below)."
    yellow "    This removes launch constraints that crash LibreOffice's Python framework"
    yellow "    on macOS Sequoia. Without it, the MCP extension cannot start its HTTP server."
    echo
    if sudo codesign --force --deep --sign - "$LO_APP"; then
      green "✓ LibreOffice re-signed ad-hoc"
    else
      red "✗ codesign failed. Continuing anyway — extension may still work on macOS Sonoma or earlier."
    fi
  fi
  echo
  yellow "    NOTE: \`brew upgrade --cask libreoffice\` will restore the original signature."
  yellow "    Re-run this installer (or just the codesign command) after every LibreOffice upgrade."
fi

# 3. Install the extension
echo
blue "==> Installing LibreOffice extension"
if pgrep -f "soffice" >/dev/null 2>&1; then
  yellow "    LibreOffice is currently running — closing it so unopkg can register the extension."
  if [[ "$(uname -s)" == "Darwin" ]]; then
    osascript -e 'tell application "LibreOffice" to quit' 2>/dev/null || true
    sleep 3
  fi
  pkill -f soffice 2>/dev/null || true
  sleep 2
fi

# Remove any previous version, then add fresh
unopkg remove org.mcp.libreoffice.extension >/dev/null 2>&1 || true
if unopkg add --suppress-license "$OXT"; then
  green "✓ Extension installed"
else
  red "✗ unopkg add failed"
  exit 1
fi

# 4. Done
echo
green "==> Done."
echo  "    Next:"
echo  "      1. Open LibreOffice (the extension auto-starts an HTTP API at http://localhost:8765)."
echo  "      2. Restart Claude Code so it picks up the libreoffice-live MCP server."
echo  "      3. Verify in Claude Code: \`claude mcp list\` should show 'libreoffice-live: ✓ Connected'."
echo
echo  "    If something hangs or returns 'No Writer document active':"
echo  "      • Make sure a Writer document is open in LibreOffice."
echo  "      • Check the extension log: tail -f /tmp/lo_mcp.log"
echo  "      • Check the extension is enabled: unopkg list | grep mcp"
