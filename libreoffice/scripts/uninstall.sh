#!/bin/bash
# Remove the LibreOffice MCP extension. Does NOT touch the LibreOffice.app
# code signature (re-sign is idempotent and harmless to leave in place).

set -euo pipefail

if pgrep -f "soffice" >/dev/null 2>&1; then
  echo "Closing LibreOffice…"
  if [[ "$(uname -s)" == "Darwin" ]]; then
    osascript -e 'tell application "LibreOffice" to quit' 2>/dev/null || true
    sleep 3
  fi
  pkill -f soffice 2>/dev/null || true
  sleep 2
fi

if unopkg remove org.mcp.libreoffice.extension; then
  echo "✓ Extension uninstalled."
else
  echo "✗ unopkg remove failed (was the extension installed?)"
  exit 1
fi
