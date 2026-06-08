#!/bin/bash
# Install / refresh the outlook-mcp auth-server LaunchAgent (+ weekly log cap)
# for the CURRENT user. Idempotent: safe to re-run after editing a plist.
#
# Run from your Mac terminal:
#   bash /Users/Rob/claude/github/outlook-mcp/deploy/install.sh
set -uo pipefail

REPO="/Users/Rob/claude/github/outlook-mcp"
DEPLOY="$REPO/deploy"
LA="$HOME/Library/LaunchAgents"
LOGDIR="$HOME/Library/Logs/outlook-mcp"
UID_NUM="$(id -u)"

SERVER_LABEL="com.robwalsh.outlook-mcp.auth-server"
LOGCAP_LABEL="com.robwalsh.outlook-mcp.logcap"

mkdir -p "$LA" "$LOGDIR"
chmod +x "$DEPLOY/cap-logs.sh"

install_agent () {
    local label="$1"
    cp "$DEPLOY/$label.plist" "$LA/$label.plist"
    launchctl bootout   "gui/$UID_NUM/$label"            2>/dev/null || true
    launchctl bootstrap "gui/$UID_NUM" "$LA/$label.plist" 2>/dev/null || true
    launchctl enable    "gui/$UID_NUM/$label"            2>/dev/null || true
    echo "Installed + loaded: $label"
}

install_agent "$SERVER_LABEL"
install_agent "$LOGCAP_LABEL"

# Start the auth server right now (the log-cap job runs on its weekly schedule)
launchctl kickstart -k "gui/$UID_NUM/$SERVER_LABEL" 2>/dev/null || true

sleep 1
echo
echo "=== auth-server status ==="
launchctl print "gui/$UID_NUM/$SERVER_LABEL" 2>/dev/null | grep -E "state =|pid =|last exit|program =" || echo "(could not read launchctl state)"
echo
echo "=== port 3333 ==="
lsof -iTCP:3333 -sTCP:LISTEN -nP 2>/dev/null || echo "(nothing listening yet — check $LOGDIR/auth-server.error.log)"
echo
echo "Logs:   $LOGDIR/auth-server.log  (errors: auth-server.error.log)"
echo "Done. The server will now start automatically at every login and restart if it crashes."
