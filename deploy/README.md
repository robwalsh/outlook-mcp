# Auto-start for the Outlook MCP auth server (macOS LaunchAgent)

Keeps `outlook-auth-server.js` (port 3333) running automatically so token re-auth
"just works" for unattended scheduled runs. Installs as a per-user **LaunchAgent**
(starts at login, restarts on crash) — not a system LaunchDaemon, because the auth
flow needs your user session and Chrome.

## What's here

| File | Purpose |
|------|---------|
| `com.robwalsh.outlook-mcp.auth-server.plist` | Runs the auth server; `RunAtLoad` + `KeepAlive`. |
| `com.robwalsh.outlook-mcp.logcap.plist` | Weekly job (Sun 03:10) that caps the log files. |
| `cap-logs.sh` | Truncates each log in place to its last 2000 lines once it passes 5 MB. |
| `install.sh` | Installs/refreshes both agents for the current user. Idempotent. |

Logs go to `~/Library/Logs/outlook-mcp/` (`auth-server.log`, `auth-server.error.log`).

## Install (one command)

```bash
bash /Users/Rob/claude/github/outlook-mcp/deploy/install.sh
```

This copies the plists to `~/Library/LaunchAgents/`, loads them, starts the server,
and prints status + a port-3333 check.

## Verify

```bash
# Is it loaded and running?
launchctl print "gui/$(id -u)/com.robwalsh.outlook-mcp.auth-server" | grep -E "state =|pid ="

# Is it listening?
lsof -iTCP:3333 -sTCP:LISTEN -nP

# Tail the log
tail -f ~/Library/Logs/outlook-mcp/auth-server.log
```

Expected: `state = running`, a pid, and a process listening on 3333.

## Manual install (if you prefer not to run the script)

```bash
mkdir -p ~/Library/LaunchAgents ~/Library/Logs/outlook-mcp
cp deploy/com.robwalsh.outlook-mcp.auth-server.plist ~/Library/LaunchAgents/
cp deploy/com.robwalsh.outlook-mcp.logcap.plist      ~/Library/LaunchAgents/
launchctl bootstrap "gui/$(id -u)" ~/Library/LaunchAgents/com.robwalsh.outlook-mcp.auth-server.plist
launchctl bootstrap "gui/$(id -u)" ~/Library/LaunchAgents/com.robwalsh.outlook-mcp.logcap.plist
launchctl kickstart -k "gui/$(id -u)/com.robwalsh.outlook-mcp.auth-server"
```

## Uninstall

```bash
launchctl bootout "gui/$(id -u)/com.robwalsh.outlook-mcp.auth-server"
launchctl bootout "gui/$(id -u)/com.robwalsh.outlook-mcp.logcap"
rm ~/Library/LaunchAgents/com.robwalsh.outlook-mcp.auth-server.plist
rm ~/Library/LaunchAgents/com.robwalsh.outlook-mcp.logcap.plist
```

## Notes / assumptions baked into the plist

- **Node path:** `/usr/local/bin/node` (from `which node`). If Node moves, update
  `ProgramArguments` in the server plist and re-run `install.sh`.
- **WorkingDirectory** is the repo root so the auth server's `dotenv.config()` finds
  `.env`. Don't remove it.
- Paths are absolute (`/Users/Rob/...`); plists can't expand `~`.
- If port 3333 is already taken (e.g. a manual `npm run auth-server` still running),
  stop that first — the LaunchAgent now owns the port.
