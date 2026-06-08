#!/bin/bash
# Keep the outlook-mcp auth-server logs from growing unbounded.
#
# Uses IN-PLACE truncation (rewrite the file keeping its last N lines) rather than
# renaming, because launchd holds an open O_APPEND file descriptor on these logs.
# Renaming/rotating would leave the server writing to the old inode; rewriting the
# same file keeps launchd's fd valid and simply continues appending afterwards.
set -uo pipefail

LOGDIR="$HOME/Library/Logs/outlook-mcp"
MAX_BYTES=$((5 * 1024 * 1024))   # 5 MB cap per file
KEEP_LINES=2000                  # lines retained when capping

for f in "$LOGDIR/auth-server.log" "$LOGDIR/auth-server.error.log"; do
    [ -f "$f" ] || continue
    size=$(stat -f%z "$f" 2>/dev/null || echo 0)
    if [ "$size" -gt "$MAX_BYTES" ]; then
        tmp="$(mktemp)"
        tail -n "$KEEP_LINES" "$f" > "$tmp"
        cat "$tmp" > "$f"        # truncate-and-rewrite in place; launchd fd stays valid
        rm -f "$tmp"
        echo "$(date '+%Y-%m-%d %H:%M:%S') capped $(basename "$f") (was ${size} bytes, kept last ${KEEP_LINES} lines)"
    fi
done
