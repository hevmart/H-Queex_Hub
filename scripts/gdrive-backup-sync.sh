#!/usr/bin/env bash
# Nightly offsite copy of backups/ to Google Drive via rclone.
#
# Uses "copy", not "sync" — deliberately never deletes anything on the Drive
# side. backups/ is locally pruned to BACKUP_RETENTION_DAYS (30 days), but
# the offsite copy is meant to survive local deletion or corruption, so it
# keeps everything ever copied rather than mirroring local retention.
#
# Writes gdrive-sync-status.json next to app.py in the same shape as
# backup-status.json ({"timestamp", "ok", "error"}) so the dashboard can
# show it via _load_gdrive_sync_status() in app.py.
set -u

APP_ROOT="/opt/hqueex-hub"
SOURCE_DIR="$APP_ROOT/backups/"
REMOTE="gdrive-hqueex:H-Queex — Working Documents/H-Queex Hub/Backups/"
STATUS_PATH="$APP_ROOT/gdrive-sync-status.json"

timestamp="$(date +%Y-%m-%dT%H:%M:%S)"

error_output="$(rclone copy "$SOURCE_DIR" "$REMOTE" --create-empty-src-dirs 2>&1)"
exit_code=$?

if [ "$exit_code" -eq 0 ]; then
  cat > "$STATUS_PATH" <<EOF
{
  "timestamp": "$timestamp",
  "ok": true,
  "error": ""
}
EOF
else
  # Escape double quotes and backslashes so the rclone error text is valid inside the JSON string.
  escaped_error="$(printf '%s' "$error_output" | tail -c 2000 | sed 's/\\/\\\\/g; s/"/\\"/g' | tr '\n' ' ')"
  cat > "$STATUS_PATH" <<EOF
{
  "timestamp": "$timestamp",
  "ok": false,
  "error": "rclone copy failed (exit $exit_code): $escaped_error"
}
EOF
fi

exit $exit_code
