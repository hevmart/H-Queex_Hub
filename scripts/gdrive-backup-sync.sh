#!/usr/bin/env bash
# Nightly offsite copy to Google Drive via rclone, two legs:
#   1. backups/ (JSON metadata snapshots) -> Backups/
#   2. OneDrive-stored file binaries (Receipts, SOPs, Delivery Logs,
#      Documents, etc., everything under the Graph "H-Queex Hub Documents"
#      root) -> Documents-Binaries/, synced remote-to-remote (OneDrive ->
#      Google Drive) via rclone rather than round-tripping through this
#      server, since rclone already has working credentials for both sides.
#
# Both legs use "copy", not "sync" — deliberately never deletes anything on
# the Drive side. backups/ is locally pruned to BACKUP_RETENTION_DAYS (30
# days), but the offsite copy is meant to survive local deletion or
# corruption, so it keeps everything ever copied rather than mirroring local
# retention. Same logic applies to the OneDrive binaries leg: a file deleted
# from OneDrive should not vanish from the offsite copy too.
#
# Destination account is the dedicated hqueexbackups@gmail.com Google
# account (see docs/deployment.md "Offsite backup" section) — kept separate
# from any personal Google account so business backup data, including any
# future client data, never sits alongside personal files.
#
# Writes gdrive-sync-status.json next to app.py in the same shape as
# backup-status.json ({"timestamp", "ok", "error"}) so the dashboard can
# show it via _load_gdrive_sync_status() in app.py. "ok" reflects both legs
# together — a failure in either leg marks the whole run as failed, since a
# half-completed offsite backup is not something the dashboard should show
# as green.
set -u

APP_ROOT="/opt/hqueex-hub"
STATUS_PATH="$APP_ROOT/gdrive-sync-status.json"

METADATA_SOURCE="$APP_ROOT/backups/"
METADATA_REMOTE="gdrive-hqueex:H-Queex — Working Documents/H-Queex Hub/Backups/"

BINARIES_SOURCE="onedrive-hqueex:H-Queex Hub Documents"
BINARIES_REMOTE="gdrive-hqueex:H-Queex — Working Documents/H-Queex Hub/Documents-Binaries/"

timestamp="$(date +%Y-%m-%dT%H:%M:%S)"

metadata_error="$(rclone copy "$METADATA_SOURCE" "$METADATA_REMOTE" --create-empty-src-dirs 2>&1)"
metadata_exit=$?

binaries_error="$(rclone copy "$BINARIES_SOURCE" "$BINARIES_REMOTE" --create-empty-src-dirs 2>&1)"
binaries_exit=$?

if [ "$metadata_exit" -eq 0 ] && [ "$binaries_exit" -eq 0 ]; then
  cat > "$STATUS_PATH" <<EOF
{
  "timestamp": "$timestamp",
  "ok": true,
  "error": ""
}
EOF
  exit 0
fi

# Escape double quotes and backslashes so the rclone error text is valid inside the JSON string.
combined_error=""
if [ "$metadata_exit" -ne 0 ]; then
  combined_error="metadata leg failed (exit $metadata_exit): $(printf '%s' "$metadata_error" | tail -c 1000)"
fi
if [ "$binaries_exit" -ne 0 ]; then
  if [ -n "$combined_error" ]; then
    combined_error="$combined_error | "
  fi
  combined_error="${combined_error}binaries leg failed (exit $binaries_exit): $(printf '%s' "$binaries_error" | tail -c 1000)"
fi
escaped_error="$(printf '%s' "$combined_error" | sed 's/\\/\\\\/g; s/"/\\"/g' | tr '\n' ' ')"
cat > "$STATUS_PATH" <<EOF
{
  "timestamp": "$timestamp",
  "ok": false,
  "error": "$escaped_error"
}
EOF

exit 1
