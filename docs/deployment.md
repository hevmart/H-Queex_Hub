# H-Queex Hub — Deployment (Hetzner VPS)

This documents the first real deployment of H-Queex Hub, done ahead of launch for
testing purposes. Local dev (Flask dev server via `Launch-HQueex-Hub.ps1`) is
unaffected and remains the day-to-day workflow — this is a separate, always-on
instance for testing the app the way it'll actually run in production.

## Current deployment

- **Server**: Hetzner VPS, Ubuntu 24.04.4 LTS, IP `167.233.110.170`
- **Domain**: `hub.h-queex.com` (A record → server IP)
- **App path**: `/opt/hqueex-hub` (git clone of `origin/main`)
- **Process manager**: systemd unit `hqueex-hub.service` running gunicorn
- **Reverse proxy / TLS**: Caddy (`hqueex-hub` proxied on `127.0.0.1:8000`),
  automatic HTTPS via Let's Encrypt
- **Deploy user**: `deploy` (passwordless sudo, SSH key auth only)

## 1. Server hardening

Root SSH login and password auth are both disabled. All administration goes
through the `deploy` user (key-based auth, passwordless sudo).

```bash
# as root, one-time setup
useradd -m -s /bin/bash deploy
usermod -aG sudo deploy
mkdir -p /home/deploy/.ssh
cp /root/.ssh/authorized_keys /home/deploy/.ssh/authorized_keys
chown -R deploy:deploy /home/deploy/.ssh
chmod 700 /home/deploy/.ssh && chmod 600 /home/deploy/.ssh/authorized_keys
echo "deploy ALL=(ALL) NOPASSWD:ALL" > /etc/sudoers.d/deploy
chmod 440 /etc/sudoers.d/deploy

# then, after confirming `ssh deploy@<ip>` + `sudo whoami` both work:
sed -i 's/^#\?PermitRootLogin.*/PermitRootLogin no/' /etc/ssh/sshd_config
sed -i 's/^#\?PasswordAuthentication.*/PasswordAuthentication no/' /etc/ssh/sshd_config
sshd -t && systemctl reload ssh
```

Firewall (ufw) allows only SSH, HTTP, HTTPS:

```bash
ufw allow OpenSSH
ufw allow 80/tcp
ufw allow 443/tcp
ufw --force enable
```

**Never disable root login before verifying the new user can actually log in
and sudo** — do it in the same session, immediately after the check passes,
so there's no window where the server is unreachable.

## 2. System dependencies

```bash
apt-get update
apt-get install -y \
  python3 python3-venv python3-pip python3-dev \
  git build-essential \
  python3-cffi python3-brotli libpango-1.0-0 libpangoft2-1.0-0 \
  libpangocairo-1.0-0 libgdk-pixbuf2.0-0 libffi-dev shared-mime-info \
  fonts-liberation \
  poppler-utils
```

- The `libpango*`/`libgdk-pixbuf*`/`libffi-dev` set is WeasyPrint's documented
  Linux dependency list (Pango ≥ 1.44 required; Ubuntu 24.04 ships 1.52, well
  above that — no PPA needed, unlike older Ubuntu releases).
- `poppler-utils` is for `pdf2image` (receipt OCR's PDF-to-image step) —
  the equivalent of the vendored Poppler binary used on the local Windows dev
  machine, but installed as a normal system package on Linux.
- Verified with `python -c "import weasyprint"` after the venv was set up —
  confirms the native libs are actually linkable, not just that apt reported
  success.

## 3. App checkout and Python environment

```bash
mkdir -p /opt/hqueex-hub && chown deploy:deploy /opt/hqueex-hub
git clone https://github.com/hevmart/H-Queex_Hub.git /opt/hqueex-hub
cd /opt/hqueex-hub
python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt
.venv/bin/pip install gunicorn
```

`gunicorn` is not in `requirements.txt` (it's a deploy-time concern, not an
app dependency) — install it into the venv separately, as above.

**To redeploy after a `git push` to `main`:**

```bash
ssh deploy@167.233.110.170
cd /opt/hqueex-hub
git pull
.venv/bin/pip install -r requirements.txt   # only if requirements changed
sudo systemctl restart hqueex-hub
```

## 4. Environment variables

Secrets and per-deployment config live outside the repo, in
`/etc/hqueex-hub/hqueex-hub.env` (root:deploy, mode 640 — not world-readable),
loaded by systemd via `EnvironmentFile=`.

```
HQ_SECRET_KEY=<64-char hex, generated fresh on the server — never reused from local .env>
HQ_ALLOWED_API_ORIGINS=https://h-queex.com,https://www.h-queex.com,https://h-queex.ie,https://www.h-queex.ie,https://h-queex.netlify.app
HQ_NETLIFY_SITE_SLUG=h-queex
```

`https://h-queex.netlify.app` (the site's default, non-custom Netlify subdomain) is listed
explicitly because `HQ_NETLIFY_SITE_SLUG`'s regex only matches deploy-preview subdomains
(`<slug>--h-queex.netlify.app`), not the bare production one — added so the lead-intake
forms work when testing against `h-queex.netlify.app` directly, before/without
`h-queex.com` DNS pointing at Netlify.

Generate a fresh secret key the same way for any future server:

```bash
python3 -c "import secrets; print(secrets.token_hex(32))"
```

**Not currently set** (add to the same file if/when needed):
- `ANTHROPIC_API_KEY` — required for the receipt OCR feature to work on this
  server. Not set on this test instance yet; OCR upload will fail without it,
  everything else works fine.

### Microsoft Graph (Company Documents → OneDrive)

Company Documents (`graph_documents.py`) stores files in OneDrive via delegated
OAuth as `hmartire@h-queex.com`, not local disk — see the module docstring and
`scripts/graph_authorize.py` for how the refresh token is originally obtained
(a one-time interactive consent run *locally*, never on the server — it needs
a real browser).

```
GRAPH_TENANT_ID=8509ed16-1e5d-4472-a776-56d4baa6d2b3
GRAPH_CLIENT_ID=6d96a8f2-3085-4496-8f6c-74f551c706e1
GRAPH_CLIENT_SECRET=<copied from local .env, never generated fresh per-server — this is the same Azure app registration>
GRAPH_REFRESH_TOKEN=<copied from local .env after running scripts/graph_authorize.py locally>
GRAPH_ENV_FILE=/etc/hqueex-hub/hqueex-hub.env
LEAD_NOTIFICATION_EMAIL=hmartire@h-queex.com   # optional; this is the default if unset
```

**Lead notification email**: `/api/leads` sends a best-effort email (via
Graph `Mail.Send`, as hmartire@h-queex.com) to `LEAD_NOTIFICATION_EMAIL`
whenever a website lead is created — a mail failure never fails the API
response since the lead is already saved by that point. This requires the
`Mail.Send` delegated permission on the same Azure app registration as the
OneDrive integration (added in Azure Portal → App registrations → API
permissions), and a refresh token obtained *after* that permission was
added — re-run `scripts/graph_authorize.py` locally any time `GRAPH_SCOPE`
in `graph_documents.py` changes, since a refresh token only carries the
scopes that were consented to at the time it was issued.

**`GRAPH_ENV_FILE` is required on the server and easy to miss.** Locally,
`graph_documents.py` persists a rotated refresh token back into the project's
own `.env` (found automatically, no config needed). The server has no such
`.env` — credentials arrive via systemd's `EnvironmentFile=` — so without
`GRAPH_ENV_FILE` pointing at that same file, a rotated token would silently
get written to a stray `/opt/hqueex-hub/.env` that nothing ever reads back,
and the real env file would keep aging until its token's 90-day window lapsed
(see `graph_documents._env_path()`).

**File permissions need one adjustment from the 640 default above**: the
`deploy` user (which owns the `hqueex-hub` systemd service) must be able to
*write* to this file, not just read it, so the rotated-token persist can
actually land. Set it to `660` (root:deploy, group read+write) rather than
the `640` used for the rest of this file's contents:

```bash
sudo chmod 660 /etc/hqueex-hub/hqueex-hub.env
```

## 5. Running as a service (gunicorn + systemd)

Unit file: `/etc/systemd/system/hqueex-hub.service`

```ini
[Unit]
Description=H-Queex Hub (gunicorn)
After=network.target

[Service]
Type=notify
User=deploy
Group=deploy
WorkingDirectory=/opt/hqueex-hub
EnvironmentFile=/etc/hqueex-hub/hqueex-hub.env
ExecStart=/opt/hqueex-hub/.venv/bin/gunicorn --workers 3 --bind 127.0.0.1:8000 --timeout 120 app:app
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable hqueex-hub
sudo systemctl start hqueex-hub
```

Gunicorn binds to `127.0.0.1:8000` only — never exposed directly to the
internet, always reached through Caddy. `app:app` refers to the module-level
`app = Flask(__name__)` object in `app.py` — the same object the Flask dev
server runs, just served by gunicorn's worker model instead of
`app.run(debug=False, use_reloader=False)`.

**Common commands:**

```bash
sudo systemctl status hqueex-hub      # is it running
sudo journalctl -u hqueex-hub -f      # tail logs
sudo systemctl restart hqueex-hub     # after a git pull / env change
```

## 6. Caddy (automatic HTTPS)

```bash
apt-get install -y debian-keyring debian-archive-keyring apt-transport-https curl
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' \
  | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' \
  | tee /etc/apt/sources.list.d/caddy-stable.list
apt-get update && apt-get install -y caddy
```

`/etc/caddy/Caddyfile`:

```
hub.h-queex.com {
	reverse_proxy 127.0.0.1:8000
}
```

```bash
sudo caddy validate --config /etc/caddy/Caddyfile
sudo systemctl restart caddy
sudo systemctl enable caddy
```

Caddy obtains and renews the Let's Encrypt certificate for `hub.h-queex.com`
automatically — no certbot, no manual renewal cron. It requires DNS to
already resolve to the server before it can complete the ACME HTTP-01
challenge; if the Caddyfile is deployed before the DNS record propagates,
Caddy just keeps retrying in the background (visible via
`journalctl -u caddy`) and picks up the cert as soon as DNS resolves — no
redeploy needed.

## 7. First-run setup on a fresh instance

Same behavior as local: with no `users.json` present, any request redirects
to `/setup`, which lets you create the first Owner account (bcrypt-hashed
password, written to `users.json` in the app directory). After that,
`/setup` stops being reachable and normal `/login` takes over.

```bash
curl -s https://hub.h-queex.com/api/health
# {"data_files_readable":true,"service":"H-Queex Hub","status":"ok",...}
```

Visit `https://hub.h-queex.com/setup` in a browser to create the Owner
account. `users.json` (like all the other JSON data files) lives directly in
`/opt/hqueex-hub` alongside the code — it is **not** committed to git (see
`.gitignore`), same as local.

## 8. Offsite backup (rclone → Google Drive)

`_backup_json_file()` in `app.py` writes every JSON save to a local
`backups/<date>/` folder immediately (works identically on every machine),
and also *attempts* a live copy to a Google Drive mount at `G:/` — which is
only ever real on the local Windows dev machine (Google Drive Desktop). On
a headless Linux server there is no `G:` drive; `_backup_json_file` checks
`GDRIVE_MOUNT_ROOT.is_dir()` before attempting anything and correctly
reports `gdrive_ok: false` there rather than fabricating a fake local `G:`
folder (an actual bug this caught and fixed — see git history). That per-save
live-mount status still shows in Settings → Backups for informational
purposes, but a headless server will always show it as `✗`, by design — it
is **not** a failure there, and does not trigger the dashboard warning
banner.

Real offsite protection for the server instead comes from a separate,
independent nightly job: `scripts/gdrive-backup-sync.sh`, run via
`hqueex-gdrive-sync.timer`/`.service`, which does two `rclone copy` legs:

1. **JSON metadata**: the whole `backups/` folder to the same Google Drive
   folder structure used locally (`My Drive/H-Queex — Working Documents/
   H-Queex Hub/Backups/`).
2. **File binaries**: everything under the Graph `H-Queex Hub Documents`
   root on OneDrive (Receipts, SOPs, Delivery Logs, Documents, and any
   future category — see `_graph_category_folder()` in `app.py`) is synced
   remote-to-remote, OneDrive → Google Drive, to `My Drive/H-Queex —
   Working Documents/H-Queex Hub/Documents-Binaries/`. Added 13 Aug 2026 —
   before this, only the JSON metadata describing a file was backed up
   offsite, not the file itself; a simultaneous loss of OneDrive and the
   server disk would have destroyed the file binaries with no recovery
   path.

Both legs use `copy`, not `sync` — deliberately never deletes anything on
the Drive side, so the offsite copy survives local/OneDrive deletion or
corruption. `backups/` is locally pruned to `BACKUP_RETENTION_DAYS` (30
days) but the offsite Backups/ copy is not pruned to match. Its combined
result (both legs) is written to `gdrive-sync-status.json` and shown on the
dashboard/Settings pages via `_load_gdrive_sync_status()` — a genuine
failure in *either* leg (not "hasn't run yet") does trigger the warning
banner, since a half-completed offsite backup isn't something the dashboard
should show as green.

**Destination account**: `hqueexbackups@gmail.com`, a dedicated free-tier
Google account created specifically for this — deliberately separate from
Hev's personal Gmail, so business backup data (which may include client
data once real engagements start) never sits in a personal account. This
was a deliberate account-separation fix made 13 Aug 2026 (previously the
`gdrive-hqueex` remote pointed at Hev's personal Google account); see git
history for the migration.

**One-time setup on a new server:**

```bash
sudo apt-get install -y rclone
```

Then configure two rclone remotes — both need an interactive OAuth consent
and can't be scripted or done by an agent on your behalf (though the actual
`rclone config create`/`update` calls that consume the resulting token can
be run non-interactively with `--non-interactive --continue`, once the
token itself has been obtained interactively elsewhere):

- **`gdrive-hqueex`** (Google Drive, full `drive` scope — not `drive.file`,
  since the sync target is an existing folder tree, not one rclone creates
  itself): run `rclone authorize "drive" --drive-scope drive` on a machine
  with a browser, signed in as `hqueexbackups@gmail.com`, then feed the
  resulting token to `rclone config create`/`update` on the server.
- **`onedrive-hqueex`** (OneDrive, business drive, signed in as
  `hmartire@h-queex.com` — this is a *separate* rclone-native OAuth app
  registration, unrelated to the Hub's own Graph app used for Documents
  uploads and lead-notification email): run `rclone authorize "onedrive"`
  on a machine with a browser, then feed the resulting token to
  `rclone config create onedrive-hqueex onedrive drive_type=business
  region=global --non-interactive`, answering the resulting state-machine
  prompts (declining token refresh, declining Shared Drive, confirming the
  one business drive found).

Neither remote's OAuth token/refresh-token should ever be pasted into a
chat session or committed to git — treat them exactly like the Graph
`GRAPH_REFRESH_TOKEN`.

Then install and enable the timer:

```bash
sudo cp scripts/gdrive-backup-sync.sh /opt/hqueex-hub/scripts/
sudo chmod +x /opt/hqueex-hub/scripts/gdrive-backup-sync.sh
sudo cp scripts/hqueex-gdrive-sync.service scripts/hqueex-gdrive-sync.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now hqueex-gdrive-sync.timer
```

Test a run immediately rather than waiting for 03:00:
`sudo systemctl start hqueex-gdrive-sync.service`, then check
`cat /opt/hqueex-hub/gdrive-sync-status.json` and
`journalctl -u hqueex-gdrive-sync.service`.

## Known gaps / not yet done

- **No CI/CD** — deployment is a manual `git pull` + `systemctl restart` over
  SSH (see §3). No GitHub Actions workflow triggers this automatically yet.
- **`ANTHROPIC_API_KEY` not set** — OCR receipt upload will fail on this
  instance until it's added to `/etc/hqueex-hub/hqueex-hub.env`.
- **Single gunicorn instance, no load balancing** — fine for a test/pre-launch
  instance; revisit `--workers` count and consider multiple app servers behind
  Caddy if real traffic before launch warrants it.
- **Rate limiting still uses in-memory storage** (`storage_uri="memory://"` in
  `app.py`), same caveat as noted in `session-handover.md` — resets on
  gunicorn restart, doesn't share state across the 3 worker processes. Not a
  problem for the login/API rate limits' intended purpose at this scale, but
  worth knowing.
