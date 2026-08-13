# H-Queex Systems Map
*Your single reference point for every tool, account, and access path.*
*Last updated: 11 August 2026 (backup system + UI fixes added)*

---

## Quick Reference Table

| What | Where | How you access it |
|---|---|---|
| Public website | h-queex.com | Just visit it in a browser — anyone can |
| Business management app (Hub) | https://hub.h-queex.com | Log in with your Owner account (email + password you set at `/setup`) |
| Hub's source code | GitHub: hevmart/H-Queex_Hub | VS Code, or github.com in browser |
| Website's source code | GitHub: hevmart/h-queex_website | VS Code, or github.com in browser |
| Website hosting | Netlify | netlify.com login |
| Hub's server | Hetzner Cloud | console.hetzner.com login |
| Domain/DNS | Squarespace Domains | squarespace.com login |
| Business plan documents | Google Drive | drive.google.com |
| Business email | Microsoft 365 | outlook.com or Outlook app |

---

## 1. The Public Website (h-queex.com)

**What it is**: the marketing site visitors see — services, pricing, the contact/intake form.

**Where it lives**: hosted on **Netlify**, which automatically rebuilds and republishes it every time changes are pushed to the `hevmart/h-queex_website` GitHub repo. You never manually "upload" anything — pushing to GitHub *is* publishing.

**How to change it**:
1. Open the `h-queex_website` folder in VS Code
2. Edit `content-model.json` for text/copy changes (never edit `index.html` directly for content — see `docs/ui-standards.md`-equivalent conventions for this repo)
3. Commit and push
4. Netlify picks it up automatically within a minute or two — no separate step needed

**Where to check it's live**: netlify.com → log in → your site → you'll see build/deploy status and history there.

---

## 2. H-Queex Hub (the business app)

**What it is**: your internal system — Finance, CRM, Operations, Compliance, everything you use to actually run the business day to day.

**Where it lives now**: as of today, it runs in two places:
- **Locally**, on your laptop, via `Launch-HQueex-Hub.ps1` — this is your everyday working copy
- **On the server**, at `https://hub.h-queex.com` — this is a live, real, internet-reachable copy for testing ahead of launch

**How to log in**: go to `https://hub.h-queex.com` in any browser, sign in with the Owner account you created during `/setup`. (If you forget that password, that's a "reset" task for a future Claude Code session — flag it if it happens.)

**How to change the app itself**: open the `H-Queex_Hub` folder in VS Code, use Claude Code to make changes, test locally, then it needs a *second* step to reach the live server (see below) — pushing to GitHub does **not** automatically update `hub.h-queex.com` the way it does for the website. That's a manual/Claude-Code-assisted deploy step each time.

---

## 3. The Server (Hetzner)

**What it is**: the actual computer, rented from Hetzner, that runs `hub.h-queex.com` 24/7.

**Login**: console.hetzner.com — this is where you'd go to check if the server is running, see costs, or (rarely) restart it.

**Server details**:
- Name: `hqueex-hub`
- IP address: `167.233.110.170`
- Location: Falkenstein, Germany
- Cost: ~€11.99/month

**Direct technical access** (only needed for deeper troubleshooting, normally Claude Code handles this): SSH in via PowerShell using `ssh deploy@167.233.110.170` (root login is disabled — this is intentional, for security).

**You do not need to log into Hetzner regularly.** Its console is mainly for billing/status checks. Day-to-day, you interact with the *app* at `hub.h-queex.com`, not the server itself.

---

## 4. Domains & DNS (Squarespace)

**What it is**: where `h-queex.com` and `h-queex.ie` were purchased and where their DNS records live.

**Login**: squarespace.com → Domains

**What's pointed where**:
- `h-queex.com` → Netlify (the website)
- `hub.h-queex.com` → `167.233.110.170` (the Hetzner server, via the A record you added)

**When you'd need this**: only if you ever add another subdomain, change hosting providers, or set up business email routing. Rare.

---

## 5. GitHub (the code itself)

**What it is**: where all the actual code lives and its full history is tracked. Both the website and the Hub are separate repositories here.

**Repos**:
- `hevmart/H-Queex_Hub` — the business app
- `hevmart/h-queex_website` — the public site

**How you interact with it**: almost always through VS Code + Claude Code, not github.com directly. You'd only visit github.com itself to browse history or check something visually.

---

## 6. OCR (Receipt Scanning)

**What it is**: not a separate tool you log into — it's a feature *inside* the Hub (Expenses → upload a receipt), powered by Anthropic's API in the background.

**Where the "engine" lives**: an `ANTHROPIC_API_KEY` is stored in two places — your local `.env` file (for when you're running the Hub on your laptop) and `/etc/hqueex-hub/hqueex-hub.env` on the server (for the live version). You don't manage this day-to-day; it just works when you upload a receipt.

**If OCR ever stops working**: that's a "check the API key" task — hand it to Claude Code, don't troubleshoot manually.

---

## 6b. Backups (Two Layers)

There are now two separate, independent backup layers protecting the Hub's data:

1. **Local backup** — every save also copies data into a `backups/<date>/`
   folder on whichever machine is running the Hub (local laptop or the
   server). This alone is not enough — if that machine is lost, this copy
   is lost too.
2. **Offsite backup (real Google Drive)** — a nightly automated job
   (`rclone`, running via systemd timer at 03:00 UTC) copies the server's
   backup files into the actual "H-Queex — Working Documents/H-Queex Hub/
   Backups/" folder in Google Drive. This is genuine offsite protection —
   verified independently in Drive itself on 11 Aug 2026.

**Where to check backup status**: log into the Hub → Settings → Backups
card. You'll see two separate lines — one for the local/live-mount check
(will correctly show ✗ on the server, since the server has no local Google
Drive mount — that's expected, not a fault) and one for the nightly offsite
sync (should show ✓ with a recent timestamp).

**Gap closed (13 Aug 2026)**: uploaded file *binaries* for Documents, Receipts,
SOPs, and Delivery Log now upload directly to OneDrive for Business
(`hmartire@h-queex.com`) via Microsoft Graph, under "H-Queex Hub Documents" →
`Receipts` / `SOPs` / `Delivery Logs` category folders (Documents already had
its own categories). Files there get real OneDrive version history — no
separate backup layer is needed for them. Local/server disk write is still
kept as a fallback during rollout (if the OneDrive upload fails, the local
copy is what's served and a warning is logged — nothing is silently lost).
**Not yet done**: existing historical files already on local/server disk have
not been backfilled into OneDrive — that's a separate follow-up migration,
not yet scheduled.

## 7. Business Plan & Documents (Google Drive)

**What it is**: the master business plan (currently V4.7) and other working documents.

**Location**: "H-Queex — Working Documents/Business Plan/" in Google Drive.

**Rule of thumb**: always check Drive for the latest version before editing — don't assume a copy in a chat or on your laptop is current.

---

## 8. Email (Microsoft 365)

**What it is**: your business mailbox, `hmartire@h-queex.com`.

**Login**: outlook.com, or the Outlook app.

---

## What You Actually Need to Remember Day-to-Day

Realistically, out of everything above, your regular touchpoints are just:

1. **`hub.h-queex.com`** — log in to run the business
2. **VS Code** — open when you want to change something, with Claude Code doing the technical work
3. **Google Drive** — for the business plan
4. **Outlook** — for email

Hetzner, Netlify, Squarespace, and GitHub are all "background" services — you rarely log into them directly. Claude Code operates most of them on your behalf when needed.

---

## If You Forget a Password

| Service | Where credentials likely are |
|---|---|
| Hetzner | Your password manager / email you signed up with |
| Netlify | Your password manager / email you signed up with |
| Squarespace | Your password manager / email you signed up with |
| GitHub | Your password manager |
| Hub Owner login | Only you know this — if lost, ask Claude Code to help reset it via the server |
| Google/Microsoft 365 | Standard account recovery via Google/Microsoft |

**Recommendation**: if you're not already using one, a password manager (like Bitwarden, free) would remove a lot of this stress — one master password unlocks everything else, rather than trying to remember or write down six separate logins.
