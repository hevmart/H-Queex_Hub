"""One-time interactive Microsoft Graph consent for H-Queex Hub Documents.

Run locally (needs a real browser to sign in):

    .venv/Scripts/python.exe scripts/graph_authorize.py

Opens a browser for you to sign in as hmartire@h-queex.com and approve the
Files.ReadWrite consent, exchanges the resulting code for tokens, writes the
refresh token into .env (GRAPH_REFRESH_TOKEN), then exercises a full
create-folder / upload / list / delete cycle against the real OneDrive to
prove the whole path works before app.py ever touches it.

Never run this against the server — it needs an interactive browser session
on the machine it's run from, same as the rclone Google Drive setup.
"""

from __future__ import annotations

import secrets
import sys
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse
import webbrowser

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from dotenv import load_dotenv  # noqa: E402

load_dotenv()

import graph_documents as graph  # noqa: E402

REDIRECT_URI = "http://localhost:53682/callback"
AUTHORIZE_URL_TEMPLATE = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"

_result: dict[str, str] = {}


class _CallbackHandler(BaseHTTPRequestHandler):
    def log_message(self, *args: object) -> None:  # silence default request logging
        pass

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return
        params = parse_qs(parsed.query)
        _result["code"] = params.get("code", [""])[0]
        _result["state"] = params.get("state", [""])[0]
        _result["error"] = params.get("error_description", [""])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        if _result["code"]:
            self.wfile.write(
                b"<html><body><h2>H-Queex Hub connected to Microsoft Graph.</h2>"
                b"You can close this window and return to the terminal.</body></html>"
            )
        else:
            self.wfile.write(b"<html><body><h2>Something went wrong.</h2>Check the terminal for details.</body></html>")


def main() -> int:
    tenant_id, client_id, _client_secret = graph.require_client_credentials()

    state = secrets.token_urlsafe(24)
    query = urlencode(
        {
            "client_id": client_id,
            "response_type": "code",
            "redirect_uri": REDIRECT_URI,
            "response_mode": "query",
            "scope": graph.GRAPH_SCOPE,
            "state": state,
            "login_hint": "hmartire@h-queex.com",
        }
    )
    authorize_url = f"{AUTHORIZE_URL_TEMPLATE.format(tenant=tenant_id)}?{query}"

    print("Opening your browser to sign in as hmartire@h-queex.com and approve access...")
    print(f"If it doesn't open automatically, visit:\n{authorize_url}\n")
    webbrowser.open(authorize_url)

    server = HTTPServer(("localhost", 53682), _CallbackHandler)
    server.handle_request()  # blocks for exactly one request, then returns
    server.server_close()

    if _result.get("error"):
        print(f"Consent failed: {_result['error']}")
        return 1
    if not _result.get("code"):
        print("No authorization code received. Aborting.")
        return 1
    if _result.get("state") != state:
        print("State mismatch on the callback — possible tampering. Aborting without exchanging the code.")
        return 1

    print("Exchanging authorization code for tokens...")
    tokens = graph.redeem_authorization_code(_result["code"], REDIRECT_URI)
    refresh_token = tokens.get("refresh_token")
    if not refresh_token:
        print("Microsoft didn't return a refresh token — check that 'offline_access' is in the requested scope.")
        return 1

    graph.persist_refresh_token(refresh_token)
    print("Refresh token saved to .env (GRAPH_REFRESH_TOKEN).")

    print("Verifying identity...")
    who = graph.whoami()
    signed_in_as = (who.get("mail") or who.get("userPrincipalName") or "").lower()
    print(f"Connected as: {who.get('displayName')} <{signed_in_as}>")
    if signed_in_as != "hmartire@h-queex.com":
        print("WARNING: signed-in account does not match hmartire@h-queex.com — check you signed in with the right account.")

    print("Running a live smoke test: create folder, upload, list, delete...")
    test_folder = f"{graph.DOCUMENTS_ROOT_FOLDER}/_connection-test"
    graph.ensure_folder(graph.DOCUMENTS_ROOT_FOLDER)
    graph.ensure_folder(test_folder)
    stamp = datetime.now(timezone.utc).isoformat(timespec="seconds")
    uploaded = graph.upload_file(
        test_folder,
        "connection-test.txt",
        f"H-Queex Hub Graph connection verified {stamp}\n".encode("utf-8"),
    )
    print(f"  uploaded: {uploaded.get('name')} (id {uploaded.get('id')})")
    listing = graph.list_folder(test_folder)
    print(f"  listed {len(listing)} item(s) in {test_folder}")
    graph.delete_file(str(uploaded["id"]))
    print("  deleted test file — cleanup complete")

    print("\nAll checks passed. Microsoft Graph is connected and ready for step 4 (wiring into the Documents feature).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
