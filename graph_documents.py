"""Microsoft Graph client for the Company Documents feature and lead-notification email.

Delegated OAuth (authorization-code + refresh-token), scoped to the
hmartire@h-queex.com account via the `Files.ReadWrite` (OneDrive) and
`Mail.Send` (lead notification emails) permissions — the refresh token is
obtained once via `scripts/graph_authorize.py` and self-renews on every use
from then on (Microsoft Entra ID does not invalidate the previous refresh
token on rotation, so a failed persist of the new one is recoverable, not a
lockout — see docs/deployment.md).

This module is deliberately standalone (no Flask/app.py imports) so
`scripts/graph_authorize.py` can exercise it before app.py ever wires it in.
"""

from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_TOKEN_URL = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
GRAPH_SCOPE = "offline_access Files.ReadWrite Mail.Send"
DOCUMENTS_ROOT_FOLDER = "H-Queex Hub Documents"

_ENV_KEYS = ("GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET")

_token_cache: dict[str, Any] = {"access_token": None, "expires_at": 0.0}


class GraphAuthError(RuntimeError):
    """Credentials missing/invalid, or the refresh token can't be redeemed."""


class GraphRequestError(RuntimeError):
    """A Graph API call itself failed (upload, download, list, etc.)."""


def _env_path() -> Path:
    """Where the rotated refresh token gets persisted. Locally this is the
    project .env that python-dotenv loads. On the server there is no .env —
    credentials come from systemd's EnvironmentFile= — so GRAPH_ENV_FILE must
    be set (to /etc/hqueex-hub/hqueex-hub.env) or a rotated token would get
    written to a file nothing ever reads back. Also overridable by tests so
    they never touch a real file."""
    override = (os.environ.get("GRAPH_ENV_FILE") or "").strip()
    if override:
        return Path(override)
    return Path(__file__).resolve().parent / ".env"


def require_client_credentials() -> tuple[str, str, str]:
    values = {key: (os.environ.get(key) or "").strip() for key in _ENV_KEYS}
    missing = [key for key, value in values.items() if not value]
    if missing:
        raise GraphAuthError(
            "Microsoft Graph isn't configured — missing " + ", ".join(missing) + " in .env"
        )
    return values["GRAPH_TENANT_ID"], values["GRAPH_CLIENT_ID"], values["GRAPH_CLIENT_SECRET"]


def _atomic_write_env_var(key: str, value: str) -> None:
    """Rewrite .env with `key` set to `value`, via temp-file + os.replace so a
    partial write can never leave the file truncated or corrupted."""
    env_path = _env_path()
    existing_lines: list[str] = []
    if env_path.exists():
        existing_lines = env_path.read_text(encoding="utf-8").splitlines(keepends=True)

    new_lines: list[str] = []
    found = False
    prefix = f"{key}="
    for line in existing_lines:
        if line.rstrip("\n").startswith(prefix):
            new_lines.append(f"{key}={value}\n")
            found = True
        else:
            new_lines.append(line if line.endswith("\n") else line + "\n")
    if not found:
        new_lines.append(f"{key}={value}\n")

    fd, tmp_path_str = tempfile.mkstemp(dir=str(env_path.parent), prefix=".env.tmp-")
    tmp_path = Path(tmp_path_str)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as tmp_file:
            tmp_file.writelines(new_lines)
        os.replace(tmp_path, env_path)
    except Exception:
        try:
            tmp_path.unlink()
        except OSError:
            pass
        raise


def persist_refresh_token(refresh_token: str) -> None:
    """Write the refresh token to .env and to this process's own environment,
    so both future runs and the rest of the current run see it immediately."""
    _atomic_write_env_var("GRAPH_REFRESH_TOKEN", refresh_token)
    os.environ["GRAPH_REFRESH_TOKEN"] = refresh_token


def _short_error(response: requests.Response) -> str:
    try:
        payload = response.json()
        message = payload.get("error_description") or payload.get("error", {}).get("message")
        if message:
            return str(message)
    except ValueError:
        pass
    return response.text[:300]


def redeem_authorization_code(code: str, redirect_uri: str) -> dict[str, Any]:
    """One-time exchange used only by scripts/graph_authorize.py."""
    tenant_id, client_id, client_secret = require_client_credentials()
    response = requests.post(
        GRAPH_TOKEN_URL.format(tenant=tenant_id),
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
            "scope": GRAPH_SCOPE,
        },
        timeout=15,
    )
    if response.status_code != 200:
        raise GraphAuthError(f"Microsoft Graph code exchange failed ({response.status_code}): {_short_error(response)}")
    return response.json()


def _redeem_refresh_token(refresh_token: str) -> dict[str, Any]:
    tenant_id, client_id, client_secret = require_client_credentials()
    response = requests.post(
        GRAPH_TOKEN_URL.format(tenant=tenant_id),
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "scope": GRAPH_SCOPE,
        },
        timeout=15,
    )
    if response.status_code != 200:
        raise GraphAuthError(f"Microsoft Graph token refresh failed ({response.status_code}): {_short_error(response)}")
    return response.json()


def get_access_token(*, force_refresh: bool = False) -> str:
    now = time.time()
    if not force_refresh and _token_cache["access_token"] and now < _token_cache["expires_at"] - 60:
        return _token_cache["access_token"]

    refresh_token = (os.environ.get("GRAPH_REFRESH_TOKEN") or "").strip()
    if not refresh_token:
        raise GraphAuthError(
            "Microsoft Graph isn't connected yet — GRAPH_REFRESH_TOKEN is missing. "
            "Run scripts/graph_authorize.py to complete the one-time consent."
        )

    token_data = _redeem_refresh_token(refresh_token)
    new_refresh_token = token_data.get("refresh_token")
    if new_refresh_token and new_refresh_token != refresh_token:
        # Entra doesn't invalidate the old refresh token on rotation, so if this
        # write fails the old one on disk keeps working until its own 90-day
        # lifetime elapses — a failed persist here is not an immediate lockout.
        persist_refresh_token(new_refresh_token)

    access_token = token_data.get("access_token")
    if not access_token:
        raise GraphAuthError("Microsoft Graph token refresh succeeded but returned no access_token")

    _token_cache["access_token"] = access_token
    _token_cache["expires_at"] = now + int(token_data.get("expires_in", 3600))
    return access_token


def _auth_headers() -> dict[str, str]:
    return {"Authorization": f"Bearer {get_access_token()}"}


def _request(method: str, url: str, **kwargs: Any) -> requests.Response:
    headers = kwargs.pop("headers", {}) or {}
    headers.update(_auth_headers())
    response = requests.request(method, url, headers=headers, timeout=30, **kwargs)
    if response.status_code == 401:
        # Access token may have expired between the cache check and this call —
        # force one refresh and retry once before giving up.
        headers["Authorization"] = f"Bearer {get_access_token(force_refresh=True)}"
        response = requests.request(method, url, headers=headers, timeout=30, **kwargs)
    return response


def ensure_folder(path: str) -> None:
    """Idempotent: creating a folder that already exists is treated as success,
    same spirit as the app's other idempotent migrations."""
    if "/" in path:
        parent_path, folder_name = path.rsplit("/", 1)
    else:
        parent_path, folder_name = "", path

    if parent_path:
        url = f"{GRAPH_API_BASE}/me/drive/root:/{quote(parent_path)}:/children"
    else:
        url = f"{GRAPH_API_BASE}/me/drive/root/children"

    response = _request(
        "POST",
        url,
        json={"name": folder_name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
    )
    if response.status_code in (200, 201):
        return
    if response.status_code == 409:
        return
    raise GraphRequestError(f"Could not create OneDrive folder '{path}' ({response.status_code}): {_short_error(response)}")


def upload_file(folder_path: str, filename: str, content: bytes) -> dict[str, Any]:
    url = f"{GRAPH_API_BASE}/me/drive/root:/{quote(folder_path)}/{quote(filename)}:/content"
    response = _request("PUT", url, data=content, headers={"Content-Type": "application/octet-stream"})
    if response.status_code not in (200, 201):
        raise GraphRequestError(f"Upload to OneDrive failed ({response.status_code}): {_short_error(response)}")
    return response.json()


def download_file(item_id: str) -> bytes:
    url = f"{GRAPH_API_BASE}/me/drive/items/{quote(item_id)}/content"
    response = _request("GET", url)
    if response.status_code != 200:
        raise GraphRequestError(f"Download from OneDrive failed ({response.status_code}): {_short_error(response)}")
    return response.content


def send_mail(to_address: str, subject: str, body_text: str) -> None:
    """Sends as hmartire@h-queex.com via /me/sendMail. Requires the Mail.Send
    delegated permission on top of Files.ReadWrite — see GRAPH_SCOPE."""
    url = f"{GRAPH_API_BASE}/me/sendMail"
    response = _request(
        "POST",
        url,
        json={
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body_text},
                "toRecipients": [{"emailAddress": {"address": to_address}}],
            },
            "saveToSentItems": True,
        },
    )
    if response.status_code != 202:
        raise GraphRequestError(f"Sending mail via Graph failed ({response.status_code}): {_short_error(response)}")


def delete_file(item_id: str) -> None:
    url = f"{GRAPH_API_BASE}/me/drive/items/{quote(item_id)}"
    response = _request("DELETE", url)
    if response.status_code not in (204, 404):
        raise GraphRequestError(f"Delete from OneDrive failed ({response.status_code}): {_short_error(response)}")


def move_file(item_id: str, new_folder_path: str) -> dict[str, Any]:
    """Used when a document's category changes — moves the driveItem to the
    new category folder rather than re-uploading it."""
    parent_url = f"{GRAPH_API_BASE}/me/drive/root:/{quote(new_folder_path)}"
    parent_lookup = _request("GET", parent_url)
    if parent_lookup.status_code != 200:
        raise GraphRequestError(
            f"Could not resolve OneDrive folder '{new_folder_path}' ({parent_lookup.status_code}): {_short_error(parent_lookup)}"
        )
    parent_id = parent_lookup.json()["id"]

    url = f"{GRAPH_API_BASE}/me/drive/items/{quote(item_id)}"
    response = _request("PATCH", url, json={"parentReference": {"id": parent_id}})
    if response.status_code != 200:
        raise GraphRequestError(f"Move within OneDrive failed ({response.status_code}): {_short_error(response)}")
    return response.json()


def list_folder(path: str) -> list[dict[str, Any]]:
    url = f"{GRAPH_API_BASE}/me/drive/root:/{quote(path)}:/children"
    response = _request("GET", url)
    if response.status_code != 200:
        raise GraphRequestError(f"Listing OneDrive folder '{path}' failed ({response.status_code}): {_short_error(response)}")
    return response.json().get("value", [])


def whoami() -> dict[str, Any]:
    response = _request("GET", f"{GRAPH_API_BASE}/me")
    if response.status_code != 200:
        raise GraphRequestError(f"Could not verify Microsoft Graph identity ({response.status_code}): {_short_error(response)}")
    return response.json()
