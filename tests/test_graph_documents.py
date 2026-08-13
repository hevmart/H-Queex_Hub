import os

import pytest

import graph_documents as graph


@pytest.fixture(autouse=True)
def _isolated_env(tmp_path, monkeypatch):
    """Every test gets its own throwaway .env and starts with no cached token,
    so nothing here can ever touch the real project .env or leak state
    between tests."""
    env_file = tmp_path / ".env"
    env_file.write_text(
        "GRAPH_TENANT_ID=test-tenant\nGRAPH_CLIENT_ID=test-client\nGRAPH_CLIENT_SECRET=test-secret\n",
        encoding="utf-8",
    )
    monkeypatch.setattr(graph, "_env_path", lambda: env_file)
    monkeypatch.setenv("GRAPH_TENANT_ID", "test-tenant")
    monkeypatch.setenv("GRAPH_CLIENT_ID", "test-client")
    monkeypatch.setenv("GRAPH_CLIENT_SECRET", "test-secret")
    monkeypatch.delenv("GRAPH_REFRESH_TOKEN", raising=False)
    graph._token_cache["access_token"] = None
    graph._token_cache["expires_at"] = 0.0
    yield env_file


class _FakeResponse:
    def __init__(self, status_code, json_data=None, text=""):
        self.status_code = status_code
        self._json_data = json_data or {}
        self.text = text or ""

    def json(self):
        if self._json_data is None:
            raise ValueError("no json")
        return self._json_data


def test_require_client_credentials_missing(monkeypatch):
    monkeypatch.delenv("GRAPH_TENANT_ID", raising=False)
    with pytest.raises(graph.GraphAuthError, match="GRAPH_TENANT_ID"):
        graph.require_client_credentials()


def test_atomic_write_env_var_adds_new_key(_isolated_env):
    graph._atomic_write_env_var("GRAPH_REFRESH_TOKEN", "abc123")
    content = _isolated_env.read_text(encoding="utf-8")
    assert "GRAPH_REFRESH_TOKEN=abc123" in content
    assert "GRAPH_TENANT_ID=test-tenant" in content  # untouched


def test_atomic_write_env_var_replaces_existing_key(_isolated_env):
    graph._atomic_write_env_var("GRAPH_REFRESH_TOKEN", "first")
    graph._atomic_write_env_var("GRAPH_REFRESH_TOKEN", "second")
    content = _isolated_env.read_text(encoding="utf-8")
    assert content.count("GRAPH_REFRESH_TOKEN=") == 1
    assert "GRAPH_REFRESH_TOKEN=second" in content
    assert "GRAPH_REFRESH_TOKEN=first" not in content


def test_persist_refresh_token_updates_process_env(_isolated_env):
    graph.persist_refresh_token("new-token")
    assert os.environ["GRAPH_REFRESH_TOKEN"] == "new-token"
    assert "GRAPH_REFRESH_TOKEN=new-token" in _isolated_env.read_text(encoding="utf-8")


def test_get_access_token_raises_without_refresh_token():
    with pytest.raises(graph.GraphAuthError, match="not connected|GRAPH_REFRESH_TOKEN"):
        graph.get_access_token()


def test_get_access_token_persists_rotated_refresh_token(monkeypatch, _isolated_env):
    monkeypatch.setenv("GRAPH_REFRESH_TOKEN", "old-refresh")

    def fake_post(url, data=None, timeout=None):
        assert data["grant_type"] == "refresh_token"
        assert data["refresh_token"] == "old-refresh"
        return _FakeResponse(200, {"access_token": "access-1", "refresh_token": "new-refresh", "expires_in": 3600})

    monkeypatch.setattr(graph.requests, "post", fake_post)

    token = graph.get_access_token()
    assert token == "access-1"
    assert os.environ["GRAPH_REFRESH_TOKEN"] == "new-refresh"
    assert "GRAPH_REFRESH_TOKEN=new-refresh" in _isolated_env.read_text(encoding="utf-8")


def test_get_access_token_uses_cache_without_refetching(monkeypatch, _isolated_env):
    monkeypatch.setenv("GRAPH_REFRESH_TOKEN", "old-refresh")
    call_count = {"n": 0}

    def fake_post(url, data=None, timeout=None):
        call_count["n"] += 1
        return _FakeResponse(200, {"access_token": "access-1", "refresh_token": "old-refresh", "expires_in": 3600})

    monkeypatch.setattr(graph.requests, "post", fake_post)

    graph.get_access_token()
    graph.get_access_token()
    assert call_count["n"] == 1


def test_get_access_token_raises_on_failed_refresh(monkeypatch, _isolated_env):
    monkeypatch.setenv("GRAPH_REFRESH_TOKEN", "old-refresh")

    def fake_post(url, data=None, timeout=None):
        return _FakeResponse(400, {"error_description": "invalid_grant"})

    monkeypatch.setattr(graph.requests, "post", fake_post)

    with pytest.raises(graph.GraphAuthError, match="invalid_grant"):
        graph.get_access_token()


def _stub_access_token(monkeypatch):
    monkeypatch.setattr(graph, "get_access_token", lambda force_refresh=False: "fake-access-token")


def test_ensure_folder_treats_conflict_as_success(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        assert method == "POST"
        return _FakeResponse(409, {"error": {"message": "already exists"}})

    monkeypatch.setattr(graph.requests, "request", fake_request)
    graph.ensure_folder("H-Queex Hub Documents")  # should not raise


def test_ensure_folder_raises_on_real_failure(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        return _FakeResponse(403, {"error": {"message": "forbidden"}})

    monkeypatch.setattr(graph.requests, "request", fake_request)
    with pytest.raises(graph.GraphRequestError, match="forbidden"):
        graph.ensure_folder("H-Queex Hub Documents")


def test_upload_file_returns_drive_item(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        assert method == "PUT"
        assert kwargs["data"] == b"hello"
        return _FakeResponse(201, {"id": "item-1", "name": "test.txt"})

    monkeypatch.setattr(graph.requests, "request", fake_request)
    result = graph.upload_file("H-Queex Hub Documents/Other", "test.txt", b"hello")
    assert result["id"] == "item-1"


def test_upload_file_raises_on_failure(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        return _FakeResponse(507, {"error": {"message": "quota exceeded"}})

    monkeypatch.setattr(graph.requests, "request", fake_request)
    with pytest.raises(graph.GraphRequestError, match="quota exceeded"):
        graph.upload_file("H-Queex Hub Documents/Other", "test.txt", b"hello")


def test_download_file_returns_bytes(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    class _ContentResponse(_FakeResponse):
        content = b"file-bytes"

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        assert method == "GET"
        return _ContentResponse(200)

    monkeypatch.setattr(graph.requests, "request", fake_request)
    assert graph.download_file("item-1") == b"file-bytes"


def test_delete_file_treats_404_as_success(monkeypatch, _isolated_env):
    _stub_access_token(monkeypatch)

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        assert method == "DELETE"
        return _FakeResponse(404)

    monkeypatch.setattr(graph.requests, "request", fake_request)
    graph.delete_file("item-1")  # should not raise


def test_request_retries_once_on_401(monkeypatch, _isolated_env):
    monkeypatch.setenv("GRAPH_REFRESH_TOKEN", "old-refresh")
    calls = {"n": 0}

    def fake_post(url, data=None, timeout=None):
        calls["n"] += 1
        return _FakeResponse(200, {"access_token": f"access-{calls['n']}", "refresh_token": "old-refresh", "expires_in": 3600})

    request_calls = {"n": 0}

    def fake_request(method, url, headers=None, timeout=None, **kwargs):
        request_calls["n"] += 1
        if request_calls["n"] == 1:
            return _FakeResponse(401, {"error": {"message": "expired"}})
        return _FakeResponse(200, {"value": []})

    monkeypatch.setattr(graph.requests, "post", fake_post)
    monkeypatch.setattr(graph.requests, "request", fake_request)

    result = graph.list_folder("H-Queex Hub Documents")
    assert result == []
    assert request_calls["n"] == 2
    assert calls["n"] == 2  # forced a second token fetch after the 401
