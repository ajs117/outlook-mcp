from unittest import mock

import pytest

from outlook_mcp import auth


class FakeApp:
    def __init__(self, silent_result, device_result):
        self.silent_result = silent_result
        self.device_result = device_result
        self.device_flow_started = False

    def get_accounts(self):
        return [{"username": "me@outlook.com"}]

    def acquire_token_silent(self, scopes, account):
        return self.silent_result

    def initiate_device_flow(self, scopes):
        self.device_flow_started = True
        return {"user_code": "ABC", "message": "go login", "expires_at": 9e12}

    def acquire_token_by_device_flow(self, flow):
        return self.device_result


@pytest.fixture
def run(monkeypatch, tmp_path):
    monkeypatch.setenv("OUTLOOK_MCP_CLIENT_ID", "client-id")
    monkeypatch.setattr(auth, "CACHE_PATH", tmp_path / "cache.json")

    def _run(app):
        with mock.patch.object(auth.msal, "PublicClientApplication", return_value=app):
            return auth.get_access_token()

    return _run


def test_cached_token_skips_device_flow(run):
    app = FakeApp({"access_token": "cached"}, None)
    assert run(app) == "cached"
    assert not app.device_flow_started


@pytest.mark.parametrize("silent_result", [
    None,
    {"error": "invalid_grant", "error_description": "refresh token expired"},
    {"error": "interaction_required"},
])
def test_failed_silent_refresh_falls_back_to_device_flow(run, silent_result):
    app = FakeApp(silent_result, {"access_token": "fresh"})
    assert run(app) == "fresh"
    assert app.device_flow_started
