import os
import stat

import pytest

from m365_mcp.auth import SCOPES, TokenProvider, ensure_owner, load_cache, login, save_cache
from m365_mcp.errors import ConfigError, M365Error, ReauthRequired

posix_only = pytest.mark.skipif(os.name != "posix", reason="nur POSIX (fcntl, Modus, UID)")


class FakeApp:
    """Steht für msal.PublicClientApplication; schreibt beim Erfolg in den Cache."""

    result: dict = {}
    accounts: list = [{"username": "olivia@example.org"}]
    login_user = "olivia@example.org"

    def __init__(self, client_id, authority, token_cache):
        self.cache = token_cache
        self.authority = authority

    def get_accounts(self, username=None):
        return [a for a in self.accounts if username is None or a["username"] == username]

    def acquire_token_silent_with_error(self, scopes, account):
        assert scopes == SCOPES
        if "access_token" in self.result:
            self.cache.deserialize('{"AccessToken": {"x": {"secret": "neu"}}}')
            self.cache.has_state_changed = True  # wie bei einem echten Refresh
        return self.result

    def initiate_device_flow(self, scopes):
        return {"user_code": "ABC", "message": "Code ABC eingeben"}

    def acquire_token_by_device_flow(self, flow):
        self.cache.deserialize('{"AccessToken": {"x": {"secret": "login"}}}')
        return {"access_token": "at", "id_token_claims": {"preferred_username": self.login_user}}


def factory(**kw):
    class App(FakeApp):
        pass

    for k, v in kw.items():
        setattr(App, k, v)
    return App


def test_scopes_have_no_reserved_entries():
    assert not {"offline_access", "openid", "profile"} & set(SCOPES)
    assert "Calendars.ReadWrite.Shared" in SCOPES and "Tasks.ReadWrite" in SCOPES


def test_token_provider_returns_token_and_saves_cache(cfg):
    provider = TokenProvider(cfg, app_factory=factory(result={"access_token": "at"}))
    assert provider() == "at"
    assert "neu" in cfg.token_cache.read_text(encoding="utf-8")


def test_token_provider_builds_app_once(cfg):
    created = []

    class Counting(factory(result={"access_token": "at"})):
        def __init__(self, *a, **kw):
            created.append(1)
            super().__init__(*a, **kw)

    provider = TokenProvider(cfg, app_factory=Counting)
    provider()
    provider()
    assert len(created) == 1


def test_corrupt_cache_needs_reauth(cfg):
    cfg.token_cache.write_text("kein json", encoding="utf-8")
    with pytest.raises(ReauthRequired, match="unlesbar"):
        TokenProvider(cfg, app_factory=factory(result={"access_token": "at"}))()


def test_token_provider_without_account_needs_reauth(cfg):
    provider = TokenProvider(cfg, app_factory=factory(accounts=[]))
    with pytest.raises(ReauthRequired, match="kein Konto"):
        provider()


def test_token_provider_invalid_grant_needs_reauth(cfg):
    provider = TokenProvider(cfg, app_factory=factory(result={"error": "invalid_grant"}))
    with pytest.raises(ReauthRequired, match="invalid_grant"):
        provider()


def test_login_saves_cache_and_returns_upn(cfg):
    lines = []
    assert login(cfg, app_factory=factory(), echo=lines.append) == "olivia@example.org"
    assert lines == ["Code ABC eingeben"]
    assert "login" in cfg.token_cache.read_text(encoding="utf-8")


def test_login_with_wrong_account_does_not_save(cfg):
    with pytest.raises(M365Error, match="erwartet olivia@example.org"):
        login(cfg, app_factory=factory(login_user="christian@example.org"), echo=lambda m: None)
    assert not cfg.token_cache.exists()


def test_load_cache_without_file_creates_nothing(cfg):
    load_cache(cfg)
    assert not cfg.token_cache.exists()


@posix_only
def test_cache_file_mode_0600(cfg):
    cache = load_cache(cfg)
    cache.deserialize('{"AccessToken": {}}')
    cache.has_state_changed = True
    save_cache(cfg, cache)
    assert stat.S_IMODE(cfg.token_cache.stat().st_mode) == 0o600


@posix_only
def test_ensure_owner_rejects_foreign_directory(tmp_path, monkeypatch):
    monkeypatch.setattr(os, "geteuid", lambda: os.stat(tmp_path).st_uid + 1)
    with pytest.raises(ConfigError, match="sudo -u"):
        ensure_owner(tmp_path)


@posix_only
def test_lock_serializes_writers(cfg):
    import fcntl

    from m365_mcp.auth import cache_lock

    with cache_lock(cfg):
        fd = os.open(cfg.token_lock, os.O_RDWR)
        try:
            with pytest.raises(BlockingIOError):
                fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)
        finally:
            os.close(fd)
