"""Device-Code-Anmeldung (MSAL) und Token-Cache mit Dateilock.

`serve` und `check` können gleichzeitig ein rotiertes Refresh-Token speichern; deshalb wird der
Cache nur unter exklusivem Lock gelesen, erneuert und atomar ersetzt.
"""

import os
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from functools import partial
from pathlib import Path

import msal

from m365_mcp.config import Config
from m365_mcp.errors import ConfigError, M365Error, ReauthRequired

try:
    import fcntl
except ImportError:  # Windows-Entwicklung; produktiv läuft der Server nur unter Linux
    fcntl = None

SCOPES = [
    "User.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Calendars.ReadWrite.Shared",
    "Tasks.ReadWrite",
]


def ensure_owner(directory: Path) -> None:
    if os.name != "posix":
        return
    owner = directory.stat().st_uid
    if owner != os.geteuid():
        raise ConfigError(
            f"{directory} gehört UID {owner}, der Prozess läuft als UID {os.geteuid()}. "
            "Als Eigentümer starten (z. B. sudo -u hermes …), sonst gehört der Token-Cache danach "
            "dem falschen Benutzer."
        )


@contextmanager
def cache_lock(cfg: Config) -> Iterator[None]:
    fd = os.open(cfg.token_lock, os.O_RDWR | os.O_CREAT, 0o600)
    try:
        if fcntl:
            fcntl.flock(fd, fcntl.LOCK_EX)
        yield
    finally:
        if fcntl:
            fcntl.flock(fd, fcntl.LOCK_UN)
        os.close(fd)


def reload_cache(cfg: Config, cache: msal.SerializableTokenCache) -> None:
    """Cache-Inhalt aus der Datei in ein bestehendes Objekt laden.

    Ein anderer Prozess kann das Refresh-Token inzwischen rotiert haben.
    """
    if not cfg.token_cache.exists():
        return
    try:
        cache.deserialize(cfg.token_cache.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ReauthRequired("Token-Cache unlesbar") from exc
    cache.has_state_changed = False


def load_cache(cfg: Config) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    reload_cache(cfg, cache)
    return cache


def save_cache(cfg: Config, cache: msal.SerializableTokenCache) -> None:
    if not cache.has_state_changed:
        return
    tmp = cfg.token_cache.with_name(cfg.token_cache.name + ".tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    try:
        if hasattr(os, "fchmod"):
            os.fchmod(fd, 0o600)  # eine liegengebliebene .tmp behielte sonst ihren alten Modus
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(cache.serialize())
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, cfg.token_cache)
    finally:
        tmp.unlink(missing_ok=True)
    cache.has_state_changed = False


def _app(cfg: Config, factory: Callable, cache: msal.SerializableTokenCache):
    return factory(
        cfg.client_id,
        authority=f"https://login.microsoftonline.com/{cfg.tenant_id}",
        token_cache=cache,
    )


class TokenProvider:
    """App und Cache-Objekt einmal anlegen: der Konstruktor der App macht Tenant-Discovery per HTTP.

    Pro Aufruf nur den Dateiinhalt unter Lock neu laden (ein anderer Prozess kann rotiert haben).
    """

    def __init__(self, cfg: Config, app_factory: Callable = msal.PublicClientApplication) -> None:
        self._cfg = cfg
        self._factory = app_factory
        self._cache = msal.SerializableTokenCache()
        self._app = None

    def __call__(self, force_refresh: bool = False) -> str:
        with cache_lock(self._cfg):
            reload_cache(self._cfg, self._cache)
            if self._app is None:
                self._app = _app(self._cfg, self._factory, self._cache)
            accounts = self._app.get_accounts(username=self._cfg.account)
            if not accounts:
                raise ReauthRequired(f"kein Konto {self._cfg.account} im Token-Cache")
            result = self._app.acquire_token_silent_with_error(SCOPES, account=accounts[0], force_refresh=force_refresh)
            save_cache(self._cfg, self._cache)
        if not result or "access_token" not in result:
            raise ReauthRequired((result or {}).get("error", "kein Token"))
        return result["access_token"]


# Ohne TTY ist stdout blockgepuffert; ungeflusht erschiene der Device-Code erst nach dem Warten.
_print_flushed = partial(print, flush=True)


def login(
    cfg: Config,
    *,
    app_factory: Callable = msal.PublicClientApplication,
    echo: Callable[[str], object] = _print_flushed,
) -> str:
    """Interaktiv; der Lock wird erst zum Speichern genommen, damit `serve` nicht minutenlang hängt."""
    cache = msal.SerializableTokenCache()
    app = _app(cfg, app_factory, cache)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise M365Error(f"Device-Code-Flow nicht möglich: {flow.get('error_description', flow)}")
    echo(flow["message"])
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise M365Error(f"Anmeldung fehlgeschlagen: {result.get('error_description', result)}")
    user = result.get("id_token_claims", {}).get("preferred_username", "")
    if user.lower() != cfg.account.lower():
        raise M365Error(f"Angemeldet als {user}, erwartet {cfg.account}; Cache nicht gespeichert")
    cache.has_state_changed = True
    with cache_lock(cfg):
        save_cache(cfg, cache)
    return user
