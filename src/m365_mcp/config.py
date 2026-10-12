"""Konfiguration (TOML). Alles Tenant-Spezifische steht nur hier, nie im Code."""

import re
import tomllib
import zoneinfo
from dataclasses import dataclass
from pathlib import Path

from m365_mcp.errors import ConfigError

_KEY = re.compile(r"^[a-z0-9]+\Z")
_ADDRESS = re.compile(r"^[^@\s]+@[^@\s]+\Z")
# Landen in der Authority-URL bzw. als Client-ID bei MSAL: GUID oder Domänenname, keine Pfadzeichen.
_IDENT = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]*\Z")


@dataclass(frozen=True)
class Config:
    path: Path
    tenant_id: str
    client_id: str
    account: str
    internal_domains: tuple[str, ...]
    reply_to: str
    external_cc: tuple[str, ...]
    calendars: dict[str, str]
    plans: dict[str, str]
    assignees: dict[str, str]
    timezone: str = "Europe/Berlin"
    max_body_chars: int = 8000

    @property
    def directory(self) -> Path:
        return self.path.parent

    @property
    def token_cache(self) -> Path:
        return self.directory / "token_cache.json"

    @property
    def token_lock(self) -> Path:
        return self.directory / "token_cache.lock"


def _text(table: dict, key: str, where: str = "") -> str:
    value = table.get(key)
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{where}{key} fehlt oder ist leer")
    return value.strip()


def _ident(table: dict, key: str) -> str:
    value = _text(table, key)
    if not _IDENT.match(value) or ".." in value:
        raise ConfigError(f"{key} {value!r} ungültig: erwartet GUID oder Domänenname")
    return value


def _internal_address(value: str, domains: tuple[str, ...], where: str) -> str:
    value = value.strip().lower()
    if not _ADDRESS.match(value) or value.rsplit("@", 1)[1] not in domains:
        raise ConfigError(f"{where}: {value} muss eine interne Adresse sein (internal_domains)")
    return value


def _mapping(value: object, where: str) -> dict[str, str]:
    if not isinstance(value, dict):
        raise ConfigError(f"{where} muss eine Tabelle sein")
    out: dict[str, str] = {}
    for key, item in value.items():
        if not _KEY.match(key) or key == "self":
            raise ConfigError(f"{where}: Schlüssel {key!r} ungültig (nur a-z und 0-9, nicht 'self')")
        if not isinstance(item, str) or not item.strip():
            raise ConfigError(f"{where}.{key} ist leer")
        out[key] = item.strip()
    return out


def load_config(path: Path) -> Config:
    try:
        raw = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, tomllib.TOMLDecodeError) as exc:
        raise ConfigError(f"Konfiguration {path} nicht lesbar: {exc}") from exc

    domains_raw = raw.get("internal_domains")
    if (
        not isinstance(domains_raw, list)
        or not domains_raw
        or not all(isinstance(d, str) and d.strip() for d in domains_raw)
    ):
        raise ConfigError("internal_domains muss eine nicht leere Liste von Domains sein")
    domains = tuple(d.strip().lower() for d in domains_raw)

    external = raw.get("external_mail")
    if not isinstance(external, dict):
        raise ConfigError("Tabelle [external_mail] fehlt")
    reply_to = _internal_address(_text(external, "reply_to", "external_mail."), domains, "reply_to")
    cc_raw = external.get("cc")
    # Pflicht und nicht leer: die CC-Regel für externe Empfänger darf nicht per Konfiguration entfallen.
    if not isinstance(cc_raw, list) or not cc_raw or not all(isinstance(a, str) for a in cc_raw):
        raise ConfigError("external_mail.cc muss eine nicht leere Liste von Adressen sein")
    external_cc = tuple(_internal_address(a, domains, "external_mail.cc") for a in cc_raw)

    calendars = {
        key: _internal_address(addr, domains, f"calendars.{key}")
        for key, addr in _mapping(raw.get("calendars", {}), "calendars").items()
    }
    planner = raw.get("planner", {})
    if not isinstance(planner, dict):
        raise ConfigError("[planner] muss eine Tabelle sein")

    max_body = raw.get("max_body_chars", 8000)
    if not isinstance(max_body, int) or isinstance(max_body, bool) or max_body <= 0:
        raise ConfigError("max_body_chars muss eine positive Zahl sein")

    timezone = raw.get("timezone", "Europe/Berlin")
    if not isinstance(timezone, str) or not timezone.strip():
        raise ConfigError("timezone muss ein nicht leerer IANA-Zonenname sein")
    timezone = timezone.strip()
    try:
        zoneinfo.ZoneInfo(timezone)
    except (zoneinfo.ZoneInfoNotFoundError, ValueError) as exc:
        raise ConfigError(f"timezone {timezone!r} ist keine bekannte IANA-Zone") from exc

    return Config(
        path=path,
        tenant_id=_ident(raw, "tenant_id"),
        client_id=_ident(raw, "client_id"),
        account=_internal_address(_text(raw, "account"), domains, "account"),
        internal_domains=domains,
        reply_to=reply_to,
        external_cc=external_cc,
        calendars=calendars,
        plans=_mapping(planner.get("plans", {}), "planner.plans"),
        assignees=_mapping(planner.get("assignees", {}), "planner.assignees"),
        timezone=timezone,
        max_body_chars=max_body,
    )
