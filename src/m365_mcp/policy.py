"""Im Code erzwungene Richtlinien. Reine Funktionen ohne I/O."""

from collections.abc import Iterable

from m365_mcp.config import Config
from m365_mcp.errors import PolicyError


def is_internal(address: str, domains: Iterable[str]) -> bool:
    address = address.strip().lower()
    if address.count("@") != 1:
        return False
    local, domain = address.split("@")
    return bool(local) and domain in {d.lower() for d in domains}


def clean_addresses(values: list[str] | None, field: str) -> list[str]:
    out = []
    for value in values or []:
        value = value.strip()
        if not value:
            continue
        if value.count("@") != 1 or value.startswith("@") or value.endswith("@"):
            raise PolicyError(f"{field}: {value!r} ist keine gültige Adresse")
        out.append(value)
    return out


def recipients(addresses: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": a}} for a in addresses]


def addresses_of(items: list[dict] | None) -> list[str]:
    return [i["emailAddress"]["address"] for i in items or [] if i.get("emailAddress")]


def apply_external_mail_policy(
    cfg: Config, to: list[str], cc: list[str], bcc: list[str]
) -> tuple[list[str], list[str] | None, list[str]]:
    """Ist irgendein Empfänger extern, kommen Reply-To und CC aus der Konfiguration dazu."""
    if all(is_internal(a, cfg.internal_domains) for a in [*to, *cc, *bcc]):
        return cc, None, []
    present = {a.lower() for a in cc}
    new_cc = list(cc)
    for a in cfg.external_cc:
        if a.lower() not in present:
            present.add(a.lower())
            new_cc.append(a)
    return new_cc, [cfg.reply_to], ["reply_to", "cc"]


def check_attendees_internal(cfg: Config, attendees: list[str]) -> None:
    external = [a for a in attendees if not is_internal(a, cfg.internal_domains)]
    if external:
        raise PolicyError(
            "Externe Teilnehmer sind nicht erlaubt (es würden Einladungen nach außen verschickt): "
            + ", ".join(external)
        )


def calendar_address(cfg: Config, key: str) -> str:
    if key == "self":
        return cfg.account
    if key in cfg.calendars:
        return cfg.calendars[key]
    allowed = ", ".join(["self", *cfg.calendars])
    raise PolicyError(f"Kalender {key!r} ist nicht freigegeben; erlaubt: {allowed}")


def calendar_root(cfg: Config, key: str) -> str:
    address = calendar_address(cfg, key)
    return "/me" if key == "self" else f"/users/{address}"


def plan_id(cfg: Config, key: str) -> str:
    if key in cfg.plans:
        return cfg.plans[key]
    raise PolicyError(f"Plan {key!r} ist nicht freigegeben; erlaubt: {', '.join(cfg.plans)}")


def assignee_id(cfg: Config, key: str) -> str:
    if key in cfg.assignees:
        return cfg.assignees[key]
    allowed = ", ".join(["self", *cfg.assignees])
    raise PolicyError(f"Zuweisung an {key!r} nicht möglich; erlaubt: {allowed}")
