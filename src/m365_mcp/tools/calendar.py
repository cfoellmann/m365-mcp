"""Kalender-Werkzeuge. calendar = "self" oder ein Schlüssel aus [calendars].

In fremden Kalendern verschickt Graph Einladungen, Aktualisierungen und Absagen im Namen des
Inhabers. Deshalb dort keine Teilnehmer und kein Ändern/Löschen von Terminen mit Teilnehmern.
Externe Teilnehmer sind in v1 auch im eigenen Kalender ausgeschlossen.
"""

import datetime
import zoneinfo

from m365_mcp.context import Ctx
from m365_mcp.errors import GraphError, PolicyError
from m365_mcp.graph import quote_id
from m365_mcp.policy import (
    addresses_of,
    calendar_address,
    calendar_root,
    check_attendees_internal,
    clean_addresses,
    is_internal,
)

_SELECT = "id,subject,start,end,location,attendees,isAllDay,type,seriesMasterId,organizer,bodyPreview"


def _tz(ctx: Ctx) -> dict:
    return {"Prefer": f'outlook.timezone="{ctx.cfg.timezone}"'}


def _local(ctx: Ctx, value: str) -> datetime.datetime:
    """Zeitangabe prüfen: ISO 8601 ohne Zone. Mit Zone oder ungültig → PolicyError, vor jedem Graph-Aufruf."""
    try:
        moment = datetime.datetime.fromisoformat(value)
    except ValueError as exc:
        raise PolicyError(f"Zeitangabe {value!r} ungültig, erwartet ISO 8601 ohne Zone (2026-10-07T09:00:00)") from exc
    if moment.tzinfo is not None:
        raise PolicyError(
            f"Zeitangabe {value!r} enthält eine Zone; erwartet ISO 8601 ohne Zone "
            f"(2026-10-07T09:00:00), interpretiert in {ctx.cfg.timezone}"
        )
    return moment


def _when(ctx: Ctx, value: str) -> dict:
    _local(ctx, value)
    return {"dateTime": value, "timeZone": ctx.cfg.timezone}


def _with_offset(ctx: Ctx, value: str) -> str:
    """calendarView liest Zeiten ohne Offset als UTC; Prefer: outlook.timezone wirkt nur auf die Antwort."""
    moment = _local(ctx, value).replace(tzinfo=zoneinfo.ZoneInfo(ctx.cfg.timezone))
    return moment.isoformat()


def _event(e: dict) -> dict:
    return {
        "id": e.get("id"),
        "subject": e.get("subject"),
        "start": (e.get("start") or {}).get("dateTime"),
        "end": (e.get("end") or {}).get("dateTime"),
        "all_day": e.get("isAllDay"),
        "location": (e.get("location") or {}).get("displayName"),
        "attendees": addresses_of(e.get("attendees")),
        "organizer": ((e.get("organizer") or {}).get("emailAddress") or {}).get("address"),
        "type": e.get("type"),
        "preview": e.get("bodyPreview"),
    }


def _explain_403(ctx: Ctx, calendar: str, exc: GraphError) -> None:
    if exc.status == 403 and calendar != "self":
        raise PolicyError(f"Kalender {calendar!r} ist für {ctx.cfg.account} nicht freigegeben") from exc


def _check_new_attendees(ctx: Ctx, calendar: str, attendees: list[str]) -> None:
    if not attendees:
        return
    if calendar != "self":
        raise PolicyError(
            "In fremden Kalendern sind keine Teilnehmer erlaubt (Einladungen gingen im Namen des Kalenderinhabers raus)"
        )
    check_attendees_internal(ctx.cfg, attendees)


def _guard_existing(ctx: Ctx, calendar: str, root: str, event_id: str) -> None:
    try:
        e = ctx.graph.get(f"{root}/events/{quote_id(event_id)}", params={"$select": "id,attendees,type,seriesMasterId"})
        attendees = addresses_of(e.get("attendees"))
        if e.get("type") in ("occurrence", "exception") and e.get("seriesMasterId"):
            master = ctx.graph.get(f"{root}/events/{quote_id(e['seriesMasterId'])}", params={"$select": "id,attendees"})
            attendees += addresses_of(master.get("attendees"))
    except GraphError as exc:
        _explain_403(ctx, calendar, exc)
        raise PolicyError(f"Termin nicht sicher prüfbar ({exc}); nichts geändert") from exc
    if calendar != "self" and attendees:
        raise PolicyError(
            "Der Termin hat Teilnehmer; Ändern oder Löschen würde Nachrichten im Namen des Kalenderinhabers verschicken"
        )
    external = [a for a in attendees if not is_internal(a, ctx.cfg.internal_domains)]
    if external:
        raise PolicyError(
            "Der Termin hat externe Teilnehmer; Ändern oder Löschen würde Nachrichten nach außen schicken"
        )


def calendar_list(ctx: Ctx, calendar: str, start: str, end: str) -> dict:
    """Termine eines Kalenders im Zeitraum (ISO 8601 ohne Zone, z. B. 2026-10-07T00:00:00)."""
    root = calendar_root(ctx.cfg, calendar)
    try:
        events = ctx.graph.get_all(
            f"{root}/calendar/calendarView",
            params={
                "startDateTime": _with_offset(ctx, start),
                "endDateTime": _with_offset(ctx, end),
                "$select": _SELECT,
                "$top": 100,
            },
            headers=_tz(ctx),
            limit=200,
        )
    except GraphError as exc:
        _explain_403(ctx, calendar, exc)
        raise
    return {"calendar": calendar, "events": [_event(e) for e in events]}


def calendar_create(
    ctx: Ctx,
    calendar: str,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    all_day: bool = False,
) -> dict:
    """Termin anlegen. Teilnehmer nur im eigenen Kalender (self) und nur interne."""
    root = calendar_root(ctx.cfg, calendar)
    attendees_ = clean_addresses(attendees, "attendees")
    _check_new_attendees(ctx, calendar, attendees_)
    payload: dict = {
        "subject": subject,
        "start": _when(ctx, start),
        "end": _when(ctx, end),
        "isAllDay": bool(all_day),
    }
    if location:
        payload["location"] = {"displayName": location}
    if body:
        payload["body"] = {"contentType": "Text", "content": body}
    if attendees_:
        payload["attendees"] = [{"emailAddress": {"address": a}, "type": "required"} for a in attendees_]
    try:
        created = ctx.graph.post(f"{root}/calendar/events", json=payload)
    except GraphError as exc:
        _explain_403(ctx, calendar, exc)
        raise
    return {"created": _event(created or {})}


def calendar_update(
    ctx: Ctx,
    calendar: str,
    event_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
) -> dict:
    """Termin ändern. Nicht angegebene Felder bleiben. Termine mit Teilnehmern in fremden Kalendern sind tabu."""
    root = calendar_root(ctx.cfg, calendar)
    for value in (start, end):
        if value is not None:
            _local(ctx, value)
    attendees_ = clean_addresses(attendees, "attendees") if attendees is not None else None
    _check_new_attendees(ctx, calendar, attendees_ or [])
    _guard_existing(ctx, calendar, root, event_id)
    patch: dict = {}
    if subject is not None:
        patch["subject"] = subject
    if start is not None:
        patch["start"] = _when(ctx, start)
    if end is not None:
        patch["end"] = _when(ctx, end)
    if location is not None:
        patch["location"] = {"displayName": location}
    if body is not None:
        patch["body"] = {"contentType": "Text", "content": body}
    if attendees_ is not None:
        patch["attendees"] = [{"emailAddress": {"address": a}, "type": "required"} for a in attendees_]
    if not patch:
        raise PolicyError("nichts zu ändern")
    try:
        updated = ctx.graph.patch(f"{root}/events/{quote_id(event_id)}", json=patch)
    except GraphError as exc:
        _explain_403(ctx, calendar, exc)
        raise
    return {"updated": _event(updated or {"id": event_id})}


def calendar_delete(ctx: Ctx, calendar: str, event_id: str) -> dict:
    """Termin löschen. Termine mit Teilnehmern in fremden Kalendern sind tabu."""
    root = calendar_root(ctx.cfg, calendar)
    _guard_existing(ctx, calendar, root, event_id)
    try:
        ctx.graph.delete(f"{root}/events/{quote_id(event_id)}")
    except GraphError as exc:
        _explain_403(ctx, calendar, exc)
        raise
    return {"deleted": True}


def calendar_free_busy(ctx: Ctx, calendars: list[str], start: str, end: str, interval_minutes: int = 30) -> dict:
    """Frei/Gebucht für self und freigegebene Kalender im Zeitraum."""
    _local(ctx, start)
    _local(ctx, end)
    keys = list(calendars)
    addresses = [calendar_address(ctx.cfg, k) for k in keys]
    data = (
        ctx.graph.post(
            "/me/calendar/getSchedule",
            json={
                "schedules": addresses,
                "startTime": _when(ctx, start),
                "endTime": _when(ctx, end),
                "availabilityViewInterval": int(interval_minutes),
            },
            headers=_tz(ctx),
        )
        or {}
    )
    by_address = {s.get("scheduleId", "").lower(): s for s in data.get("value", [])}
    schedules = []
    for key, address in zip(keys, addresses, strict=True):
        s = by_address.get(address.lower(), {})
        schedules.append(
            {
                "calendar": key,
                "availability_view": s.get("availabilityView"),
                "busy": [
                    {
                        "start": (i.get("start") or {}).get("dateTime"),
                        "end": (i.get("end") or {}).get("dateTime"),
                        "status": i.get("status"),
                    }
                    for i in s.get("scheduleItems", [])
                ],
            }
        )
    return {"schedules": schedules}


TOOLS = [calendar_list, calendar_create, calendar_update, calendar_delete, calendar_free_busy]
