"""Microsoft To Do: nur eigene bzw. für den Agenten sichtbare Listen (/me/todo)."""

import datetime
import re

from m365_mcp.context import Ctx
from m365_mcp.errors import PolicyError
from m365_mcp.graph import quote_id

_ISO_DATE = re.compile(r"\d{4}-\d{2}-\d{2}")


def parse_due(due: str) -> str:
    # date.fromisoformat akzeptiert auch 20261008; erlaubt ist nur die Form YYYY-MM-DD.
    try:
        if not _ISO_DATE.fullmatch(due):
            raise ValueError(due)
        return datetime.date.fromisoformat(due).isoformat()
    except ValueError as exc:
        raise PolicyError(f"Fälligkeitsdatum {due!r} ungültig, erwartet YYYY-MM-DD") from exc


def _task(t: dict) -> dict:
    due = (t.get("dueDateTime") or {}).get("dateTime")
    return {
        "id": t.get("id"),
        "title": t.get("title"),
        "status": t.get("status"),
        "due": due[:10] if due else None,
        "importance": t.get("importance"),
    }


def _tasks_path(list_id: str) -> str:
    return f"/me/todo/lists/{quote_id(list_id)}/tasks"


def _tz(ctx: Ctx) -> dict:
    # Ohne diesen Header liefert Graph dueDateTime in UTC (Mitternacht Berlin = Vortag 22:00).
    return {"Prefer": f'outlook.timezone="{ctx.cfg.timezone}"'}


def todo_lists(ctx: Ctx) -> dict:
    """Alle eigenen To-Do-Listen. truncated=true: mehr als 200 Listen, nur die ersten geliefert."""
    lists, truncated = ctx.graph.get_all("/me/todo/lists")
    return {
        "lists": [
            {
                "id": lst["id"],
                "name": lst.get("displayName"),
                "is_owner": lst.get("isOwner"),
                "is_shared": lst.get("isShared"),
            }
            for lst in lists
        ],
        "truncated": truncated,
    }


def todo_list_tasks(ctx: Ctx, list_id: str, include_completed: bool = False) -> dict:
    """Aufgaben einer Liste; standardmäßig nur offene.
    truncated=true: mehr als 200 Aufgaben, nur die ersten geliefert."""
    params = {} if include_completed else {"$filter": "status ne 'completed'"}
    tasks, truncated = ctx.graph.get_all(_tasks_path(list_id), params=params, headers=_tz(ctx))
    return {"tasks": [_task(t) for t in tasks], "truncated": truncated}


def _fields(ctx: Ctx, title: str | None, due: str | None, note: str | None) -> dict:
    payload: dict = {}
    if title is not None:
        payload["title"] = title
    if due is not None:
        payload["dueDateTime"] = {"dateTime": f"{parse_due(due)}T00:00:00", "timeZone": ctx.cfg.timezone}
    if note is not None:
        payload["body"] = {"contentType": "text", "content": note}
    return payload


def todo_create(ctx: Ctx, list_id: str, title: str, due: str | None = None, note: str | None = None) -> dict:
    """Aufgabe anlegen (due als YYYY-MM-DD)."""
    created = ctx.graph.post(_tasks_path(list_id), json=_fields(ctx, title, due, note), headers=_tz(ctx))
    return {"created": _task(created or {})}


def todo_update(
    ctx: Ctx, list_id: str, task_id: str, title: str | None = None, due: str | None = None, note: str | None = None
) -> dict:
    """Aufgabe ändern; nicht angegebene Felder bleiben."""
    payload = _fields(ctx, title, due, note)
    if not payload:
        raise PolicyError("nichts zu ändern")
    updated = ctx.graph.patch(f"{_tasks_path(list_id)}/{quote_id(task_id)}", json=payload, headers=_tz(ctx))
    return {"updated": _task(updated or {"id": task_id})}


def todo_complete(ctx: Ctx, list_id: str, task_id: str) -> dict:
    """Aufgabe als erledigt markieren."""
    ctx.graph.patch(f"{_tasks_path(list_id)}/{quote_id(task_id)}", json={"status": "completed"})
    return {"completed": True}


TOOLS = [todo_lists, todo_list_tasks, todo_create, todo_update, todo_complete]
