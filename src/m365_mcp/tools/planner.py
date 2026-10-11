"""Planner: nur Pläne aus [planner.plans]. Auch Aufgaben-IDs werden gegen den Plan geprüft."""

from collections.abc import Iterator
from contextlib import contextmanager

from m365_mcp.context import Ctx
from m365_mcp.errors import GraphError, PolicyError
from m365_mcp.graph import quote_id
from m365_mcp.policy import assignee_id, config_key, plan_id
from m365_mcp.tools.todo import parse_due


@contextmanager
def _explain_403(ctx: Ctx) -> Iterator[None]:
    """Häufigste Ursache: Konto ist nur Besitzer der Gruppe des Plans – Planner verlangt Mitgliedschaft."""
    try:
        yield
    except GraphError as exc:
        if exc.status == 403:
            raise PolicyError(
                f"{exc}: ist {ctx.cfg.account} Mitglied der Gruppe des Plans? Besitz allein reicht nicht"
            ) from exc
        raise


def _assignee_keys(ctx: Ctx, assignments: dict) -> list[str]:
    reverse = {v: k for k, v in ctx.cfg.assignees.items()}
    return [reverse.get(uid, uid) for uid in assignments or {}]


def _task(ctx: Ctx, t: dict) -> dict:
    return {
        "id": t.get("id"),
        "title": t.get("title"),
        "bucket_id": t.get("bucketId"),
        "percent_complete": t.get("percentComplete"),
        "due": (t.get("dueDateTime") or "")[:10] or None,
        "assignees": _assignee_keys(ctx, t.get("assignments")),
    }


def _task_in_plan(ctx: Ctx, plan: str, task_id: str) -> dict:
    pid = plan_id(ctx.cfg, plan)
    t = ctx.graph.get(f"/planner/tasks/{quote_id(task_id)}")
    if t.get("planId") != pid:
        raise PolicyError(f"Aufgabe {task_id} ist nicht im Plan {plan!r}")
    return t


def _resolve_assignees(ctx: Ctx, keys: list[str]) -> dict:
    ids = []
    for key in keys:
        if config_key(key) == "self":
            ids.append(ctx.graph.get("/me", params={"$select": "id"})["id"])
        else:
            ids.append(assignee_id(ctx.cfg, key))
    return {uid: {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"} for uid in ids}


def planner_tasks(ctx: Ctx, plan: str, include_completed: bool = False) -> dict:
    """Aufgaben eines freigegebenen Plans; standardmäßig nur offene.
    truncated=true: der Plan hat mehr als 200 Aufgaben (auch erledigte), nur die ersten wurden geprüft."""
    with _explain_403(ctx):
        tasks, truncated = ctx.graph.get_all(f"/planner/plans/{quote_id(plan_id(ctx.cfg, plan))}/tasks")
    if not include_completed:
        tasks = [t for t in tasks if (t.get("percentComplete") or 0) < 100]
    return {"tasks": [_task(ctx, t) for t in tasks], "truncated": truncated}


def planner_buckets(ctx: Ctx, plan: str) -> dict:
    """Buckets (Spalten) eines freigegebenen Plans. truncated=true: mehr als 200 Buckets, nur die ersten geliefert."""
    with _explain_403(ctx):
        buckets, truncated = ctx.graph.get_all(f"/planner/plans/{quote_id(plan_id(ctx.cfg, plan))}/buckets")
    return {"buckets": [{"id": b["id"], "name": b.get("name")} for b in buckets], "truncated": truncated}


def planner_create(
    ctx: Ctx,
    plan: str,
    title: str,
    bucket_id: str | None = None,
    due: str | None = None,
    assignees: list[str] | None = None,
) -> dict:
    """Aufgabe anlegen. assignees: Schlüssel aus [planner.assignees] oder "self"."""
    payload: dict = {"planId": plan_id(ctx.cfg, plan), "title": title}
    if bucket_id:
        payload["bucketId"] = bucket_id
    if due is not None:
        payload["dueDateTime"] = f"{parse_due(due)}T00:00:00Z"
    if assignees:
        payload["assignments"] = _resolve_assignees(ctx, assignees)
    with _explain_403(ctx):
        created = ctx.graph.post("/planner/tasks", json=payload)
    return {"created": _task(ctx, created or {})}


def planner_update(
    ctx: Ctx,
    plan: str,
    task_id: str,
    title: str | None = None,
    due: str | None = None,
    bucket_id: str | None = None,
) -> dict:
    """Aufgabe ändern; nicht angegebene Felder bleiben."""
    payload: dict = {}
    if title is not None:
        payload["title"] = title
    if due is not None:
        payload["dueDateTime"] = f"{parse_due(due)}T00:00:00Z"
    if bucket_id is not None:
        payload["bucketId"] = bucket_id
    if not payload:
        raise PolicyError("nichts zu ändern")
    with _explain_403(ctx):
        t = _task_in_plan(ctx, plan, task_id)
        ctx.graph.patch(f"/planner/tasks/{quote_id(task_id)}", json=payload, headers={"If-Match": t["@odata.etag"]})
    return {"updated": True}


def planner_complete(ctx: Ctx, plan: str, task_id: str) -> dict:
    """Aufgabe als erledigt markieren (100 %)."""
    with _explain_403(ctx):
        t = _task_in_plan(ctx, plan, task_id)
        ctx.graph.patch(
            f"/planner/tasks/{quote_id(task_id)}", json={"percentComplete": 100}, headers={"If-Match": t["@odata.etag"]}
        )
    return {"completed": True}


TOOLS = [planner_tasks, planner_buckets, planner_create, planner_update, planner_complete]
