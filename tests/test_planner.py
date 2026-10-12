import pytest

from m365_mcp.errors import PolicyError
from m365_mcp.tools.planner import (
    planner_buckets,
    planner_complete,
    planner_create,
    planner_tasks,
    planner_update,
)


def task(tid, plan="plan1", pct=0):
    return {
        "id": tid,
        "planId": plan,
        "title": "T",
        "bucketId": "b1",
        "percentComplete": pct,
        "dueDateTime": None,
        "assignments": {"uid-c": {}},
        "@odata.etag": 'W/"etag"',
    }


def test_unknown_plan_rejected(ctx, router):
    with pytest.raises(PolicyError):
        planner_tasks(ctx, "andere")
    assert router.calls == []


def test_tasks_hides_completed_and_maps_assignees(ctx, router):
    router.add("GET", "/planner/plans/plan1/tasks", json_body={"value": [task("t1"), task("t2", pct=100)]})
    out = planner_tasks(ctx, "familie")
    assert [t["id"] for t in out["tasks"]] == ["t1"]
    assert out["tasks"][0]["assignees"] == ["christian"]


def test_buckets(ctx, router):
    router.add("GET", "/planner/plans/plan1/buckets", json_body={"value": [{"id": "b1", "name": "Haus"}]})
    assert planner_buckets(ctx, "familie") == {"buckets": [{"id": "b1", "name": "Haus"}], "truncated": False}


def test_create_with_assignees(ctx, router):
    router.add("GET", "/me", json_body={"id": "uid-o"})
    router.add("POST", "/planner/tasks", json_body=task("t1"))
    planner_create(ctx, "familie", "Rasen", bucket_id="b1", due="2026-10-10", assignees=["christian", "self"])
    payload = router.sent("POST", "/planner/tasks")[0][2]
    assert payload["planId"] == "plan1"
    assert payload["dueDateTime"] == "2026-10-09T22:00:00Z"  # Mitternacht Europe/Berlin (MESZ)
    assert set(payload["assignments"]) == {"uid-c", "uid-o"}
    assert payload["assignments"]["uid-c"]["@odata.type"] == "#microsoft.graph.plannerAssignment"


def test_create_unknown_assignee_rejected(ctx, router):
    with pytest.raises(PolicyError):
        planner_create(ctx, "familie", "Rasen", assignees=["nina"])
    assert not router.sent("POST", "/planner/tasks")


def test_update_uses_etag(ctx, router):
    router.add("GET", "/planner/tasks/t1", json_body=task("t1"))
    router.add("PATCH", "/planner/tasks/t1", status=204)
    planner_update(ctx, "familie", "t1", title="Neu")
    req = router.sent("PATCH", "/planner/tasks/t1")[0]
    assert req[2] == {"title": "Neu"}
    assert req[3].headers["If-Match"] == 'W/"etag"'


@pytest.mark.parametrize("fn", ["update", "complete"])
def test_task_from_other_plan_rejected(ctx, router, fn):
    router.add("GET", "/planner/tasks/t9", json_body=task("t9", plan="geheim"))
    with pytest.raises(PolicyError, match="nicht im Plan"):
        if fn == "update":
            planner_update(ctx, "familie", "t9", title="x")
        else:
            planner_complete(ctx, "familie", "t9")
    assert not router.sent("PATCH", "/planner/tasks/t9")


def test_complete_sets_100(ctx, router):
    router.add("GET", "/planner/tasks/t1", json_body=task("t1"))
    router.add("PATCH", "/planner/tasks/t1", status=204)
    planner_complete(ctx, "familie", "t1")
    assert router.sent("PATCH", "/planner/tasks/t1")[0][2] == {"percentComplete": 100}


@pytest.mark.parametrize("due", ["", "20261008"])
def test_create_invalid_due_rejected(ctx, router, due):
    with pytest.raises(PolicyError, match="YYYY-MM-DD"):
        planner_create(ctx, "familie", "Rasen", due=due)
    assert not router.sent("POST", "/planner/tasks")


def test_tasks_reports_truncation(ctx, router):
    router.add(
        "GET",
        "/planner/plans/plan1/tasks",
        json_body={
            "value": [task(f"t{i}", pct=100) for i in range(200)],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/planner/plans/plan1/tasks?$skiptoken=x",
        },
    )
    out = planner_tasks(ctx, "familie")
    assert out["tasks"] == []
    assert out["truncated"] is True


def test_plan_and_assignee_keys_ignore_case(ctx, router):
    router.add("GET", "/me", json_body={"id": "uid-o"})
    router.add("POST", "/planner/tasks", json_body=task("t1"))
    planner_create(ctx, "Familie", "Rasen", assignees=["CHRISTIAN", " Self "])
    payload = router.sent("POST", "/planner/tasks")[0][2]
    assert payload["planId"] == "plan1"
    assert set(payload["assignments"]) == {"uid-c", "uid-o"}


def test_403_hints_at_group_membership(ctx, router):
    router.add(
        "GET",
        "/planner/plans/plan1/tasks",
        status=403,
        json_body={"error": {"code": "Forbidden", "message": "x"}},
    )
    with pytest.raises(PolicyError, match="Mitglied der Gruppe"):
        planner_tasks(ctx, "familie")


def test_due_uses_configured_timezone(ctx, router):
    import dataclasses

    ctx = dataclasses.replace(ctx, cfg=dataclasses.replace(ctx.cfg, timezone="America/New_York"))
    router.add("POST", "/planner/tasks", json_body=task("t9") | {"dueDateTime": "2026-12-01T05:00:00Z"})
    out = planner_create(ctx, "familie", "Rasen", due="2026-12-01")
    assert router.calls[0][2]["dueDateTime"] == "2026-12-01T05:00:00Z"
    assert out["created"]["due"] == "2026-12-01"


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("2026-10-09T22:00:00Z", "2026-10-10"),
        ("2026-10-10T00:00:00Z", "2026-10-10"),
        ("2026-10-10T10:00:00.0000000Z", "2026-10-10"),
        ("kaputt", "kaputt"),
    ],
)
def test_due_read_back_in_configured_timezone(ctx, router, raw, expected):
    router.add("GET", "/planner/plans/plan1/tasks", json_body={"value": [task("t1") | {"dueDateTime": raw}]})
    assert planner_tasks(ctx, "familie")["tasks"][0]["due"] == expected
