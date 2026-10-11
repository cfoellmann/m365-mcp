import pytest

from m365_mcp.errors import PolicyError
from m365_mcp.tools.todo import todo_complete, todo_create, todo_list_tasks, todo_lists, todo_update


def test_lists(ctx, router):
    router.add(
        "GET",
        "/me/todo/lists",
        json_body={"value": [{"id": "l1", "displayName": "Einkauf", "isOwner": True, "isShared": True}]},
    )
    assert todo_lists(ctx) == {"lists": [{"id": "l1", "name": "Einkauf", "is_owner": True, "is_shared": True}]}


def test_open_tasks_filter(ctx, router):
    router.add(
        "GET",
        "/me/todo/lists/l1/tasks",
        json_body={
            "value": [
                {
                    "id": "t1",
                    "title": "Milch",
                    "status": "notStarted",
                    "dueDateTime": {"dateTime": "2026-10-08T00:00:00.0000000"},
                }
            ]
        },
    )
    out = todo_list_tasks(ctx, "l1")
    assert router.calls[0][3].url.params["$filter"] == "status ne 'completed'"
    assert router.calls[0][3].headers["Prefer"] == 'outlook.timezone="Europe/Berlin"'
    assert out["tasks"][0]["due"] == "2026-10-08"


def test_create_with_due_and_note(ctx, router):
    router.add("POST", "/me/todo/lists/l1/tasks", json_body={"id": "t1", "title": "Milch", "status": "notStarted"})
    todo_create(ctx, "l1", "Milch", due="2026-10-08", note="fettarm")
    payload = router.calls[0][2]
    assert payload["dueDateTime"] == {"dateTime": "2026-10-08T00:00:00", "timeZone": "Europe/Berlin"}
    assert payload["body"] == {"contentType": "text", "content": "fettarm"}


def test_invalid_due_rejected(ctx):
    with pytest.raises(PolicyError, match="YYYY-MM-DD"):
        todo_create(ctx, "l1", "x", due="morgen")


def test_update_and_complete(ctx, router):
    router.add("PATCH", "/me/todo/lists/l1/tasks/t1", json_body={"id": "t1", "title": "Brot", "status": "completed"})
    todo_update(ctx, "l1", "t1", title="Brot")
    assert router.calls[0][2] == {"title": "Brot"}
    todo_complete(ctx, "l1", "t1")
    assert router.calls[1][2] == {"status": "completed"}


def test_update_without_fields_rejected(ctx):
    with pytest.raises(PolicyError, match="nichts"):
        todo_update(ctx, "l1", "t1")
