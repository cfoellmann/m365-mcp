import asyncio
import logging

import pytest
from mcp.server.mcpserver.exceptions import ToolError

from m365_mcp.server import ALL_TOOLS, build_server

EXPECTED = {
    "mail_list",
    "mail_get",
    "mail_send",
    "mail_reply",
    "mail_mark_read",
    "mail_move",
    "calendar_list",
    "calendar_create",
    "calendar_update",
    "calendar_delete",
    "calendar_free_busy",
    "todo_lists",
    "todo_list_tasks",
    "todo_create",
    "todo_update",
    "todo_complete",
    "planner_tasks",
    "planner_buckets",
    "planner_create",
    "planner_update",
    "planner_complete",
}


def test_all_tools_registered_without_ctx(ctx):
    assert {f.__name__ for f in ALL_TOOLS} == EXPECTED
    tools = asyncio.run(build_server(ctx).list_tools())
    assert {t.name for t in tools} == EXPECTED
    for t in tools:
        assert "ctx" not in t.input_schema.get("properties", {}), t.name
        assert t.description


def test_send_schema_requires_to_subject_body(ctx):
    tools = {t.name: t for t in asyncio.run(build_server(ctx).list_tools())}
    assert set(tools["mail_send"].input_schema["required"]) == {"to", "subject", "body"}


def test_policy_error_becomes_tool_error(ctx, router):
    # MCPServer.call_tool wirft ToolError; der Protokoll-Handler macht daraus ein Ergebnis mit
    # isError=True, der Server läuft weiter. Hier wird die Ausnahme selbst geprüft.
    with pytest.raises(ToolError, match="nicht freigegeben"):
        asyncio.run(build_server(ctx).call_tool("calendar_list", {"calendar": "boss", "start": "a", "end": "b"}))
    assert router.calls == []


def test_reauth_error_names_login_command(ctx, router):
    router.add("GET", "/me/todo/lists", status=401, json_body={"error": {"code": "InvalidAuthenticationToken"}})
    with pytest.raises(ToolError, match="m365-mcp login --config"):
        asyncio.run(build_server(ctx).call_tool("todo_lists", {}))


def test_success_returns_result(ctx, router):
    router.add("GET", "/me/todo/lists", json_body={"value": []})
    asyncio.run(build_server(ctx).call_tool("todo_lists", {}))


def test_no_addresses_in_logs(ctx, router, caplog):
    from m365_mcp.cli import configure_logging

    configure_logging()
    with caplog.at_level(logging.DEBUG):
        with pytest.raises(ToolError):
            asyncio.run(
                build_server(ctx).call_tool(
                    "calendar_create",
                    {
                        "calendar": "self",
                        "subject": "x",
                        "start": "2026-10-07T09:00:00",
                        "end": "2026-10-07T10:00:00",
                        "attendees": ["x@extern.example"],
                    },
                )
            )
    shown = [r for r in caplog.records if r.levelno >= logging.getLogger(r.name).getEffectiveLevel()]
    assert not any("@" in r.getMessage() for r in shown)


def test_graph_error_text_does_not_leak_message(ctx, router):
    router.add(
        "GET",
        "/me/todo/lists",
        status=403,
        json_body={"error": {"code": "ErrorAccessDenied", "message": "Zugriff für geheim@example.org verweigert"}},
    )
    with pytest.raises(ToolError) as exc:
        asyncio.run(build_server(ctx).call_tool("todo_lists", {}))
    assert "403" in str(exc.value) and "ErrorAccessDenied" in str(exc.value)
    assert "geheim@example.org" not in str(exc.value)
