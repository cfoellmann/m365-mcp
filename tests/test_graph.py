import httpx
import pytest

from m365_mcp.errors import GraphError, ReauthRequired
from m365_mcp.graph import GraphClient, quote_id


def client(router, sleeps=None, token=None):
    sleeps = sleeps if sleeps is not None else []
    return GraphClient(
        token or (lambda force_refresh=False: "tok"),
        transport=httpx.MockTransport(router),
        sleep=sleeps.append,
        clock=lambda: sum(sleeps),  # Zeit vergeht nur durch Warten
    )


def test_get_sends_bearer_and_returns_json(router):
    router.add("GET", "/me", json_body={"userPrincipalName": "olivia@example.org"})
    assert client(router).get("/me")["userPrincipalName"] == "olivia@example.org"
    assert router.calls[0][3].headers["Authorization"] == "Bearer tok"


def test_retry_on_429_with_retry_after(router):
    router.add("GET", "/me", status=429, headers={"Retry-After": "7"})
    router.add("GET", "/me", status=503)
    router.add("GET", "/me", json_body={"ok": True})
    sleeps: list[float] = []
    assert client(router, sleeps).get("/me") == {"ok": True}
    assert sleeps == [7.0, 2.0]


def test_retry_gives_up_after_three(router):
    router.add("GET", "/me", status=504)
    with pytest.raises(GraphError) as exc:
        client(router).get("/me")
    assert exc.value.status == 504
    assert len(router.calls) == 4


def test_post_not_retried_on_504(router):
    router.add("POST", "/me/sendMail", status=504)
    with pytest.raises(GraphError) as exc:
        client(router).post("/me/sendMail", json={})
    assert exc.value.status == 504
    assert len(router.calls) == 1


def test_post_retried_on_429(router):
    router.add("POST", "/me/sendMail", status=429, headers={"Retry-After": "1"})
    router.add("POST", "/me/sendMail", status=202)
    client(router).post("/me/sendMail", json={})
    assert len(router.calls) == 2


def test_transport_error_is_graph_error():
    def boom(request):
        raise httpx.ConnectError("weg", request=request)

    graph = GraphClient(lambda force_refresh=False: "tok", transport=httpx.MockTransport(boom), sleep=lambda s: None)
    with pytest.raises(GraphError, match="Ergebnis unklar") as exc:
        graph.get("/me")
    assert exc.value.code == "transport"


def test_retry_after_is_capped(router):
    router.add("GET", "/me", status=429, headers={"Retry-After": "3600"})
    router.add("GET", "/me", json_body={})
    sleeps: list[float] = []
    client(router, sleeps).get("/me")
    assert sleeps == [30.0]


def test_401_means_reauth(router):
    router.add("GET", "/me", status=401, json_body={"error": {"code": "InvalidAuthenticationToken"}})
    with pytest.raises(ReauthRequired):
        client(router).get("/me")
    assert len(router.calls) == 2


def test_401_retries_once_with_forced_refresh(router):
    router.add("GET", "/me", status=401, json_body={"error": {"code": "InvalidAuthenticationToken"}})
    router.add("GET", "/me", json_body={"ok": True})
    forced: list[bool] = []

    def token(force_refresh=False):
        forced.append(force_refresh)
        return "neu" if force_refresh else "alt"

    assert client(router, token=token).get("/me") == {"ok": True}
    assert forced == [False, True]
    assert router.calls[1][3].headers["Authorization"] == "Bearer neu"


def test_401_refresh_also_for_post(router):
    router.add("POST", "/me/sendMail", status=401, json_body={"error": {"code": "InvalidAuthenticationToken"}})
    router.add("POST", "/me/sendMail", status=202)
    client(router).post("/me/sendMail", json={})
    assert len(router.calls) == 2


def test_error_body_with_string_error_is_graph_error(router):
    router.add("GET", "/me", status=400, json_body={"error": "invalid_request"})
    with pytest.raises(GraphError) as exc:
        client(router).get("/me")
    assert (exc.value.status, exc.value.code) == (400, "unknown")


@pytest.mark.parametrize(("header", "expected"), [("-5", 0.0), ("nan", 2.0), ("inf", 2.0), ("abc", 2.0)])
def test_retry_after_invalid_values(router, header, expected):
    router.add("GET", "/me", status=429, headers={"Retry-After": header})
    router.add("GET", "/me", json_body={})
    sleeps: list[float] = []
    client(router, sleeps).get("/me")
    assert sleeps == [expected]


def test_retries_stop_at_time_budget(router):
    router.add("GET", "/me", status=429, headers={"Retry-After": "30"})
    sleeps: list[float] = []
    with pytest.raises(GraphError) as exc:
        client(router, sleeps).get("/me")
    assert exc.value.status == 429
    assert sleeps == [30.0]  # ein zweites Warten sprengte das Budget
    assert len(router.calls) == 2


def test_request_timeout_shrinks_with_budget(router):
    router.add("GET", "/me", status=429, headers={"Retry-After": "30"})
    router.add("GET", "/me", json_body={})
    client(router).get("/me")
    first, second = (c[3].extensions["timeout"]["read"] for c in router.calls)
    assert first == 30.0
    assert second < first


def test_error_body_is_mapped(router):
    router.add("GET", "/me", status=403, json_body={"error": {"code": "ErrorAccessDenied", "message": "nein"}})
    with pytest.raises(GraphError) as exc:
        client(router).get("/me")
    assert (exc.value.status, exc.value.code) == (403, "ErrorAccessDenied")


def test_204_returns_none(router):
    router.add("DELETE", "/me/events/e1", status=204)
    assert client(router).delete("/me/events/e1") is None


def test_get_all_follows_next_link(router):
    router.add(
        "GET",
        "/me/todo/lists",
        json_body={
            "value": [{"id": "1"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=x",
        },
    )
    router.add("GET", "/me/todo/lists", json_body={"value": [{"id": "2"}]})
    items, truncated = client(router).get_all("/me/todo/lists")
    assert [i["id"] for i in items] == ["1", "2"]
    assert truncated is False


def test_get_all_respects_limit(router):
    router.add(
        "GET",
        "/me/todo/lists",
        json_body={
            "value": [{"id": str(i)} for i in range(5)],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?n=2",
        },
    )
    items, truncated = client(router).get_all("/me/todo/lists", limit=3)
    assert len(items) == 3
    assert truncated is True
    assert len(router.calls) == 1


def test_get_all_exactly_limit_without_next_link_is_not_truncated(router):
    router.add("GET", "/me/todo/lists", json_body={"value": [{"id": str(i)} for i in range(3)]})
    items, truncated = client(router).get_all("/me/todo/lists", limit=3)
    assert len(items) == 3
    assert truncated is False


def test_get_all_full_page_with_next_link_is_truncated(router):
    router.add(
        "GET",
        "/me/todo/lists",
        json_body={
            "value": [{"id": str(i)} for i in range(3)],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?n=2",
        },
    )
    items, truncated = client(router).get_all("/me/todo/lists", limit=3)
    assert len(items) == 3
    assert truncated is True


def test_quote_id():
    assert quote_id("AAMk/a+b=") == "AAMk%2Fa%2Bb%3D"


def test_get_all_overshoot_on_later_page_is_truncated(router):
    router.add(
        "GET",
        "/me/todo/lists",
        json_body={
            "value": [{"id": "1"}, {"id": "2"}],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/todo/lists?$skiptoken=x",
        },
    )
    router.add("GET", "/me/todo/lists", json_body={"value": [{"id": "3"}, {"id": "4"}]})
    items, truncated = client(router).get_all("/me/todo/lists", limit=3)
    assert [i["id"] for i in items] == ["1", "2", "3"]
    assert truncated is True
    assert len(router.calls) == 2


def test_get_all_rejects_foreign_next_link(router):
    router.add("GET", "/me/messages", json_body={"value": [{"id": "1"}], "@odata.nextLink": "https://evil.example/x"})
    with pytest.raises(GraphError, match="foreign_next_link"):
        client(router).get_all("/me/messages")
    assert len(router.calls) == 1


def test_delete_404_after_5xx_retry_counts_as_done(router):
    router.add("DELETE", "/me/events/e1", status=503)
    router.add("DELETE", "/me/events/e1", status=404, json_body={"error": {"code": "ErrorItemNotFound"}})
    assert client(router).delete("/me/events/e1") is None


def test_delete_404_without_retry_is_error(router):
    router.add("DELETE", "/me/events/e1", status=404, json_body={"error": {"code": "ErrorItemNotFound"}})
    with pytest.raises(GraphError) as exc:
        client(router).delete("/me/events/e1")
    assert exc.value.status == 404


def test_non_json_success_body_is_graph_error():
    transport = httpx.MockTransport(lambda request: httpx.Response(200, text="<html>Proxy</html>"))
    graph = GraphClient(lambda force_refresh=False: "tok", transport=transport, sleep=lambda s: None)
    with pytest.raises(GraphError) as exc:
        graph.get("/me")
    assert exc.value.code == "invalid_response"
