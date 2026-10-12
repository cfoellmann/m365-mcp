import json
from pathlib import Path

import httpx
import pytest

from m365_mcp.config import Config


class Router:
    """Ersetzt Graph: (Methode, Pfad ohne /v1.0) → Antwort(en). Zeichnet alle Aufrufe auf."""

    def __init__(self) -> None:
        self.routes: dict[tuple[str, str], list[tuple[int, object, dict]]] = {}
        self.calls: list[tuple[str, str, object, httpx.Request]] = []

    def add(self, method, path, status=200, json_body=None, headers=None):
        self.routes.setdefault((method, path), []).append((status, json_body, headers or {}))
        return self

    def __call__(self, request: httpx.Request) -> httpx.Response:
        path = request.url.path.removeprefix("/v1.0")
        body = json.loads(request.content) if request.content else None
        self.calls.append((request.method, path, body, request))
        queue = self.routes.get((request.method, path))
        if not queue:
            return httpx.Response(599, json={"error": {"code": "unrouted", "message": f"{request.method} {path}"}})
        status, json_body, headers = queue.pop(0) if len(queue) > 1 else queue[0]
        if status == 204:
            return httpx.Response(204, headers=headers)
        return httpx.Response(status, json=json_body if json_body is not None else {}, headers=headers)

    def sent(self, method, path):
        return [c for c in self.calls if c[0] == method and c[1] == path]


@pytest.fixture
def cfg(tmp_path: Path) -> Config:
    return Config(
        path=tmp_path / "config.toml",
        tenant_id="tid",
        client_id="cid",
        account="olivia@example.org",
        internal_domains=("example.org",),
        reply_to="chef@example.org",
        external_cc=("chef@example.org",),
        calendars={"christian": "christian@example.org", "nina": "nina@example.org"},
        plans={"familie": "plan1"},
        assignees={"christian": "uid-c"},
    )


@pytest.fixture
def router() -> Router:
    return Router()


@pytest.fixture
def ctx(cfg, router):
    from m365_mcp.context import Ctx
    from m365_mcp.graph import GraphClient

    graph = GraphClient(lambda force_refresh=False: "tok", transport=httpx.MockTransport(router), sleep=lambda s: None)
    return Ctx(cfg=cfg, graph=graph)
