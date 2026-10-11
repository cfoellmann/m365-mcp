"""Dünner Graph-Client: Bearer-Token, Retry bei 429/503/504, Fehlerzuordnung."""

import time
from collections.abc import Callable
from urllib.parse import quote

import httpx

from m365_mcp.errors import GraphError, ReauthRequired

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_RETRY = {429, 503, 504}
_MAX_WAIT = 30.0


def quote_id(value: str) -> str:
    return quote(value, safe="")


def _retry_after(response: httpx.Response) -> float:
    try:
        return min(float(response.headers.get("Retry-After", "2")), _MAX_WAIT)
    except ValueError:
        return 2.0


def _error(response: httpx.Response) -> tuple[str, str]:
    try:
        err = response.json()["error"]
        return err.get("code", "unknown"), err.get("message", "")
    except (ValueError, KeyError, TypeError):
        return "unknown", response.text[:200]


class GraphClient:
    def __init__(
        self,
        token: Callable[[], str],
        *,
        transport: httpx.BaseTransport | None = None,
        sleep: Callable[[float], None] = time.sleep,
        max_retries: int = 3,
    ) -> None:
        self._token = token
        self._sleep = sleep
        self._max_retries = max_retries
        self._http = httpx.Client(base_url=GRAPH_BASE, transport=transport, timeout=30.0)

    def request(self, method, path, *, params=None, json=None, headers=None):
        # POST (sendMail, send, Anlegen) ist nicht idempotent: 503/504 können heißen, dass Graph den
        # Aufruf trotzdem ausgeführt hat. Nur 429 sagt sicher „nicht verarbeitet“.
        retry_on = _RETRY if method != "POST" else {429}
        for attempt in range(self._max_retries + 1):
            merged = {"Authorization": f"Bearer {self._token()}", **(headers or {})}
            try:
                response = self._http.request(method, path, params=params, json=json, headers=merged)
            except httpx.TransportError as exc:
                raise GraphError(0, "transport", type(exc).__name__) from exc
            if response.status_code in retry_on and attempt < self._max_retries:
                self._sleep(_retry_after(response))
                continue
            break
        if response.status_code == 401:
            raise ReauthRequired(_error(response)[0])
        if response.status_code >= 400:
            code, message = _error(response)
            raise GraphError(response.status_code, code, message)
        if response.status_code == 204 or not response.content:
            return None
        return response.json()

    def get(self, path, params=None, headers=None) -> dict:
        return self.request("GET", path, params=params, headers=headers) or {}

    def get_all(self, path, params=None, headers=None, limit: int = 200) -> tuple[list[dict], bool]:
        """Folgt @odata.nextLink bis `limit`. Zweiter Wert: True, wenn es mehr Einträge gab als geliefert."""
        items: list[dict] = []
        page = self.get(path, params=params, headers=headers)
        while True:
            items.extend(page.get("value", []))
            next_link = page.get("@odata.nextLink")
            if len(items) >= limit or not next_link:
                return items[:limit], len(items) > limit or bool(next_link)
            page = self.get(next_link, headers=headers)

    def post(self, path, json=None, headers=None):
        return self.request("POST", path, json=json, headers=headers)

    def patch(self, path, json=None, headers=None):
        return self.request("PATCH", path, json=json, headers=headers)

    def delete(self, path, headers=None) -> None:
        self.request("DELETE", path, headers=headers)
