"""Dünner Graph-Client: Bearer-Token, Retry bei 429/503/504, Fehlerzuordnung."""

import math
import time
from collections.abc import Callable
from urllib.parse import quote

import httpx

from m365_mcp.errors import GraphError, ReauthRequired

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_RETRY = {429, 503, 504}
_MAX_WAIT = 30.0
_TIMEOUT = 30.0
# Gesamtzeit pro Aufruf: viele MCP-Clients brechen nach ~60 s ab, das Ergebnis ginge dann verloren.
_BUDGET = 50.0


def quote_id(value: str) -> str:
    return quote(value, safe="")


def _retry_after(response: httpx.Response) -> float:
    try:
        wait = float(response.headers.get("Retry-After", "2"))
    except ValueError:
        return 2.0
    if not math.isfinite(wait):
        return 2.0
    return min(max(wait, 0.0), _MAX_WAIT)


def _error(response: httpx.Response) -> tuple[str, str]:
    try:
        err = response.json()["error"]
        return err.get("code", "unknown"), err.get("message", "")
    except (ValueError, KeyError, TypeError, AttributeError):
        return "unknown", response.text[:200]


class GraphClient:
    def __init__(
        self,
        token: Callable[..., str],
        *,
        transport: httpx.BaseTransport | None = None,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
        max_retries: int = 3,
    ) -> None:
        """`token(force_refresh=False)` liefert ein Access-Token; mit True wird es neu geholt."""
        self._token = token
        self._sleep = sleep
        self._clock = clock
        self._max_retries = max_retries
        self._http = httpx.Client(base_url=GRAPH_BASE, transport=transport, timeout=_TIMEOUT)

    def request(self, method, path, *, params=None, json=None, headers=None):
        # POST (sendMail, send, Anlegen) ist nicht idempotent: 503/504 können heißen, dass Graph den
        # Aufruf trotzdem ausgeführt hat. Nur 429 sagt sicher „nicht verarbeitet“.
        retry_on = _RETRY if method != "POST" else {429}
        start = self._clock()
        attempt = 0
        force_refresh = refreshed = server_retry = False
        while True:
            remaining = _BUDGET - (self._clock() - start)
            merged = {"Authorization": f"Bearer {self._token(force_refresh=force_refresh)}", **(headers or {})}
            force_refresh = False
            try:
                response = self._http.request(
                    method, path, params=params, json=json, headers=merged, timeout=min(_TIMEOUT, max(remaining, 1.0))
                )
            except httpx.TransportError as exc:
                raise GraphError(0, "transport", type(exc).__name__) from exc
            # 401 heißt „Token abgelehnt“, der Aufruf wurde nicht ausgeführt: einmal mit frisch
            # geholtem Token wiederholen (auch POST), bevor eine Neuanmeldung verlangt wird.
            if response.status_code == 401 and not refreshed:
                force_refresh = refreshed = True
                continue
            if response.status_code in retry_on and attempt < self._max_retries:
                wait = _retry_after(response)
                if self._clock() - start + wait < _BUDGET:
                    self._sleep(wait)
                    attempt += 1
                    server_retry = server_retry or response.status_code != 429
                    continue
            break
        if response.status_code == 401:
            raise ReauthRequired(_error(response)[0])
        # Nach 503/504 kann das erste DELETE serverseitig gegriffen haben; 404 heißt dann „erledigt“.
        if method == "DELETE" and response.status_code == 404 and server_retry:
            return None
        if response.status_code >= 400:
            code, message = _error(response)
            raise GraphError(response.status_code, code, message)
        if response.status_code == 204 or not response.content:
            return None
        try:
            return response.json()
        except ValueError as exc:
            raise GraphError(response.status_code, "invalid_response", response.text[:200]) from exc

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
            # nextLink stammt aus der Antwort und bekommt das Bearer-Token: nur zu Graph folgen.
            if not next_link.startswith(GRAPH_BASE + "/"):
                raise GraphError(0, "foreign_next_link", next_link[:200])
            page = self.get(next_link, headers=headers)

    def post(self, path, json=None, headers=None):
        return self.request("POST", path, json=json, headers=headers)

    def patch(self, path, json=None, headers=None):
        return self.request("PATCH", path, json=json, headers=headers)

    def delete(self, path, headers=None) -> None:
        self.request("DELETE", path, headers=headers)
