"""Fehlerklassen. Der Text jeder M365Error geht als Werkzeugfehler unverändert an das Modell."""


class M365Error(Exception):
    """Basis aller erwarteten Fehler."""


class ConfigError(M365Error):
    """Konfiguration fehlt, ist unlesbar oder verletzt eine Regel."""


class PolicyError(M365Error):
    """Eine im Code erzwungene Richtlinie verbietet den Aufruf."""


class ReauthRequired(M365Error):
    """Kein gültiges Token mehr; ein Mensch muss `m365-mcp login` ausführen."""

    def __init__(self, detail: str = "") -> None:
        super().__init__("M365-Neuanmeldung nötig" + (f" ({detail})" if detail else ""))


class GraphError(M365Error):
    """Graph hat mit einem Fehlerstatus geantwortet."""

    def __init__(self, status: int, code: str, message: str) -> None:
        # Der Rohtext von Graph (kann Adressen, Betreffs, URLs enthalten) bleibt im Attribut und
        # geht nie in str(exc), also nie an das Modell.
        if code == "transport":
            text = "Graph-Verbindungsfehler, Ergebnis unklar – nicht automatisch wiederholen"
        else:
            text = f"Graph {status} {code}"
        super().__init__(text)
        self.status = status
        self.code = code
        self.message = message
