# m365-mcp

Schmaler MCP-Server (stdio) für Microsoft 365: Mail, Kalender, To Do und Planner über Microsoft
Graph mit delegierter Anmeldung (Device Code). Gedacht für einen KI-Agenten mit eigenem
M365-Benutzer.

## Im Code erzwungene Richtlinien

- **Mail:** nur das eigene Postfach. Ist ein Empfänger (To, CC, BCC) außerhalb von
  `internal_domains`, setzt der Server `Reply-To` und CC aus `[external_mail]`. Antworten laufen
  über `createReply` → prüfen → PATCH → `send`, nie über `/reply`. `mail_move` nur in Posteingang,
  Archiv, Gelöschte Elemente, Junk-E-Mail und deren Unterordner.
- **Kalender:** nur `self` und die Schlüssel aus `[calendars]`. In fremden Kalendern keine
  Teilnehmer und kein Ändern/Löschen von Terminen mit Teilnehmern (auch Serien). Externe
  Teilnehmer sind überall ausgeschlossen.
- **Planner:** nur Pläne aus `[planner.plans]`, auch bei Aufgaben-IDs geprüft; Zuweisungen nur an
  `[planner.assignees]` und `self`.
- **To Do:** nur `/me/todo`.

## Einrichtung

1. Entra-App: öffentlicher Client, „Allow public client flows“ an, delegierte Rechte `User.Read`,
   `offline_access`, `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`,
   `Calendars.ReadWrite.Shared`, `Tasks.ReadWrite`, Admin-Consent; Enterprise App mit
   „Zuweisung erforderlich“ und nur dem Agenten-Benutzer zugewiesen.
2. `config.toml` (Beispiel unten) in einen Ordner, der nur dem ausführenden Benutzer gehört
   (`0700`). Token-Cache und Lock entstehen daneben. `login`/`serve`/`check` brechen ab, wenn der
   Ordner einem anderen Benutzer gehört oder für Gruppe/andere schreibbar ist.
3. `m365-mcp login --config config.toml` (Device Code), dann `m365-mcp check --config config.toml`.
4. Im MCP-Client: `m365-mcp serve --config config.toml` als stdio-Server.

```toml
tenant_id = "00000000-0000-0000-0000-000000000000"
client_id = "00000000-0000-0000-0000-000000000000"
account = "agent@example.org"
internal_domains = ["example.org"]
timezone = "Europe/Berlin"
max_body_chars = 8000

[external_mail]
reply_to = "mensch@example.org"
cc = ["mensch@example.org"]

[calendars]
mensch = "mensch@example.org"

[planner.plans]
familie = "<plan-id>"

[planner.assignees]
mensch = "<entra-object-id>"
```

## Update-Hinweise

**0.1.4:** `login`, `serve` und `check` brechen jetzt ab, wenn der Ordner der `config.toml` für
Gruppe oder andere schreibbar ist (vorher wurde nur der Eigentümer geprüft). Vor dem Update auf dem
Server prüfen und ggf. korrigieren:

```sh
ls -ld /pfad/zum/config-ordner   # erwartet drwx------
chmod 0700 /pfad/zum/config-ordner
```

Außerdem müssen `tenant_id` und `client_id` eine GUID oder ein Domänenname sein. Planner-Fälligkeiten
werden jetzt als Mitternacht in `timezone` gespeichert (in Graph z. B. `2026-10-09T22:00:00Z` für
den 10.10. in Europe/Berlin); angezeigt wird weiterhin das richtige Datum.

## Exit-Codes

`0` ok, `1` Fehler, `2` Aufruffehler (argparse), `3` Neuanmeldung nötig (`m365-mcp login`).

## Installation ohne uv

`requirements.lock` enthält alle Abhängigkeiten mit Hashes (ohne das Paket selbst):
`python3 -m venv venv && venv/bin/pip install --require-hashes -r requirements.lock`, danach entweder
`venv/bin/pip install --no-deps .` (legt den Befehl `m365-mcp` an) oder ohne Installation
`venv/bin/python -m m365_mcp` mit `src/` im Pfad (so macht es hermes-home über eine `.pth`-Datei).
