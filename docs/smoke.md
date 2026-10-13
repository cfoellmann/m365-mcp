# Smoke-Test gegen den echten Tenant

Voraussetzung: `m365-mcp check` meldet `OK <upn>`. Jeden Schritt über den Agenten (Chat) auslösen
und das Ergebnis im Postfach/Kalender eines Menschen prüfen.

1. Mail an eine interne Adresse → kommt an, ohne Reply-To, ohne zusätzliche CC.
2. Mail an eine externe Adresse → kommt an mit Reply-To und CC laut `[external_mail]`;
   Werkzeugantwort `policy_applied: ["reply_to", "cc"]`.
3. Externe Mail an den Agenten → wird abgewiesen (NDR), sofern das Postfach nur interne Absender annimmt.
4. Antwort auf eine interne Mail → `policy_applied: []`. Antworttext mit Zeilenumbrüchen und
   `<tag>` schreiben; Umbrüche und Zeichen müssen unverändert ankommen. Beide Formate prüfen:
   einmal auf eine in Outlook geschriebene Mail (Text) und einmal auf eine HTML-Mail
   (z. B. eine Planner-Benachrichtigung) – bei HTML gingen Umbrüche und `<tag>` früher verloren.
5. Termin ohne Teilnehmer in einem freigegebenen Kalender anlegen, ändern, löschen → klappt.
   „Was steht morgen an?“ liefert genau die Termine des Tages (Zeitzone des Abfragefensters).
6. Termin mit Teilnehmern in einem freigegebenen Kalender ändern → abgelehnt.
7. Termin im eigenen Kalender mit externem Teilnehmer → abgelehnt.
8. To Do: Liste anzeigen, Aufgabe mit Fälligkeit anlegen (Datum stimmt in To Do, kein Vortag) und erledigen. Zusätzlich prüfen: Erscheint eine Liste,
   die ein Mensch mit dem Agenten geteilt hat, in `todo_lists`? (Offener Punkt der Spec.)
9. Planner: Aufgabe im freigegebenen Plan mit Fälligkeit anlegen (Datum stimmt in Planner, kein
   Vortag), zuweisen, erledigen; Aufgabe aus einem anderen
   Plan ändern → abgelehnt.
