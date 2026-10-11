import pytest

from m365_mcp.errors import GraphError, PolicyError
from m365_mcp.tools.calendar import (
    calendar_create,
    calendar_delete,
    calendar_free_busy,
    calendar_list,
    calendar_update,
)

NINA = "/users/nina@example.org"


def att(address):
    return {"emailAddress": {"address": address}, "type": "required"}


def event(eid, attendees=(), etype="singleInstance", master=None):
    return {
        "id": eid,
        "subject": "Termin",
        "start": {"dateTime": "2026-10-07T09:00:00"},
        "end": {"dateTime": "2026-10-07T10:00:00"},
        "attendees": [att(a) for a in attendees],
        "type": etype,
        "seriesMasterId": master,
    }


def test_list_foreign_calendar(ctx, router):
    router.add("GET", f"{NINA}/calendar/calendarView", json_body={"value": [event("e1")]})
    out = calendar_list(ctx, "nina", "2026-10-07T00:00:00", "2026-10-08T00:00:00")
    assert out["events"][0]["id"] == "e1"
    assert router.calls[0][3].headers["Prefer"] == 'outlook.timezone="Europe/Berlin"'
    params = router.calls[0][3].url.params
    assert params["startDateTime"] == "2026-10-07T00:00:00+02:00"
    assert params["endDateTime"] == "2026-10-08T00:00:00+02:00"


def test_list_rejects_invalid_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_list(ctx, "self", "morgen", "2026-10-08T00:00:00")
    assert router.calls == []


def test_unknown_calendar_rejected_without_call(ctx, router):
    with pytest.raises(PolicyError):
        calendar_list(ctx, "fremd@example.org", "2026-10-07T09:00:00", "2026-10-07T10:00:00")
    assert router.calls == []


def test_403_on_foreign_calendar_is_explained(ctx, router):
    router.add(
        "GET",
        f"{NINA}/calendar/calendarView",
        status=403,
        json_body={"error": {"code": "ErrorAccessDenied", "message": "x"}},
    )
    with pytest.raises(PolicyError, match="nicht freigegeben"):
        calendar_list(ctx, "nina", "2026-10-07T00:00:00", "2026-10-08T00:00:00")


def test_create_foreign_without_attendees(ctx, router):
    router.add("POST", f"{NINA}/calendar/events", json_body=event("e1"))
    out = calendar_create(ctx, "nina", "Zahnarzt", "2026-10-07T09:00:00", "2026-10-07T10:00:00", location="Praxis")
    payload = router.calls[0][2]
    assert payload["start"] == {"dateTime": "2026-10-07T09:00:00", "timeZone": "Europe/Berlin"}
    assert payload["location"] == {"displayName": "Praxis"}
    assert "attendees" not in payload
    assert out["created"]["id"] == "e1"


def test_create_foreign_with_attendees_rejected(ctx, router):
    with pytest.raises(PolicyError, match="fremden Kalendern"):
        calendar_create(
            ctx, "nina", "X", "2026-10-07T09:00:00", "2026-10-07T10:00:00", attendees=["christian@example.org"]
        )
    assert router.calls == []


def test_create_self_with_internal_attendee(ctx, router):
    router.add("POST", "/me/calendar/events", json_body=event("e1", ["nina@example.org"]))
    calendar_create(ctx, "self", "X", "2026-10-07T09:00:00", "2026-10-07T10:00:00", attendees=["nina@example.org"])
    assert router.calls[0][2]["attendees"] == [att("nina@example.org")]


def test_create_self_with_external_attendee_rejected(ctx, router):
    with pytest.raises(PolicyError, match="Externe Teilnehmer"):
        calendar_create(ctx, "self", "X", "2026-10-07T09:00:00", "2026-10-07T10:00:00", attendees=["x@extern.example"])
    assert router.calls == []


def test_update_foreign_without_attendees(ctx, router):
    router.add("GET", f"{NINA}/events/e1", json_body=event("e1"))
    router.add("PATCH", f"{NINA}/events/e1", json_body=event("e1"))
    calendar_update(ctx, "nina", "e1", subject="Neu")
    assert router.sent("PATCH", f"{NINA}/events/e1")[0][2] == {"subject": "Neu"}


def test_update_foreign_meeting_rejected(ctx, router):
    router.add("GET", f"{NINA}/events/e1", json_body=event("e1", ["christian@example.org"]))
    with pytest.raises(PolicyError, match="Teilnehmer"):
        calendar_update(ctx, "nina", "e1", subject="Neu")
    assert not router.sent("PATCH", f"{NINA}/events/e1")


def test_delete_foreign_occurrence_of_meeting_series_rejected(ctx, router):
    router.add("GET", f"{NINA}/events/o1", json_body=event("o1", etype="occurrence", master="m1"))
    router.add("GET", f"{NINA}/events/m1", json_body=event("m1", ["christian@example.org"], etype="seriesMaster"))
    with pytest.raises(PolicyError, match="Teilnehmer"):
        calendar_delete(ctx, "nina", "o1")
    assert not router.sent("DELETE", f"{NINA}/events/o1")


def test_delete_when_check_impossible_rejected(ctx, router):
    router.add("GET", f"{NINA}/events/e1", status=500, json_body={"error": {"code": "x", "message": "y"}})
    with pytest.raises(PolicyError, match="nicht sicher"):
        calendar_delete(ctx, "nina", "e1")
    assert not router.sent("DELETE", f"{NINA}/events/e1")


def test_delete_self_with_external_attendee_rejected(ctx, router):
    router.add("GET", "/me/events/e1", json_body=event("e1", ["x@extern.example"]))
    with pytest.raises(PolicyError, match="externe Teilnehmer"):
        calendar_delete(ctx, "self", "e1")


def test_update_self_adding_external_rejected(ctx, router):
    with pytest.raises(PolicyError, match="Externe Teilnehmer"):
        calendar_update(ctx, "self", "e1", attendees=["x@extern.example"])
    assert router.calls == []


def test_403_on_foreign_patch_is_explained(ctx, router):
    router.add("GET", f"{NINA}/events/e1", json_body=event("e1"))
    router.add(
        "PATCH", f"{NINA}/events/e1", status=403, json_body={"error": {"code": "ErrorAccessDenied", "message": "x"}}
    )
    with pytest.raises(PolicyError, match="nicht freigegeben"):
        calendar_update(ctx, "nina", "e1", subject="Neu")


def test_delete_self_internal_meeting(ctx, router):
    router.add("GET", "/me/events/e1", json_body=event("e1", ["nina@example.org"]))
    router.add("DELETE", "/me/events/e1", status=204)
    assert calendar_delete(ctx, "self", "e1") == {"deleted": True}


def test_update_without_fields_rejected(ctx, router):
    router.add("GET", "/me/events/e1", json_body=event("e1"))
    with pytest.raises(PolicyError, match="nichts"):
        calendar_update(ctx, "self", "e1")


def test_free_busy_maps_keys(ctx, router):
    router.add(
        "POST",
        "/me/calendar/getSchedule",
        json_body={
            "value": [
                {"scheduleId": "olivia@example.org", "availabilityView": "00", "scheduleItems": []},
                {
                    "scheduleId": "nina@example.org",
                    "availabilityView": "02",
                    "scheduleItems": [{"status": "busy", "start": {"dateTime": "s"}, "end": {"dateTime": "e"}}],
                },
            ]
        },
    )
    out = calendar_free_busy(ctx, ["self", "nina"], "2026-10-07T09:00:00", "2026-10-07T10:00:00")
    assert router.calls[0][2]["schedules"] == ["olivia@example.org", "nina@example.org"]
    assert [s["calendar"] for s in out["schedules"]] == ["self", "nina"]
    assert out["schedules"][1]["busy"] == [{"start": "s", "end": "e", "status": "busy"}]


def test_free_busy_unknown_key_rejected(ctx, router):
    with pytest.raises(PolicyError):
        calendar_free_busy(ctx, ["boss"], "2026-10-07T09:00:00", "2026-10-07T10:00:00")


def test_list_rejects_zoned_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_list(ctx, "self", "2026-10-07T09:00:00+02:00", "2026-10-08T00:00:00")
    assert router.calls == []


def test_create_rejects_zoned_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_create(ctx, "self", "X", "2026-10-07T09:00:00+02:00", "2026-10-07T10:00:00")
    assert router.calls == []


def test_update_rejects_zoned_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_update(ctx, "self", "e1", start="2026-10-07T09:00:00+02:00")
    assert router.calls == []


def test_free_busy_rejects_zoned_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_free_busy(ctx, ["self"], "2026-10-07T09:00:00+02:00", "2026-10-07T10:00:00")
    assert router.calls == []


def test_create_rejects_invalid_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_create(ctx, "self", "X", "morgen", "2026-10-07T10:00:00")
    assert router.calls == []


def test_free_busy_rejects_invalid_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_free_busy(ctx, ["self"], "morgen", "2026-10-07T10:00:00")
    assert router.calls == []


def test_create_all_day_date_only_is_normalized(ctx, router):
    router.add("POST", "/me/calendar/events", json_body=event("e1"))
    calendar_create(ctx, "self", "Urlaub", "2026-10-07", "2026-10-08", all_day=True)
    payload = router.calls[0][2]
    assert payload["start"]["dateTime"] == "2026-10-07T00:00:00"
    assert payload["end"]["dateTime"] == "2026-10-08T00:00:00"


def test_update_rejects_invalid_time(ctx, router):
    with pytest.raises(PolicyError, match="ISO 8601"):
        calendar_update(ctx, "self", "e1", start="morgen")
    assert router.calls == []


def test_free_busy_403_on_foreign_calendar_is_explained(ctx, router):
    router.add(
        "POST",
        "/me/calendar/getSchedule",
        status=403,
        json_body={"error": {"code": "ErrorAccessDenied", "message": "x"}},
    )
    with pytest.raises(PolicyError, match="nicht freigegeben"):
        calendar_free_busy(ctx, ["self", "nina"], "2026-10-07T09:00:00", "2026-10-07T10:00:00")


def test_free_busy_403_on_self_is_reraised(ctx, router):
    router.add(
        "POST",
        "/me/calendar/getSchedule",
        status=403,
        json_body={"error": {"code": "ErrorAccessDenied", "message": "x"}},
    )
    with pytest.raises(GraphError):
        calendar_free_busy(ctx, ["self"], "2026-10-07T09:00:00", "2026-10-07T10:00:00")


def test_list_reports_truncation(ctx, router):
    router.add(
        "GET",
        f"{NINA}/calendar/calendarView",
        json_body={
            "value": [event(f"e{i}") for i in range(200)],
            "@odata.nextLink": f"https://graph.microsoft.com/v1.0{NINA}/calendar/calendarView?$skip=200",
        },
    )
    out = calendar_list(ctx, "nina", "2026-10-07T00:00:00", "2026-10-08T00:00:00")
    assert len(out["events"]) == 200
    assert out["truncated"] is True
    assert len(router.calls) == 1


def test_list_not_truncated(ctx, router):
    router.add("GET", f"{NINA}/calendar/calendarView", json_body={"value": [event("e1")]})
    out = calendar_list(ctx, "nina", "2026-10-07T00:00:00", "2026-10-08T00:00:00")
    assert out["truncated"] is False
