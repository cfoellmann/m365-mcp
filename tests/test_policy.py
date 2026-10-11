import pytest

from m365_mcp.errors import PolicyError
from m365_mcp.policy import (
    addresses_of,
    apply_external_mail_policy,
    assignee_id,
    calendar_address,
    calendar_root,
    check_attendees_internal,
    clean_addresses,
    is_internal,
    plan_id,
    recipients,
)


@pytest.mark.parametrize(
    "address,expected",
    [
        ("chef@example.org", True),
        ("Chef@Example.ORG", True),
        ("  chef@example.org ", True),
        ("a@sub.example.org", False),
        ("a@example.org.evil.example", False),
        ("a@extern.example", False),
        ("kein-at", False),
        ("a@b@example.org", False),
    ],
)
def test_is_internal(address, expected):
    assert is_internal(address, ("example.org",)) is expected


def test_clean_addresses():
    assert clean_addresses([" a@example.org ", ""], "to") == ["a@example.org"]
    assert clean_addresses(None, "cc") == []
    with pytest.raises(PolicyError, match="to"):
        clean_addresses(["kaputt"], "to")


def test_recipient_roundtrip():
    objs = recipients(["a@example.org"])
    assert objs == [{"emailAddress": {"address": "a@example.org"}}]
    assert addresses_of(objs) == ["a@example.org"]
    assert addresses_of(None) == []


def test_internal_only_mail_unchanged(cfg):
    cc, reply_to, applied = apply_external_mail_policy(cfg, ["nina@example.org"], [], [])
    assert (cc, reply_to, applied) == ([], None, [])


@pytest.mark.parametrize(
    "to,cc,bcc",
    [
        (["x@extern.example"], [], []),
        (["nina@example.org"], ["x@extern.example"], []),
        (["nina@example.org"], [], ["x@extern.example"]),
        (["a@sub.example.org"], [], []),
    ],
)
def test_external_anywhere_applies_policy(cfg, to, cc, bcc):
    new_cc, reply_to, applied = apply_external_mail_policy(cfg, to, cc, bcc)
    assert reply_to == ["chef@example.org"]
    assert "chef@example.org" in new_cc
    assert applied == ["reply_to", "cc"]


def test_external_policy_does_not_duplicate_cc(cfg):
    new_cc, _, _ = apply_external_mail_policy(cfg, ["x@extern.example"], ["CHEF@example.org"], [])
    assert len(new_cc) == 1


def test_attendees_internal(cfg):
    check_attendees_internal(cfg, ["nina@example.org"])
    with pytest.raises(PolicyError, match="x@extern.example"):
        check_attendees_internal(cfg, ["nina@example.org", "x@extern.example"])


def test_calendar_resolution(cfg):
    assert calendar_root(cfg, "self") == "/me"
    assert calendar_root(cfg, "nina") == "/users/nina@example.org"
    assert calendar_address(cfg, "self") == "olivia@example.org"
    assert calendar_address(cfg, "christian") == "christian@example.org"
    with pytest.raises(PolicyError, match="fremd@example.org"):
        calendar_root(cfg, "fremd@example.org")


def test_plan_and_assignee_resolution(cfg):
    assert plan_id(cfg, "familie") == "plan1"
    assert assignee_id(cfg, "christian") == "uid-c"
    with pytest.raises(PolicyError):
        plan_id(cfg, "andere")
    with pytest.raises(PolicyError):
        assignee_id(cfg, "nina")


def test_external_policy_deduplicates_within_external_cc(cfg):
    import dataclasses

    cfg2 = dataclasses.replace(cfg, external_cc=("chef@example.org", "CHEF@example.org"))
    new_cc, _, _ = apply_external_mail_policy(cfg2, ["x@extern.example"], [], [])
    assert [a.lower() for a in new_cc] == ["chef@example.org"]


def test_keys_ignore_case_and_whitespace(cfg):
    assert plan_id(cfg, " Familie ") == "plan1"
    assert assignee_id(cfg, "CHRISTIAN") == "uid-c"
    assert calendar_address(cfg, "Christian") == "christian@example.org"
    assert calendar_root(cfg, " SELF ") == "/me"
    with pytest.raises(PolicyError, match="'Andere'"):
        plan_id(cfg, "Andere")
