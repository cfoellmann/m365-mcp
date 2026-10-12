import pytest

from m365_mcp.errors import GraphError, PolicyError
from m365_mcp.tools.mail import (
    mail_get,
    mail_list,
    mail_mark_read,
    mail_move,
    mail_reply,
    mail_send,
)


def msg(mid, sender):
    return {
        "id": mid,
        "subject": "Betreff",
        "from": {"emailAddress": {"address": sender}},
        "receivedDateTime": "2026-10-06T08:00:00Z",
        "isRead": False,
        "bodyPreview": "Hallo",
    }


def test_mail_list_marks_source(ctx, router):
    router.add(
        "GET",
        "/me/mailFolders/inbox/messages",
        json_body={"value": [msg("1", "nina@example.org"), msg("2", "x@extern.example")]},
    )
    out = mail_list(ctx, unread_only=True, top=500)
    assert [m["source"] for m in out["messages"]] == ["internal", "external"]
    params = router.calls[0][3].url.params
    assert params["$filter"] == "isRead eq false"
    assert params["$top"] == "50"


def test_mail_source_external_if_sender_differs_from_from(ctx, router):
    spoofed = msg("1", "chef@example.org") | {"sender": {"emailAddress": {"address": "x@extern.example"}}}
    same = msg("2", "nina@example.org") | {"sender": {"emailAddress": {"address": "nina@example.org"}}}
    router.add("GET", "/me/mailFolders/inbox/messages", json_body={"value": [spoofed, same]})
    out = mail_list(ctx)
    assert [m["source"] for m in out["messages"]] == ["external", "internal"]
    assert "sender" in router.calls[0][3].url.params["$select"].split(",")


def test_mail_get_truncates_and_requests_text(ctx, router, cfg):
    body = "x" * (cfg.max_body_chars + 10)
    router.add(
        "GET",
        "/me/messages/m1",
        json_body={
            **msg("m1", "nina@example.org"),
            "body": {"content": body},
            "toRecipients": [],
            "ccRecipients": [],
        },
    )
    out = mail_get(ctx, "m1")
    assert out["body"].startswith("x" * cfg.max_body_chars)
    assert "gekürzt" in out["body"]
    assert router.calls[0][3].headers["Prefer"] == 'outlook.body-content-type="text"'


def test_send_internal_unchanged(ctx, router):
    router.add("POST", "/me/sendMail", status=204)
    out = mail_send(ctx, ["nina@example.org"], "Hi", "Text")
    sent = router.sent("POST", "/me/sendMail")[0][2]["message"]
    assert "replyTo" not in sent
    assert sent["ccRecipients"] == []
    assert out == {"sent": True, "policy_applied": []}


def test_send_external_bcc_applies_policy(ctx, router):
    router.add("POST", "/me/sendMail", status=204)
    out = mail_send(ctx, ["nina@example.org"], "Hi", "Text", bcc=["x@extern.example"])
    sent = router.sent("POST", "/me/sendMail")[0][2]["message"]
    assert sent["replyTo"] == [{"emailAddress": {"address": "chef@example.org"}}]
    assert {"emailAddress": {"address": "chef@example.org"}} in sent["ccRecipients"]
    assert out["policy_applied"] == ["reply_to", "cc"]


def test_send_needs_recipient(ctx):
    with pytest.raises(PolicyError, match="Empfänger"):
        mail_send(ctx, [], "Hi", "Text")


def test_reply_to_external_sets_policy_before_send(ctx, router):
    router.add(
        "POST",
        "/me/messages/m1/createReply",
        json_body={
            "id": "d1",
            "toRecipients": [{"emailAddress": {"address": "x@extern.example"}}],
            "ccRecipients": [],
        },
    )
    router.add("PATCH", "/me/messages/d1", json_body={"id": "d1"})
    router.add("POST", "/me/messages/d1/send", status=202)
    out = mail_reply(ctx, "m1", "Danke")
    order = [(c[0], c[1]) for c in router.calls]
    assert order == [
        ("POST", "/me/messages/m1/createReply"),
        ("PATCH", "/me/messages/d1"),
        ("POST", "/me/messages/d1/send"),
    ]
    assert router.calls[0][2] == {"comment": "Danke"}
    patch = router.calls[1][2]
    assert patch["replyTo"] == [{"emailAddress": {"address": "chef@example.org"}}]
    assert {"emailAddress": {"address": "chef@example.org"}} in patch["ccRecipients"]
    assert out["policy_applied"] == ["reply_to", "cc"]
    assert not router.sent("POST", "/me/messages/m1/reply")


def test_reply_all_internal_no_patch(ctx, router):
    router.add(
        "POST",
        "/me/messages/m1/createReplyAll",
        json_body={
            "id": "d1",
            "toRecipients": [{"emailAddress": {"address": "nina@example.org"}}],
        },
    )
    router.add("POST", "/me/messages/d1/send", status=202)
    out = mail_reply(ctx, "m1", "Ok", reply_all=True)
    assert not router.sent("PATCH", "/me/messages/d1")
    assert out["policy_applied"] == []


def test_reply_failure_deletes_draft(ctx, router):
    router.add(
        "POST",
        "/me/messages/m1/createReply",
        json_body={
            "id": "d1",
            "toRecipients": [{"emailAddress": {"address": "nina@example.org"}}],
        },
    )
    router.add(
        "POST",
        "/me/messages/d1/send",
        status=500,
        json_body={"error": {"code": "x", "message": "kaputt"}},
    )
    router.add("DELETE", "/me/messages/d1", status=204)
    with pytest.raises(GraphError):
        mail_reply(ctx, "m1", "Ok")
    assert router.sent("DELETE", "/me/messages/d1")


def test_mark_read(ctx, router):
    router.add("PATCH", "/me/messages/m1", json_body={"id": "m1"})
    assert mail_mark_read(ctx, "m1") == {"ok": True}
    assert router.calls[0][2] == {"isRead": True}


def test_move_to_wellknown(ctx, router):
    router.add("POST", "/me/messages/m1/move", json_body={"id": "m1b"})
    assert mail_move(ctx, "m1", "Archive") == {"moved": True, "id": "m1b"}
    assert router.calls[0][2] == {"destinationId": "archive"}


def test_move_to_subfolder_of_inbox(ctx, router):
    for name, fid in [
        ("inbox", "INBOX"),
        ("archive", "ARCH"),
        ("deleteditems", "DEL"),
        ("junkemail", "JUNK"),
    ]:
        router.add("GET", f"/me/mailFolders/{name}", json_body={"id": fid})
    router.add("GET", "/me/mailFolders/sub1", json_body={"id": "sub1", "parentFolderId": "INBOX"})
    router.add("POST", "/me/messages/m1/move", json_body={"id": "m1b"})
    assert mail_move(ctx, "m1", "sub1")["moved"] is True


@pytest.mark.parametrize("target", ["drafts", "outbox", "sentitems"])
def test_move_to_forbidden_wellknown(ctx, router, target):
    for name, fid in [
        ("inbox", "INBOX"),
        ("archive", "ARCH"),
        ("deleteditems", "DEL"),
        ("junkemail", "JUNK"),
    ]:
        router.add("GET", f"/me/mailFolders/{name}", json_body={"id": fid})
    router.add(
        "GET",
        f"/me/mailFolders/{target}",
        json_body={"id": target.upper(), "parentFolderId": "ROOT"},
    )
    router.add("GET", "/me/mailFolders/ROOT", json_body={"id": "ROOT"})
    with pytest.raises(PolicyError, match="Zielordner"):
        mail_move(ctx, "m1", target)
    assert not router.sent("POST", "/me/messages/m1/move")
