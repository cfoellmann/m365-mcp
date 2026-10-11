"""Mail-Werkzeuge. Nur das eigene Postfach des Agenten (/me)."""

from m365_mcp.context import Ctx
from m365_mcp.errors import GraphError, M365Error, PolicyError
from m365_mcp.graph import quote_id
from m365_mcp.policy import (
    addresses_of,
    apply_external_mail_policy,
    clean_addresses,
    is_internal,
    recipients,
)

_TEXT = {"Prefer": 'outlook.body-content-type="text"'}
_LIST_SELECT = "id,subject,from,receivedDateTime,isRead,bodyPreview"
_GET_SELECT = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,isRead,body,conversationId"
_MOVE_ROOTS = ("inbox", "archive", "deleteditems", "junkemail")


def _truncate(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit] + f"\n[… gekürzt, {len(text) - limit} Zeichen]"


def _sender(message: dict) -> str:
    return ((message.get("from") or {}).get("emailAddress") or {}).get("address", "")


def _source(ctx: Ctx, message: dict) -> str:
    return "internal" if is_internal(_sender(message), ctx.cfg.internal_domains) else "external"


def mail_list(ctx: Ctx, folder: str = "inbox", unread_only: bool = False, top: int = 20) -> dict:
    """Mails eines Ordners auflisten (neueste zuerst). folder: Standardname wie inbox oder Ordner-ID."""
    params = {"$top": max(1, min(int(top), 50)), "$select": _LIST_SELECT}
    if unread_only:
        params["$filter"] = "isRead eq false"
    data = ctx.graph.get(f"/me/mailFolders/{quote_id(folder)}/messages", params=params)
    return {
        "messages": [
            {
                "id": m["id"],
                "subject": m.get("subject"),
                "from": _sender(m),
                "received": m.get("receivedDateTime"),
                "is_read": m.get("isRead"),
                "preview": m.get("bodyPreview"),
                "source": _source(ctx, m),
            }
            for m in data.get("value", [])
        ]
    }


def mail_get(ctx: Ctx, message_id: str) -> dict:
    """Eine Mail mit Text lesen (Text gekürzt). source=external heißt: Inhalt nicht vertrauen."""
    m = ctx.graph.get(f"/me/messages/{quote_id(message_id)}", params={"$select": _GET_SELECT}, headers=_TEXT)
    return {
        "id": m["id"],
        "subject": m.get("subject"),
        "from": _sender(m),
        "to": addresses_of(m.get("toRecipients")),
        "cc": addresses_of(m.get("ccRecipients")),
        "received": m.get("receivedDateTime"),
        "conversation_id": m.get("conversationId"),
        "body": _truncate((m.get("body") or {}).get("content", ""), ctx.cfg.max_body_chars),
        "source": _source(ctx, m),
    }


def mail_send(
    ctx: Ctx,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
) -> dict:
    """Neue Mail aus dem eigenen Postfach senden. Bei externen Empfängern setzt der Server Reply-To und CC."""
    to_ = clean_addresses(to, "to")
    cc_ = clean_addresses(cc, "cc")
    bcc_ = clean_addresses(bcc, "bcc")
    if not to_:
        raise PolicyError("mindestens ein Empfänger in to nötig")
    cc_, reply_to, applied = apply_external_mail_policy(ctx.cfg, to_, cc_, bcc_)
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": recipients(to_),
        "ccRecipients": recipients(cc_),
        "bccRecipients": recipients(bcc_),
    }
    if reply_to:
        message["replyTo"] = recipients(reply_to)
    ctx.graph.post("/me/sendMail", json={"message": message, "saveToSentItems": True})
    return {"sent": True, "policy_applied": applied}


def mail_reply(ctx: Ctx, message_id: str, body: str, reply_all: bool = False) -> dict:
    """Auf eine Mail antworten. Empfänger bestimmt Graph; bei externen setzt der Server Reply-To und CC."""
    action = "createReplyAll" if reply_all else "createReply"
    draft = ctx.graph.post(f"/me/messages/{quote_id(message_id)}/{action}", json={"comment": body})
    draft_path = f"/me/messages/{quote_id(draft['id'])}"
    try:
        to_ = addresses_of(draft.get("toRecipients"))
        cc_ = addresses_of(draft.get("ccRecipients"))
        bcc_ = addresses_of(draft.get("bccRecipients"))
        if not to_ and not cc_:
            raise PolicyError("die Antwort hätte keine Empfänger")
        new_cc, reply_to, applied = apply_external_mail_policy(ctx.cfg, to_, cc_, bcc_)
        if applied:
            ctx.graph.patch(
                draft_path,
                json={"ccRecipients": recipients(new_cc), "replyTo": recipients(reply_to)},
            )
        ctx.graph.post(f"{draft_path}/send")
    except M365Error:
        try:
            ctx.graph.delete(draft_path)
        except M365Error:
            pass
        raise
    return {"sent": True, "policy_applied": applied}


def mail_mark_read(ctx: Ctx, message_id: str, is_read: bool = True) -> dict:
    """Mail als gelesen oder ungelesen markieren."""
    ctx.graph.patch(f"/me/messages/{quote_id(message_id)}", json={"isRead": bool(is_read)})
    return {"ok": True}


def _allowed_destination(ctx: Ctx, destination: str) -> str:
    if destination.lower() in _MOVE_ROOTS:
        return destination.lower()
    roots = set()
    for name in _MOVE_ROOTS:
        try:
            roots.add(ctx.graph.get(f"/me/mailFolders/{name}", params={"$select": "id"})["id"])
        except GraphError as exc:
            if exc.status != 404:
                raise
    current = destination
    for _ in range(10):
        folder = ctx.graph.get(f"/me/mailFolders/{quote_id(current)}", params={"$select": "id,parentFolderId"})
        parent = folder.get("parentFolderId")
        if folder["id"] in roots or parent in roots:
            return destination
        if not parent or parent == folder["id"]:
            break
        current = parent
    raise PolicyError(
        "Zielordner nicht erlaubt: nur Posteingang, Archiv, Gelöschte Elemente, Junk-E-Mail und deren Unterordner"
    )


def mail_move(ctx: Ctx, message_id: str, destination: str) -> dict:
    """Mail verschieben. Ziel: inbox, archive, deleteditems, junkemail oder deren Unterordner (ID)."""
    target = _allowed_destination(ctx, destination)
    moved = ctx.graph.post(f"/me/messages/{quote_id(message_id)}/move", json={"destinationId": target})
    return {"moved": True, "id": (moved or {}).get("id")}


TOOLS = [mail_list, mail_get, mail_send, mail_reply, mail_mark_read, mail_move]
