from __future__ import annotations

import html
import re
from html.parser import HTMLParser
from typing import Any

from .models import CalendarEvent, FolderInfo, MailDetail, MailSummary


def record(value: object) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def records(value: object) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def string(value: object) -> str:
    return str(value) if value is not None else ""


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        del attrs
        if tag.casefold() in {"br", "p", "div", "li", "tr", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.casefold() in {"p", "div", "li", "tr", "h1", "h2", "h3", "h4"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        self.parts.append(data)


def html_to_text(value: str) -> str:
    parser = _TextExtractor()
    try:
        parser.feed(value)
        raw = "".join(parser.parts)
    except Exception:
        raw = re.sub(r"<[^>]+>", " ", value)
    raw = html.unescape(raw).replace("\r\n", "\n").replace("\r", "\n")
    raw = re.sub(r"[\t\f\v ]+", " ", raw)
    raw = re.sub(r" *\n *", "\n", raw)
    return re.sub(r"\n{3,}", "\n\n", raw).strip()


def mailbox(value: object) -> tuple[str, str]:
    current = record(value)
    nested = record(current.get("Mailbox"))
    if nested:
        current = nested
    nested_address = record(current.get("EmailAddress"))
    if nested_address:
        current = nested_address
    return (
        string(current.get("Name") or current.get("DisplayName")),
        string(current.get("EmailAddress") or current.get("Address")),
    )


def sender(item: dict[str, Any]) -> tuple[str, str]:
    for key in ("From", "Sender", "Organizer"):
        name, address = mailbox(item.get(key))
        if name or address:
            return name, address
    return "", ""


def recipient_text(value: object) -> str:
    formatted: list[str] = []
    for raw in records(value):
        name, address = mailbox(raw)
        if name and address and name.casefold() != address.casefold():
            formatted.append(f"{name} <{address}>")
        elif address:
            formatted.append(address)
        elif name:
            formatted.append(name)
    return "; ".join(formatted)


def item_id(item: dict[str, Any]) -> str:
    return string(record(item.get("ItemId")).get("Id"))


def change_key(item: dict[str, Any]) -> str:
    return string(record(item.get("ItemId")).get("ChangeKey"))


def body_text(item: dict[str, Any]) -> str:
    body = record(item.get("Body"))
    value = string(body.get("Value"))
    body_type = string(body.get("BodyType"))
    if body_type.casefold() == "html" or "<" in value:
        return html_to_text(value)
    return value.replace("\r\n", "\n").replace("\r", "\n")


def preview_text(item: dict[str, Any]) -> str:
    preview = string(item.get("Preview"))
    if "<" in preview:
        return html_to_text(preview)
    return preview.strip()


def received_at(item: dict[str, Any]) -> str:
    return string(
        item.get("DateTimeReceived") or item.get("DateTimeSent") or item.get("DateTimeCreated")
    )


def parse_mail_summary(item: dict[str, Any], *, include_preview: bool) -> MailSummary:
    sender_name, sender_address = sender(item)
    preview: str | None = None
    if include_preview:
        preview = preview_text(item)[:500]
    return MailSummary(
        email_id=item_id(item),
        subject=string(item.get("Subject")),
        sender_name=sender_name,
        sender_address=sender_address,
        received_at=received_at(item),
        is_unread=not bool(item.get("IsRead")),
        has_attachments=bool(item.get("HasAttachments")),
        body_preview=preview,
    )


def parse_mail_detail(item: dict[str, Any], *, max_body_characters: int) -> MailDetail:
    sender_name, sender_address = sender(item)
    body = body_text(item)
    truncated = len(body) > max_body_characters
    return MailDetail(
        email_id=item_id(item),
        subject=string(item.get("Subject")),
        sender_name=sender_name,
        sender_address=sender_address,
        to=recipient_text(item.get("ToRecipients")),
        cc=recipient_text(item.get("CcRecipients")),
        received_at=received_at(item),
        is_unread=not bool(item.get("IsRead")),
        has_attachments=bool(item.get("HasAttachments")),
        body=body[:max_body_characters],
        body_truncated=truncated,
    )


def parse_folder(
    item: dict[str, Any],
    *,
    folder_path: str,
    fallback_parent_id: str | None = None,
) -> FolderInfo:
    folder_id = string(record(item.get("FolderId")).get("Id"))
    parent_id = string(record(item.get("ParentFolderId")).get("Id")) or fallback_parent_id
    child_count = int(item.get("ChildFolderCount") or 0)
    return FolderInfo(
        folder_id=folder_id,
        name=string(item.get("FolderDisplayName") or item.get("DisplayName")),
        folder_path=folder_path,
        parent_folder_id=parent_id,
        unread_item_count=int(item.get("UnreadCount") or 0),
        total_item_count=int(item.get("TotalCount") or 0),
        has_children=child_count > 0,
    )


def _location(item: dict[str, Any]) -> str:
    value = item.get("Location")
    if isinstance(value, str):
        return value
    current = record(value)
    if current:
        return string(current.get("DisplayName") or current.get("LocationUri"))
    enhanced = record(item.get("EnhancedLocation"))
    return string(enhanced.get("DisplayName"))


def parse_calendar_event(item: dict[str, Any]) -> CalendarEvent:
    organizer_name, organizer_address = mailbox(item.get("Organizer"))
    organizer = organizer_address or organizer_name
    calendar_type = string(item.get("CalendarItemType"))
    is_recurring = bool(item.get("Recurrence")) or calendar_type in {
        "RecurringMaster",
        "Occurrence",
        "Exception",
    }
    return CalendarEvent(
        event_id=item_id(item),
        subject=string(item.get("Subject")),
        starts_at=string(item.get("Start")),
        ends_at=string(item.get("End")),
        location=_location(item),
        organizer=organizer,
        is_all_day=bool(item.get("IsAllDayEvent")),
        is_recurring=is_recurring,
        busy_status=string(item.get("FreeBusyType") or "Unknown"),
    )
