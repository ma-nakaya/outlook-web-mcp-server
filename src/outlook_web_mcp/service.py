from __future__ import annotations

from collections import deque
from contextlib import suppress
from datetime import UTC, datetime, timedelta
from typing import Any

from .api import OutlookWebApi, response_items, response_messages, root_folders, root_items
from .auth import AuthManager
from .config import STORE_ID, Settings
from .errors import ApiError, OutlookWebMcpError
from .models import CalendarEvent, FolderInfo, MailSummary
from .parsers import (
    change_key,
    item_id,
    parse_calendar_event,
    parse_folder,
    parse_mail_detail,
    parse_mail_summary,
    record,
    records,
    string,
)

DISTINGUISHED_FOLDERS = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sent_mail": "sentitems",
    "sentitems": "sentitems",
    "drafts": "drafts",
    "calendar": "calendar",
}


def _property(field_uri: str) -> dict[str, str]:
    return {"__type": "PropertyUri:#Exchange", "FieldURI": field_uri}


def _folder_id(folder_id: str, *, distinguished: bool = False) -> dict[str, str]:
    return {
        "__type": ("DistinguishedFolderId:#Exchange" if distinguished else "FolderId:#Exchange"),
        "Id": folder_id,
    }


def _target_folder(folder_id: str, *, distinguished: bool = False) -> dict[str, object]:
    return {
        "__type": "TargetFolderId:#Exchange",
        "BaseFolderId": _folder_id(folder_id, distinguished=distinguished),
    }


def _response_folders(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        folders = records(message.get("Folders"))
        if folders:
            return folders
    return []


def _parse_iso(value: str, parameter_name: str) -> datetime:
    normalized = value.strip().replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError as exc:
        raise ValueError(
            f"{parameter_name} must be an ISO 8601 date and time with an offset"
        ) from exc
    if parsed.tzinfo is None:
        raise ValueError(f"{parameter_name} must include a timezone offset")
    return parsed


def _utc_string(value: datetime) -> str:
    return value.astimezone(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")


def _calendar_string(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000")


class OutlookService:
    def __init__(
        self,
        auth: AuthManager,
        api: OutlookWebApi,
        settings: Settings | None = None,
    ) -> None:
        self.auth = auth
        self.api = api
        self.settings = settings or auth.settings

    async def auth_status(self) -> dict[str, object]:
        return await self.auth.ensure_status()

    @staticmethod
    def _validate_store(store_id: str) -> None:
        if store_id != STORE_ID:
            raise OutlookWebMcpError(
                "storeId is not an Outlook Web mailbox handle. Run the Outlook Web "
                "search or folder tool again and use the returned storeId.",
                "store_id_invalid",
            )

    async def _get_folder(
        self,
        folder_id: str,
        *,
        distinguished: bool,
        fallback_name: str = "",
    ) -> FolderInfo:
        data = await self.api.call(
            "GetFolder",
            {
                "__type": "GetFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "Default",
                    "AdditionalProperties": [
                        _property("ParentFolderId"),
                        _property("ChildFolderCount"),
                        _property("TotalCount"),
                        _property("UnreadCount"),
                    ],
                },
                "FolderIds": [_folder_id(folder_id, distinguished=distinguished)],
            },
        )
        folders = _response_folders(data)
        if not folders:
            raise ApiError("Outlook Web returned no folder.", "folder_not_found")
        raw = folders[0]
        name = string(raw.get("FolderDisplayName") or raw.get("DisplayName")) or fallback_name
        return parse_folder(raw, folder_path=name)

    async def _find_child_folders(self, parent_id: str, limit: int) -> list[dict[str, Any]]:
        data = await self.api.call(
            "FindFolder",
            {
                "__type": "FindFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "Default",
                    "AdditionalProperties": [
                        _property("ParentFolderId"),
                        _property("ChildFolderCount"),
                        _property("TotalCount"),
                        _property("UnreadCount"),
                        _property("FolderClass"),
                    ],
                },
                "ParentFolderIds": [_folder_id(parent_id)],
                "Traversal": "Shallow",
                "Paging": {
                    "__type": "IndexedPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "Offset": 0,
                    "MaxEntriesReturned": min(max(limit, 1), 200),
                },
            },
        )
        return root_folders(data)

    async def _resolve_folder(self, name: str) -> FolderInfo:
        normalized = name.strip().casefold()
        distinguished = DISTINGUISHED_FOLDERS.get(normalized)
        if distinguished is None:
            raise ValueError("Supported folders: inbox, sent, drafts")
        return await self._get_folder(
            distinguished,
            distinguished=True,
            fallback_name=normalized,
        )

    async def list_mail_folders(
        self,
        *,
        folder: str = "inbox",
        parent_folder_id: str | None = None,
        store_id: str | None = None,
        recursive: bool = True,
        max_results: int = 100,
        max_depth: int = 5,
    ) -> dict[str, object]:
        if not 1 <= max_results <= 200:
            raise ValueError("maxResults must be between 1 and 200")
        if not 1 <= max_depth <= 10:
            raise ValueError("maxDepth must be between 1 and 10")
        if parent_folder_id:
            if store_id is None:
                raise ValueError("storeId is required when parentFolderId is supplied")
            self._validate_store(store_id)
            root = await self._get_folder(
                parent_folder_id,
                distinguished=False,
                fallback_name="Selected folder",
            )
        else:
            root = await self._resolve_folder(folder)

        output: list[FolderInfo] = [root]
        queue: deque[tuple[FolderInfo, int]] = deque([(root, 0)])
        while recursive and queue and len(output) < max_results:
            parent, depth = queue.popleft()
            if depth >= max_depth:
                continue
            remaining = max_results - len(output)
            children = await self._find_child_folders(parent.folder_id, remaining)
            if children:
                parent.has_children = True
            for raw in children[:remaining]:
                name = string(raw.get("FolderDisplayName") or raw.get("DisplayName"))
                path = f"{parent.folder_path}\\{name}" if parent.folder_path else name
                child = parse_folder(
                    raw,
                    folder_path=path,
                    fallback_parent_id=parent.folder_id,
                )
                output.append(child)
                queue.append((child, depth + 1))
                if len(output) >= max_results:
                    break
        return {"result": [item.to_dict(STORE_ID) for item in output]}

    async def _mail_page(
        self,
        folder_id: str,
        *,
        unread_only: bool,
        include_preview: bool,
        offset: int,
        limit: int,
    ) -> tuple[list[dict[str, Any]], bool, int]:
        properties = [
            _property("Subject"),
            _property("From"),
            _property("Sender"),
            _property("DateTimeReceived"),
            _property("IsRead"),
            _property("HasAttachments"),
        ]
        if include_preview:
            properties.append(_property("Preview"))

        data = await self.api.call(
            "FindItem",
            {
                "__type": "FindItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "AdditionalProperties": properties,
                },
                "ParentFolderIds": [_folder_id(folder_id)],
                "Traversal": "Shallow",
                "Paging": {
                    "__type": "IndexedPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "Offset": offset,
                    "MaxEntriesReturned": limit,
                },
                "ViewFilter": "Unread" if unread_only else "All",
                "SortOrder": [
                    {
                        "__type": "SortResults:#Exchange",
                        "Order": "Descending",
                        "Path": _property("DateTimeReceived"),
                    }
                ],
            },
        )
        page = root_items(data)
        includes_last = True
        next_offset = offset + len(page)
        for message in response_messages(data):
            root = record(message.get("RootFolder"))
            if root:
                includes_last = bool(root.get("IncludesLastItemInRange", True))
                with suppress(TypeError, ValueError):
                    next_offset = int(root.get("IndexedPagingOffset") or next_offset)
                break
        return page, includes_last, next_offset

    async def _search_folder_ids(
        self,
        root: FolderInfo,
        *,
        include_subfolders: bool,
    ) -> list[str]:
        ids = [root.folder_id]
        if not include_subfolders:
            return ids
        queue: deque[tuple[str, int]] = deque([(root.folder_id, 0)])
        while queue and len(ids) < self.settings.max_scanned_folders:
            parent_id, depth = queue.popleft()
            if depth >= 10:
                continue
            children = await self._find_child_folders(
                parent_id,
                self.settings.max_scanned_folders - len(ids),
            )
            for raw in children:
                folder_class = string(raw.get("FolderClass"))
                if folder_class and not folder_class.startswith("IPF.Note"):
                    continue
                child_id = string(record(raw.get("FolderId")).get("Id"))
                if not child_id or child_id in ids:
                    continue
                ids.append(child_id)
                queue.append((child_id, depth + 1))
                if len(ids) >= self.settings.max_scanned_folders:
                    break
        return ids

    async def search_emails(
        self,
        *,
        folder: str = "inbox",
        query: str | None = None,
        days_back: int = 30,
        max_results: int = 20,
        include_body_preview: bool = False,
        folder_id: str | None = None,
        store_id: str | None = None,
        include_subfolders: bool = False,
        unread_only: bool = False,
    ) -> dict[str, object]:
        if not 1 <= days_back <= 365:
            raise ValueError("daysBack must be between 1 and 365")
        if not 1 <= max_results <= 50:
            raise ValueError("maxResults must be between 1 and 50")
        if folder_id:
            if store_id is None:
                raise ValueError("storeId is required when folderId is supplied")
            self._validate_store(store_id)
            root = await self._get_folder(
                folder_id,
                distinguished=False,
                fallback_name="Selected folder",
            )
        else:
            root = await self._resolve_folder(folder)
        folder_ids = await self._search_folder_ids(
            root,
            include_subfolders=include_subfolders,
        )
        cutoff_dt = datetime.now(UTC) - timedelta(days=days_back)
        normalized_query = (query or "").strip().casefold()
        found: list[MailSummary] = []
        seen: set[str] = set()
        scanned = 0
        page_size = 200

        for current_folder_id in folder_ids:
            offset = 0
            while scanned < self.settings.max_scanned_items:
                request_limit = min(page_size, self.settings.max_scanned_items - scanned)
                page, includes_last, next_offset = await self._mail_page(
                    current_folder_id,
                    unread_only=unread_only,
                    include_preview=include_body_preview,
                    offset=offset,
                    limit=request_limit,
                )
                scanned += len(page)
                reached_cutoff = False
                for raw in page:
                    summary = parse_mail_summary(
                        raw,
                        include_preview=include_body_preview,
                    )
                    if not summary.email_id or summary.email_id in seen:
                        continue
                    try:
                        received = _parse_iso(summary.received_at, "DateTimeReceived")
                    except ValueError:
                        received = None
                    if received is not None and received < cutoff_dt:
                        reached_cutoff = True
                        continue
                    if unread_only and not summary.is_unread:
                        continue
                    searchable = "\n".join(
                        [summary.subject, summary.sender_name, summary.sender_address]
                    ).casefold()
                    if normalized_query and normalized_query not in searchable:
                        continue
                    seen.add(summary.email_id)
                    found.append(summary)
                if reached_cutoff or includes_last or not page or next_offset <= offset:
                    break
                offset = next_offset
                if not normalized_query and not include_subfolders and len(found) >= max_results:
                    break
            if scanned >= self.settings.max_scanned_items:
                break

        found.sort(key=lambda item: item.received_at, reverse=True)
        selected = found[:max_results]
        return {"result": [item.to_dict(STORE_ID) for item in selected]}

    async def _get_item_raw(self, email_id: str, *, body_type: str = "Text") -> dict[str, Any]:
        data = await self.api.call(
            "GetItem",
            {
                "__type": "GetItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "AllProperties",
                    "BodyType": body_type,
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": email_id}],
            },
        )
        items = response_items(data)
        if not items:
            raise ApiError("The Outlook item was not found.", "item_not_found")
        return items[0]

    async def get_email(
        self,
        *,
        email_id: str,
        store_id: str,
        max_body_characters: int = 20_000,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        if not email_id.strip():
            raise ValueError("emailId is required")
        if not 1 <= max_body_characters <= 50_000:
            raise ValueError("maxBodyCharacters must be between 1 and 50000")
        raw = await self._get_item_raw(email_id)
        detail = parse_mail_detail(raw, max_body_characters=max_body_characters)
        return detail.to_dict(STORE_ID)

    async def set_email_read_state(
        self,
        *,
        email_id: str,
        store_id: str,
        is_read: bool,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        raw = await self._get_item_raw(email_id)
        key = change_key(raw)
        item_ref: dict[str, object] = {
            "__type": "ItemId:#Exchange",
            "Id": email_id,
        }
        if key:
            item_ref["ChangeKey"] = key
        await self.api.call(
            "UpdateItem",
            {
                "__type": "UpdateItemRequest:#Exchange",
                "ItemChanges": [
                    {
                        "__type": "ItemChange:#Exchange",
                        "ItemId": item_ref,
                        "Updates": [
                            {
                                "__type": "SetItemField:#Exchange",
                                "Path": _property("IsRead"),
                                "Item": {
                                    "__type": "Message:#Exchange",
                                    "IsRead": is_read,
                                },
                            }
                        ],
                    }
                ],
                "ConflictResolution": "AutoResolve",
                "MessageDisposition": "SaveOnly",
            },
        )
        summary = parse_mail_summary(raw, include_preview=False)
        return {
            "emailId": email_id,
            "storeId": STORE_ID,
            "subject": summary.subject,
            "receivedAt": summary.received_at,
            "isRead": is_read,
            "status": "read_state_updated",
        }

    async def create_reply_draft(
        self,
        *,
        email_id: str,
        store_id: str,
        body: str,
        reply_all: bool = False,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        if not body.strip():
            raise ValueError("body is required")
        if len(body) > 20_000:
            raise ValueError("body cannot exceed 20000 characters")
        original = await self._get_item_raw(email_id)
        key = change_key(original)
        reference: dict[str, object] = {
            "__type": "ItemId:#Exchange",
            "Id": email_id,
        }
        if key:
            reference["ChangeKey"] = key
        subject = string(original.get("Subject"))
        reply_subject = subject if subject.casefold().startswith("re:") else f"Re: {subject}"
        reply_item = {
            "__type": ("ReplyAllToItem:#Exchange" if reply_all else "ReplyToItem:#Exchange"),
            "ReferenceItemId": reference,
            "ShouldIgnoreChangeKey": True,
            "NewBodyContent": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "Text",
                "Value": body,
            },
            "Subject": reply_subject,
        }
        data = await self.api.call(
            "CreateItem",
            {
                "__type": "CreateItemRequest:#Exchange",
                "MessageDisposition": "SaveOnly",
                "SavedItemFolderId": _target_folder("drafts", distinguished=True),
                "Items": [reply_item],
            },
        )
        created = response_items(data)
        draft_id = item_id(created[0]) if created else ""
        if not draft_id:
            raise ApiError(
                "Outlook Web saved no identifiable reply draft.",
                "draft_id_missing",
            )
        try:
            draft = parse_mail_detail(
                await self._get_item_raw(draft_id),
                max_body_characters=1,
            )
            to = draft.to
            cc = draft.cc
            saved_subject = draft.subject or reply_subject
        except ApiError:
            to = ""
            cc = ""
            saved_subject = reply_subject
        return {
            "draftId": draft_id,
            "storeId": STORE_ID,
            "subject": saved_subject,
            "to": to,
            "cc": cc,
            "isReplyAll": reply_all,
            "status": "saved_to_drafts",
        }

    async def list_calendar_events(
        self,
        *,
        starts_after: str,
        ends_before: str,
        max_results: int = 50,
    ) -> dict[str, object]:
        start = _parse_iso(starts_after, "startsAfter")
        end = _parse_iso(ends_before, "endsBefore")
        if end <= start:
            raise ValueError("endsBefore must be later than startsAfter")
        if end - start > timedelta(days=31):
            raise ValueError("Calendar range cannot exceed 31 days")
        if not 1 <= max_results <= 100:
            raise ValueError("maxResults must be between 1 and 100")
        calendar = await self._get_folder(
            "calendar",
            distinguished=True,
            fallback_name="Calendar",
        )
        data = await self.api.call(
            "GetCalendarView",
            {
                "__type": "GetCalendarViewRequest:#Exchange",
                "CalendarId": _target_folder(calendar.folder_id),
                "RangeStart": _calendar_string(start),
                "RangeEnd": _calendar_string(end),
                "ClientSupportsIrm": True,
                "OptimizeExtendedPropertyLoading": True,
            },
            app="Calendar",
        )
        response_body = record(record(data).get("Body"))
        body_items = records(response_body.get("Items"))
        raw_events = body_items or response_items(data) or root_items(data)
        events: list[CalendarEvent] = [parse_calendar_event(raw) for raw in raw_events]
        events.sort(key=lambda item: item.starts_at)
        return {"result": [item.to_dict(STORE_ID) for item in events[:max_results]]}
