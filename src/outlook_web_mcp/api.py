from __future__ import annotations

import json
import urllib.parse
import uuid
from typing import Any

import httpx

from .auth import AuthManager, jwt_claims
from .config import OWA_ENDPOINT, OWA_ORIGIN, Settings
from .errors import ApiError
from .models import CredentialBundle

URL_POST_DATA_MAX = 4_096


def _record(value: object) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _records(value: object) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def response_messages(data: object) -> list[dict[str, Any]]:
    root = _record(data)
    body = _record(root.get("Body"))
    messages = _record(body.get("ResponseMessages"))
    return _records(messages.get("Items"))


def response_items(data: object) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for message in response_messages(data):
        items.extend(_records(message.get("Items")))
    return items


def root_items(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        root = _record(message.get("RootFolder"))
        items = _records(root.get("Items"))
        if items:
            return items
    return []


def root_folders(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        root = _record(message.get("RootFolder"))
        folders = _records(root.get("Folders"))
        if folders:
            return folders
    return []


def _first_error(value: object) -> tuple[str, str] | None:
    if isinstance(value, dict):
        response_class = str(value.get("ResponseClass") or "")
        response_code = str(value.get("ResponseCode") or "")
        if response_class.casefold() == "error" or (response_code and response_code != "NoError"):
            message = str(value.get("MessageText") or response_code or "OWA request failed")
            return response_code or "owa_error", message[:500]
        error_code = value.get("ErrorCode")
        if error_code not in (None, 0, "0", "NoError"):
            message = str(value.get("FaultMessage") or value.get("Message") or error_code)
            return str(error_code), message[:500]
        for child in value.values():
            error = _first_error(child)
            if error is not None:
                return error
    elif isinstance(value, list):
        for child in value:
            error = _first_error(child)
            if error is not None:
                return error
    return None


class OutlookWebApi:
    """Thin client for the private EWS-JSON endpoint used by Outlook Web."""

    def __init__(
        self,
        auth: AuthManager,
        settings: Settings | None = None,
        http: httpx.AsyncClient | None = None,
    ) -> None:
        self.auth = auth
        self.settings = settings or auth.settings
        self._http = http
        self._owned_http: httpx.AsyncClient | None = None
        self._session_id = str(uuid.uuid4())

    async def __aenter__(self) -> OutlookWebApi:
        return self

    async def __aexit__(self, *_: object) -> None:
        await self.aclose()

    async def aclose(self) -> None:
        if self._owned_http is not None:
            await self._owned_http.aclose()
            self._owned_http = None

    def _client(self) -> httpx.AsyncClient:
        if self._http is not None:
            return self._http
        if self._owned_http is None:
            self._owned_http = httpx.AsyncClient(
                timeout=self.settings.request_timeout_seconds,
                follow_redirects=False,
                headers={"User-Agent": "outlook-web-mcp-server/0.1"},
            )
        return self._owned_http

    def envelope(self, action: str, body: dict[str, Any]) -> dict[str, Any]:
        return {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": self.settings.time_zone_id,
                    },
                },
            },
            "Body": body,
        }

    async def call(
        self,
        action: str,
        body: dict[str, Any],
        *,
        app: str = "Mail",
    ) -> dict[str, Any]:
        envelope = self.envelope(action, body)
        response = await self._send(action, envelope, app=app, force_refresh=False)
        if response.status_code == 401:
            response = await self._send(action, envelope, app=app, force_refresh=True)
        if 300 <= response.status_code < 400:
            raise ApiError(
                "Outlook Web unexpectedly redirected an API request; "
                "refusing to forward credentials.",
                "unexpected_redirect",
                http_status=response.status_code,
            )
        if len(response.content) > self.settings.max_response_bytes:
            raise ApiError(
                "Outlook Web response exceeded the configured size limit.",
                "response_too_large",
                http_status=response.status_code,
            )
        try:
            parsed = response.json()
        except json.JSONDecodeError as exc:
            raise ApiError(
                f"Outlook Web returned a non-JSON response (HTTP {response.status_code}).",
                "invalid_json",
                http_status=response.status_code,
            ) from exc
        if not isinstance(parsed, dict):
            raise ApiError(
                "Outlook Web returned an unexpected JSON shape.",
                "invalid_shape",
                http_status=response.status_code,
            )
        if response.status_code != 200:
            error = _first_error(parsed)
            response_code = error[0] if error else None
            message = error[1] if error else f"HTTP {response.status_code}"
            raise ApiError(
                f"Outlook Web {action} failed: {message}",
                "http_error",
                http_status=response.status_code,
                response_code=response_code,
            )
        error = _first_error(parsed)
        if error is not None:
            raise ApiError(
                f"Outlook Web {action} failed: {error[1]}",
                "ews_error",
                http_status=response.status_code,
                response_code=error[0],
            )
        return parsed

    async def _send(
        self,
        action: str,
        envelope: dict[str, Any],
        *,
        app: str,
        force_refresh: bool,
    ) -> httpx.Response:
        token = await self.auth.get_access_token(force=force_refresh)
        identity = self.auth.identity()
        raw = json.dumps(envelope, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
        encoded = urllib.parse.quote_from_bytes(raw, safe="")
        headers = self._headers(action, app, token, identity)
        content: bytes | None
        if len(encoded.encode("ascii")) <= URL_POST_DATA_MAX:
            headers["X-OWA-UrlPostData"] = encoded
            content = None
        else:
            content = raw
        url = f"{OWA_ENDPOINT}?action={urllib.parse.quote(action)}&app={urllib.parse.quote(app)}"
        return await self._client().post(url, headers=headers, content=content)

    def _headers(
        self,
        action: str,
        app: str,
        token: str,
        identity: CredentialBundle,
    ) -> dict[str, str]:
        claims = jwt_claims(token)
        puid = str(claims.get("puid") or identity.puid)
        tenant_id = str(claims.get("tid") or identity.tenant_id)
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
            "Action": action,
            "Prefer": (
                'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"'
            ),
            "X-OWA-CorrelationId": str(uuid.uuid4()),
            "X-OWA-SessionId": self._session_id,
            "X-OWA-Hosted-UX": "false",
            "X-Req-Source": app,
            "Origin": OWA_ORIGIN,
        }
        if puid and tenant_id:
            headers["X-AnchorMailbox"] = f"PUID:{puid}@{tenant_id}"
        return headers
