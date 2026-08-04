from __future__ import annotations

import base64
import json
import urllib.parse

import httpx
import pytest

from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.config import Settings
from outlook_web_mcp.models import CredentialBundle


def _jwt(claims: dict[str, object]) -> str:
    raw = json.dumps(claims, separators=(",", ":")).encode()
    payload = base64.urlsafe_b64encode(raw).decode().rstrip("=")
    return f"e30.{payload}.signature"


class StubAuth:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.forces: list[bool] = []
        self.token = _jwt({"puid": "12345", "tid": "tenant-guid"})

    async def get_access_token(self, *, force: bool = False) -> str:
        self.forces.append(force)
        return self.token

    def identity(self) -> CredentialBundle:
        return CredentialBundle("unused", "client", "tenant-guid", puid="12345")


@pytest.mark.asyncio
async def test_call_uses_fixed_owa_headers_and_retries_one_401() -> None:
    settings = Settings(credential_profile="api-test")
    auth = StubAuth(settings)
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        if len(requests) == 1:
            return httpx.Response(401, json={"error": "expired"})
        return httpx.Response(
            200,
            json={
                "Body": {
                    "ResponseMessages": {
                        "Items": [{"ResponseClass": "Success", "ResponseCode": "NoError"}]
                    }
                }
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        result = await api.call(
            "GetFolder",
            {
                "__type": "GetFolderRequest:#Exchange",
                "FolderIds": [],
            },
        )

    assert "Body" in result
    assert auth.forces == [False, True]
    assert len(requests) == 2
    request = requests[-1]
    assert request.url.host == "outlook.cloud.microsoft"
    assert request.url.path == "/owa/service.svc"
    assert request.url.params["action"] == "GetFolder"
    assert request.headers["authorization"] == f"Bearer {auth.token}"
    assert request.headers["x-anchormailbox"] == "PUID:12345@tenant-guid"
    assert request.headers["origin"] == "https://outlook.cloud.microsoft"
    encoded = request.headers["x-owa-urlpostdata"]
    envelope = json.loads(urllib.parse.unquote(encoded))
    assert envelope["__type"] == "GetFolderJsonRequest:#Exchange"
    assert envelope["Header"]["RequestServerVersion"] == "V2018_01_08"


@pytest.mark.asyncio
async def test_ews_error_becomes_sanitized_api_error() -> None:
    settings = Settings(credential_profile="api-error-test")
    auth = StubAuth(settings)

    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "Body": {
                    "ResponseMessages": {
                        "Items": [
                            {
                                "ResponseClass": "Error",
                                "ResponseCode": "ErrorInvalidIdMalformed",
                                "MessageText": "The id is malformed",
                            }
                        ]
                    }
                }
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        with pytest.raises(Exception, match="The id is malformed") as raised:
            await api.call("GetItem", {"__type": "GetItemRequest:#Exchange"})

    assert "Bearer" not in str(raised.value)
