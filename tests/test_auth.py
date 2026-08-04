from __future__ import annotations

import base64
import json
import time
from typing import Any

import httpx
import pytest

from outlook_web_mcp.auth import AuthManager, DeviceCodePrompt, PendingDeviceLogin
from outlook_web_mcp.config import OUTLOOK_AUDIENCE, OUTLOOK_SCOPE, Settings
from outlook_web_mcp.models import CredentialBundle


def _jwt(claims: dict[str, object]) -> str:
    def part(value: dict[str, object]) -> str:
        raw = json.dumps(value, separators=(",", ":")).encode()
        return base64.urlsafe_b64encode(raw).decode().rstrip("=")

    return f"{part({'alg': 'none'})}.{part(claims)}.signature"


class MemoryStore:
    def __init__(self, bundle: CredentialBundle | None = None) -> None:
        self.bundle = bundle
        self.save_count = 0

    def load(self) -> CredentialBundle | None:
        return self.bundle

    def save(self, bundle: CredentialBundle) -> None:
        self.bundle = bundle
        self.save_count += 1

    def delete(self) -> None:
        self.bundle = None


@pytest.mark.asyncio
async def test_device_code_uses_outlook_audience_scope() -> None:
    seen: dict[str, str] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        form = dict(httpx.QueryParams(request.content.decode()))
        seen.update(form)
        return httpx.Response(
            200,
            json={
                "device_code": "device-secret",
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "interval": 5,
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        auth = AuthManager(Settings(credential_profile="device-test"), MemoryStore(), client)
        device_code, prompt = await auth.request_device_code()

    assert device_code == "device-secret"
    assert prompt.user_code == "ABCD-EFGH"
    assert seen["scope"] == OUTLOOK_SCOPE
    assert seen["client_id"] == Settings().client_id


@pytest.mark.asyncio
async def test_device_code_redemption_has_no_browser_origin(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = _jwt(
        {
            "aud": OUTLOOK_AUDIENCE,
            "scp": "OWA.AccessAsUser.All",
            "tid": "tenant-id",
            "puid": "puid-value",
            "exp": int(time.time()) + 3600,
        }
    )

    async def no_sleep(_: float) -> None:
        return None

    monkeypatch.setattr("outlook_web_mcp.auth.asyncio.sleep", no_sleep)

    def handler(request: httpx.Request) -> httpx.Response:
        assert "origin" not in request.headers
        form = dict(httpx.QueryParams(request.content.decode()))
        assert form["grant_type"] == "urn:ietf:params:oauth:grant-type:device_code"
        return httpx.Response(
            200,
            json={
                "access_token": token,
                "refresh_token": "refresh-secret",
                "expires_in": 3600,
            },
        )

    prompt = DeviceCodePrompt("CODE", "https://example.test", "", 900, 1)
    pending = PendingDeviceLogin(
        "device-secret",
        prompt,
        Settings().client_id,
        time.time() + 900,
    )
    store = MemoryStore()
    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        auth = AuthManager(Settings(credential_profile="redeem-test"), store, client)
        bundle = await auth.complete_login(pending)

    assert bundle.tenant_id == "tenant-id"
    assert store.bundle is bundle


@pytest.mark.asyncio
async def test_refresh_rotates_secret_and_status_never_exposes_it() -> None:
    token = _jwt(
        {
            "aud": OUTLOOK_AUDIENCE,
            "scp": "OWA.AccessAsUser.All Mail.ReadWrite",
            "tid": "tenant-id",
            "puid": "puid-value",
            "name": "Mailbox User",
            "exp": int(time.time()) + 3600,
        }
    )
    store = MemoryStore(
        CredentialBundle(
            refresh_token="old-refresh-secret",
            client_id=Settings().client_id,
            tenant_id="tenant-id",
        )
    )
    request_form: dict[str, str] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        request_form.update(dict(httpx.QueryParams(request.content.decode())))
        assert "origin" not in request.headers
        return httpx.Response(
            200,
            json={
                "access_token": token,
                "refresh_token": "rotated-refresh-secret",
                "expires_in": 3600,
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        auth = AuthManager(Settings(credential_profile="refresh-test"), store, client)
        assert await auth.get_access_token() == token
        status: dict[str, Any] = await auth.ensure_status()

    assert request_form["grant_type"] == "refresh_token"
    assert request_form["scope"] == OUTLOOK_SCOPE
    assert store.bundle is not None
    assert store.bundle.refresh_token == "rotated-refresh-secret"
    assert store.bundle.puid == "puid-value"
    assert store.save_count == 1
    serialized = json.dumps(status)
    assert "refresh_token" not in serialized
    assert "old-refresh-secret" not in serialized
    assert "rotated-refresh-secret" not in serialized
    assert status["usable"] is True
