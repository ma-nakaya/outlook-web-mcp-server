from __future__ import annotations

import os
from dataclasses import dataclass, field

from platformdirs import user_cache_path

DEFAULT_OUTLOOK_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48"
OUTLOOK_AUDIENCE = "https://outlook.office.com"
OUTLOOK_SCOPE = f"{OUTLOOK_AUDIENCE}/.default openid profile offline_access"
OWA_ORIGIN = "https://outlook.cloud.microsoft"
OWA_ENDPOINT = f"{OWA_ORIGIN}/owa/service.svc"
STORE_ID = "outlook-web"


def _int_env(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None:
        return default
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc


@dataclass(frozen=True, slots=True)
class Settings:
    credential_profile: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_CREDENTIAL_PROFILE", "default").strip() or "default"
        )
    )
    client_id: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_CLIENT_ID", DEFAULT_OUTLOOK_CLIENT_ID).strip()
            or DEFAULT_OUTLOOK_CLIENT_ID
        )
    )
    time_zone_id: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_TIME_ZONE_ID", "Tokyo Standard Time").strip()
            or "Tokyo Standard Time"
        )
    )
    request_timeout_seconds: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS", 30)
    )
    max_response_bytes: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_RESPONSE_BYTES", 10_000_000)
    )
    max_scanned_items: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_SCANNED_ITEMS", 5_000)
    )
    max_scanned_folders: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_SCANNED_FOLDERS", 100)
    )
    log_level: str = field(
        default_factory=lambda: os.getenv("OUTLOOK_WEB_LOG_LEVEL", "INFO").strip().upper()
    )

    @property
    def lock_path(self) -> str:
        root = user_cache_path("outlook-web-mcp-server", "ma-nakaya", ensure_exists=True)
        return str(root / f"auth-{self.credential_profile}.lock")

    def validate(self) -> None:
        if not self.client_id:
            raise ValueError("OUTLOOK_WEB_CLIENT_ID must not be empty")
        if not self.time_zone_id:
            raise ValueError("OUTLOOK_WEB_TIME_ZONE_ID must not be empty")
        if self.request_timeout_seconds < 1:
            raise ValueError("OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS must be positive")
        if self.max_response_bytes < 1024:
            raise ValueError("OUTLOOK_WEB_MAX_RESPONSE_BYTES must be at least 1024")
        if not 1 <= self.max_scanned_items <= 50_000:
            raise ValueError("OUTLOOK_WEB_MAX_SCANNED_ITEMS must be between 1 and 50000")
        if not 1 <= self.max_scanned_folders <= 1_000:
            raise ValueError("OUTLOOK_WEB_MAX_SCANNED_FOLDERS must be between 1 and 1000")
