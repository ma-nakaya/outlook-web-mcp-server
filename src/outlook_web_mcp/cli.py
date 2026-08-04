from __future__ import annotations

import argparse
import asyncio
import json
import sys

from .api import OutlookWebApi
from .auth import AuthManager, DeviceCodePrompt, PendingDeviceLogin
from .config import Settings
from .credentials import KeyringPendingLoginStore
from .errors import OutlookWebMcpError


def _json(value: object) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _show_code(prompt: DeviceCodePrompt) -> None:
    print(f"Open: {prompt.verification_uri}")
    print(f"Code: {prompt.user_code}")


async def _run(args: argparse.Namespace) -> int:
    settings = Settings()
    pending_store = KeyringPendingLoginStore(settings.credential_profile)
    async with AuthManager(settings) as auth:
        if args.command == "login":
            bundle = await auth.login(
                open_browser=not args.no_browser,
                on_code=_show_code,
            )
            _json(
                {
                    "authenticated": True,
                    "profile": settings.credential_profile,
                    "tenant_id": bundle.tenant_id,
                    "display_name": bundle.display_name,
                }
            )
            return 0
        if args.command == "login-start":
            pending = await auth.begin_login(open_browser=not args.no_browser)
            pending_store.save(pending.to_secret_dict())
            _json(
                {
                    "verification_uri": pending.prompt.verification_uri,
                    "verification_uri_complete": pending.prompt.verification_uri_complete,
                    "user_code": pending.prompt.user_code,
                    "expires_in": pending.prompt.expires_in,
                }
            )
            return 0
        if args.command == "login-complete":
            raw = pending_store.load()
            if raw is None:
                raise OutlookWebMcpError(
                    "No pending login. Run `outlook-web-mcp login-start` first.",
                    "pending_missing",
                )
            pending = PendingDeviceLogin.from_secret_dict(raw)
            try:
                bundle = await auth.complete_login(pending)
            finally:
                pending_store.delete()
            _json(
                {
                    "authenticated": True,
                    "profile": settings.credential_profile,
                    "tenant_id": bundle.tenant_id,
                    "display_name": bundle.display_name,
                }
            )
            return 0
        if args.command == "status":
            _json(await auth.ensure_status())
            return 0
        if args.command == "logout":
            auth.logout()
            pending_store.delete()
            _json({"authenticated": False, "profile": settings.credential_profile})
            return 0
        if args.command == "probe":
            async with OutlookWebApi(auth, settings) as api:
                data = await api.call(
                    "GetFolder",
                    {
                        "__type": "GetFolderRequest:#Exchange",
                        "FolderShape": {
                            "__type": "FolderResponseShape:#Exchange",
                            "BaseShape": "IdOnly",
                        },
                        "FolderIds": [
                            {
                                "__type": "DistinguishedFolderId:#Exchange",
                                "Id": "inbox",
                            }
                        ],
                    },
                )
            _json(
                {
                    "success": True,
                    "response_has_body": isinstance(data.get("Body"), dict),
                }
            )
            return 0
    return 2


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(
        prog="outlook-web-mcp",
        description="Unofficial Outlook Web MCP server without Microsoft Graph or COM.",
    )
    sub = root.add_subparsers(dest="command", required=True)
    login = sub.add_parser("login", help="Sign in with Microsoft's device-code flow")
    login.add_argument("--no-browser", action="store_true")
    login_start = sub.add_parser(
        "login-start",
        help="Start sign-in and securely save the short-lived device code",
    )
    login_start.add_argument("--no-browser", action="store_true")
    sub.add_parser("login-complete", help="Complete a previously started sign-in")
    sub.add_parser("status", help="Check and refresh authentication")
    sub.add_parser("logout", help="Remove the saved refresh token")
    sub.add_parser("probe", help="Run a content-free Inbox API health check")
    sub.add_parser("serve", help="Run the stdio MCP server")
    return root


def main() -> None:
    args = parser().parse_args()
    if args.command == "serve":
        from .server import main as server_main

        server_main()
        return
    try:
        raise SystemExit(asyncio.run(_run(args)))
    except OutlookWebMcpError as exc:
        print(f"{exc.code}: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
