from __future__ import annotations

import pytest

from outlook_web_mcp.server import mcp


@pytest.mark.asyncio
async def test_expected_tools_and_annotations_are_registered() -> None:
    tools = await mcp.list_tools()
    by_name = {tool.name: tool for tool in tools}
    assert set(by_name) == {
        "outlook_auth_status",
        "search_emails",
        "list_mail_folders",
        "get_email",
        "set_email_read_state",
        "list_calendar_events",
        "create_reply_draft",
    }
    assert by_name["search_emails"].annotations.readOnlyHint is True
    assert by_name["get_email"].annotations.readOnlyHint is True
    assert by_name["set_email_read_state"].annotations.readOnlyHint is False
    assert by_name["create_reply_draft"].annotations.readOnlyHint is False
    assert by_name["create_reply_draft"].annotations.idempotentHint is False

    search_properties = by_name["search_emails"].inputSchema["properties"]
    assert "daysBack" in search_properties
    assert "maxResults" in search_properties
    assert "includeBodyPreview" in search_properties
    assert "includeSubfolders" in search_properties
