"""
stdio MCP bridge to the LibreOffice extension HTTP API on localhost:8765.
Discovers tools dynamically from GET /tools and forwards calls to POST /tools/{name}.
"""

import asyncio
import json
import os
import sys

import httpx
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

ENDPOINT = os.environ.get("LIBREOFFICE_MCP_URL", "http://localhost:8765")
TIMEOUT = httpx.Timeout(5.0, connect=1.0)

server = Server("libreoffice-live")


@server.list_tools()
async def list_tools() -> list[Tool]:
    try:
        async with httpx.AsyncClient(timeout=TIMEOUT) as client:
            resp = await client.get(f"{ENDPOINT}/tools")
            resp.raise_for_status()
            data = resp.json()
    except Exception as e:
        return [
            Tool(
                name="libreoffice_unavailable",
                description=(
                    f"LibreOffice extension HTTP API at {ENDPOINT} is not reachable: {e}. "
                    "Open LibreOffice and click 'MCP Server → Start MCP Server' in the menu bar."
                ),
                inputSchema={"type": "object", "properties": {}},
            )
        ]

    return [
        Tool(
            name=t["name"],
            description=t.get("description", ""),
            inputSchema=t.get("parameters") or {"type": "object", "properties": {}},
        )
        for t in data.get("tools", [])
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict | None) -> list[TextContent]:
    if name == "libreoffice_unavailable":
        return [TextContent(type="text", text=f"LibreOffice extension HTTP API at {ENDPOINT} is not reachable.")]

    payload = arguments or {}
    try:
        async with httpx.AsyncClient(timeout=TIMEOUT) as client:
            resp = await client.post(f"{ENDPOINT}/tools/{name}", json=payload)
        body = resp.text
        try:
            body = json.dumps(resp.json(), ensure_ascii=False, indent=2)
        except Exception:
            pass
        if resp.is_error:
            return [TextContent(type="text", text=f"HTTP {resp.status_code}: {body}")]
        return [TextContent(type="text", text=body)]
    except httpx.HTTPError as e:
        return [TextContent(type="text", text=f"Request to {ENDPOINT}/tools/{name} failed: {e}")]


async def main() -> None:
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        sys.exit(0)
