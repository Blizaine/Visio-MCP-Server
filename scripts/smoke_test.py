"""
End-to-end smoke test for the Visio MCP Server.

Spawns the server as a subprocess via the MCP stdio transport, runs each
tool against a real Visio instance, and prints the envelope each tool
returns. Visio will become visible during the test.

Usage (from the repo root, after `uv pip install -e .`):

    .\\.venv\\Scripts\\python.exe scripts\\smoke_test.py

Exits 0 on success, 1 on any tool-level failure or transport error.
"""

import asyncio
import json
import os
import sys
import time

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def _server_params() -> StdioServerParameters:
    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    venv_python = os.path.join(repo_root, ".venv", "Scripts", "python.exe")
    if not os.path.exists(venv_python):
        venv_python = sys.executable  # fall back to current interpreter
    return StdioServerParameters(
        command=venv_python,
        args=["-m", "visio_mcp_server.visio_server"],
        cwd=repo_root,
        env=None,
    )


def _envelope_ok(result) -> tuple[bool, dict]:
    """Parse an MCP CallToolResult into (ok, envelope_dict)."""
    if not result.content:
        return False, {"raw": "(no content)"}
    text = result.content[0].text
    try:
        env = json.loads(text)
    except json.JSONDecodeError:
        return False, {"raw": text}
    return bool(env.get("ok")), env


async def _expect_ok(session: ClientSession, name: str, args: dict) -> dict:
    """Call a tool and assert the envelope reports ok=true. Returns data."""
    print(f"\n--- call {name}({args}) ---")
    result = await session.call_tool(name, args)
    ok, env = _envelope_ok(result)
    print(json.dumps(env, indent=2))
    if not ok:
        raise SystemExit(f"FAIL: tool {name} returned ok=false")
    return env.get("data", {})


async def run() -> None:
    test_path = os.path.join(
        os.environ.get("TEMP", os.path.expanduser("~")),
        f"visio_mcp_smoke_{int(time.time())}.vsdx",
    )

    print(f"Test file will be saved to: {test_path}")
    print("Spawning the MCP server...")

    async with stdio_client(_server_params()) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            print("MCP session initialized")

            tools = await session.list_tools()
            tool_names = sorted(t.name for t in tools.tools)
            print(f"Tools advertised: {tool_names}")
            expected = {"add_shape", "add_text", "close_document",
                        "connect_shapes", "create_visio_file",
                        "list_shapes", "open_visio_file"}
            missing = expected - set(tool_names)
            if missing:
                raise SystemExit(f"FAIL: tools missing from server: {missing}")

            await _expect_ok(session, "create_visio_file", {"save_path": test_path})

            data = await _expect_ok(session, "add_shape", {
                "file_path": test_path, "shape_type": "Rectangle",
                "x": 2.0, "y": 2.0, "width": 1.5, "height": 1.0,
            })
            id1 = data["shape_id"]

            data = await _expect_ok(session, "add_shape", {
                "file_path": test_path, "shape_type": "Circle",
                "x": 5.0, "y": 2.0, "width": 1.0, "height": 1.0,
            })
            id2 = data["shape_id"]

            await _expect_ok(session, "connect_shapes", {
                "file_path": test_path,
                "shape1_id": id1, "shape2_id": id2,
                "connector_type": "Straight",
            })

            await _expect_ok(session, "add_text", {
                "file_path": test_path, "shape_id": id1, "text": "Start",
            })
            await _expect_ok(session, "add_text", {
                "file_path": test_path, "shape_id": id2, "text": "End",
            })

            data = await _expect_ok(session, "list_shapes", {"file_path": test_path})
            shape_count = len(data.get("shapes", []))
            print(f"\nlist_shapes reports {shape_count} shapes on the page")
            if shape_count < 3:
                raise SystemExit(
                    f"FAIL: expected at least 3 shapes (2 rectangles/circles + 1 connector), got {shape_count}"
                )

            await _expect_ok(session, "close_document", {
                "file_path": test_path, "save_changes": True,
            })

    print(f"\nSUCCESS — all tools exercised. File saved at: {test_path}")
    print("Open it in Visio and verify visually:")
    print("  - one rectangle labeled 'Start' (left)")
    print("  - one circle labeled 'End' (right)")
    print("  - a STRAIGHT connector between them (must be visible — this checks the LinePattern fix)")


if __name__ == "__main__":
    asyncio.run(run())
