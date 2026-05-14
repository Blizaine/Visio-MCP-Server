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
            expected = {"add_page", "add_shape", "add_shapes", "add_text",
                        "close_document", "connect_shapes", "connect_shapes_bulk",
                        "create_visio_file", "delete_page", "delete_shapes",
                        "duplicate_page", "export_page", "export_pdf",
                        "list_pages", "list_shapes", "open_visio_file",
                        "save_document", "set_active_page", "set_shape_fill",
                        "set_shape_line", "set_shape_text_format",
                        "style_shapes", "transform_shapes"}
            missing = expected - set(tool_names)
            if missing:
                raise SystemExit(f"FAIL: tools missing from server: {missing}")

            await _expect_ok(session, "create_visio_file", {"save_path": test_path})

            # Sanity-check the new doc's starting page list.
            data = await _expect_ok(session, "list_pages", {"file_path": test_path})
            initial_pages = data["pages"]
            if len(initial_pages) != 1:
                raise SystemExit(f"FAIL: new document should have 1 page, got {len(initial_pages)}")
            page1_name = initial_pages[0]["name"]
            print(f"\nInitial page: {page1_name!r}")

            # Original flowchart, drawn on page 1 by default.
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
            if shape_count < 3:
                raise SystemExit(
                    f"FAIL: expected at least 3 shapes on {page1_name!r}, got {shape_count}"
                )

            # Phase 3 coverage: add a second page and draw on it explicitly.
            data = await _expect_ok(session, "add_page", {
                "file_path": test_path, "name": "Page 2",
                "width": 11.0, "height": 8.5,
            })
            page2_name = data["name"]

            data = await _expect_ok(session, "list_pages", {"file_path": test_path})
            if len(data["pages"]) != 2:
                raise SystemExit(f"FAIL: expected 2 pages after add_page, got {len(data['pages'])}")

            data = await _expect_ok(session, "add_shape", {
                "file_path": test_path, "shape_type": "Rectangle",
                "x": 3.0, "y": 3.0, "width": 2.0, "height": 1.0,
                "page_name": page2_name,
            })
            page2_shape_id = data["shape_id"]
            if data.get("page_name") != page2_name:
                raise SystemExit(f"FAIL: add_shape on {page2_name!r} reported page_name={data.get('page_name')!r}")

            await _expect_ok(session, "add_text", {
                "file_path": test_path, "shape_id": page2_shape_id, "text": "On Page 2",
                "page_name": page2_name,
            })

            data = await _expect_ok(session, "list_shapes", {
                "file_path": test_path, "page_name": page2_name,
            })
            page2_shapes = data["shapes"]
            if len(page2_shapes) != 1 or page2_shapes[0]["text"] != "On Page 2":
                raise SystemExit(f"FAIL: page 2 should have 1 shape with text 'On Page 2', got {page2_shapes}")

            # duplicate the second page and verify count
            data = await _expect_ok(session, "duplicate_page", {
                "file_path": test_path, "source_page_name": page2_name,
                "new_name": "Page 2 Copy",
            })
            copy_name = data["new_name"]

            data = await _expect_ok(session, "list_pages", {"file_path": test_path})
            if len(data["pages"]) != 3:
                raise SystemExit(f"FAIL: expected 3 pages after duplicate, got {len(data['pages'])}")

            # set_active_page back to the first page
            await _expect_ok(session, "set_active_page", {
                "file_path": test_path, "page_name": page1_name,
            })

            # delete the duplicate
            await _expect_ok(session, "delete_page", {
                "file_path": test_path, "page_name": copy_name,
            })

            data = await _expect_ok(session, "list_pages", {"file_path": test_path})
            if len(data["pages"]) != 2:
                raise SystemExit(f"FAIL: expected 2 pages after delete, got {len(data['pages'])}")

            # Phase 4 coverage: styling the page-1 shapes.
            await _expect_ok(session, "set_shape_fill", {
                "file_path": test_path, "shape_id": id1, "color": "#90CAF9",
            })
            await _expect_ok(session, "set_shape_line", {
                "file_path": test_path, "shape_id": id2,
                "color": "#1565C0", "weight": 2.5, "pattern": 1,
            })
            await _expect_ok(session, "set_shape_text_format", {
                "file_path": test_path, "shape_id": id1,
                "size": 16, "bold": True, "color": "#0D47A1", "align": "center",
            })

            # Phase 4 coverage: exports. Save outputs next to the source.
            base = os.path.splitext(test_path)[0]
            png_out = base + ".png"
            pdf_out = base + ".pdf"

            png_data = await _expect_ok(session, "export_page", {
                "file_path": test_path, "output_path": png_out,
            })
            if png_data["bytes"] <= 0 or not os.path.exists(png_out):
                raise SystemExit(f"FAIL: PNG export did not produce a file at {png_out}")
            print(f"PNG size: {png_data['bytes']} bytes")

            pdf_data = await _expect_ok(session, "export_pdf", {
                "file_path": test_path, "output_path": pdf_out,
            })
            if pdf_data["bytes"] <= 0 or not os.path.exists(pdf_out):
                raise SystemExit(f"FAIL: PDF export did not produce a file at {pdf_out}")
            with open(pdf_out, "rb") as fh:
                magic = fh.read(4)
            if magic != b"%PDF":
                raise SystemExit(f"FAIL: PDF magic bytes wrong: got {magic!r}")
            print(f"PDF size: {pdf_data['bytes']} bytes, magic OK")

            # ---------------------------------------------------------------
            # Phase 5 coverage: batch tools, delete, transform, explicit save.
            # We build a second diagram entirely via batch tools and time it
            # so the perf win is observable.
            # ---------------------------------------------------------------
            batch_path = os.path.join(
                os.path.dirname(test_path),
                f"visio_mcp_batch_{int(time.time())}.vsdx",
            )

            await _expect_ok(session, "create_visio_file", {"save_path": batch_path})

            t0 = time.perf_counter()
            # 5 boxes + 2 circles, all in one call.
            batch_shapes = [
                {"shape_type": "Rectangle", "x": 0.5, "y": 6.5, "width": 1.2, "height": 0.8, "text": "A"},
                {"shape_type": "Rectangle", "x": 2.5, "y": 6.5, "width": 1.2, "height": 0.8, "text": "B"},
                {"shape_type": "Rectangle", "x": 4.5, "y": 6.5, "width": 1.2, "height": 0.8, "text": "C"},
                {"shape_type": "Rectangle", "x": 0.5, "y": 4.5, "width": 1.2, "height": 0.8, "text": "D"},
                {"shape_type": "Rectangle", "x": 2.5, "y": 4.5, "width": 1.2, "height": 0.8, "text": "E"},
                {"shape_type": "Circle",    "x": 4.5, "y": 4.5, "width": 1.0, "height": 1.0, "text": "F"},
                {"shape_type": "Circle",    "x": 1.5, "y": 2.5, "width": 1.0, "height": 1.0, "text": "G"},
            ]
            data = await _expect_ok(session, "add_shapes", {
                "file_path": batch_path, "shapes": batch_shapes,
            })
            shape_ids = [s["shape_id"] for s in data["shapes"]]
            if len(shape_ids) != len(batch_shapes):
                raise SystemExit(f"FAIL: add_shapes returned {len(shape_ids)} shapes, expected {len(batch_shapes)}")

            # 6 connectors in one call.
            connections = [
                {"shape1_id": shape_ids[0], "shape2_id": shape_ids[1], "connector_type": "Straight"},
                {"shape1_id": shape_ids[1], "shape2_id": shape_ids[2], "connector_type": "Straight"},
                {"shape1_id": shape_ids[0], "shape2_id": shape_ids[3], "connector_type": "Straight"},
                {"shape1_id": shape_ids[3], "shape2_id": shape_ids[4], "connector_type": "Straight"},
                {"shape1_id": shape_ids[4], "shape2_id": shape_ids[5], "connector_type": "Straight"},
                {"shape1_id": shape_ids[3], "shape2_id": shape_ids[6], "connector_type": "Straight"},
            ]
            await _expect_ok(session, "connect_shapes_bulk", {
                "file_path": batch_path, "connections": connections,
            })

            # Style every shape blue with white bold text — in one call.
            style_updates = [
                {
                    "shape_id": sid,
                    "fill": {"color": "#1976D2"},
                    "text_format": {"color": "#FFFFFF", "bold": True, "size": 14, "align": "center"},
                }
                for sid in shape_ids
            ]
            await _expect_ok(session, "style_shapes", {
                "file_path": batch_path, "updates": style_updates,
            })

            elapsed = time.perf_counter() - t0
            print(f"\n>>> Batch build of 7 shapes + 6 connectors + 7 styles: {elapsed:.2f}s")

            # Explicit save (singles no longer auto-save).
            await _expect_ok(session, "save_document", {"file_path": batch_path})

            # transform_shapes: enlarge shape A and rotate it.
            await _expect_ok(session, "transform_shapes", {
                "file_path": batch_path,
                "updates": [
                    {"shape_id": shape_ids[0], "width": 2.0, "height": 1.2},
                    {"shape_id": shape_ids[6], "x": 5.5, "y": 2.5, "angle_degrees": 45.0},
                ],
            })

            # delete_shapes: remove G (the last circle).
            data = await _expect_ok(session, "delete_shapes", {
                "file_path": batch_path, "shape_ids": [shape_ids[6]],
            })
            if data["count"] != 1 or shape_ids[6] not in data["deleted_ids"]:
                raise SystemExit(f"FAIL: delete_shapes didn't remove the expected shape: {data}")

            # Verify the deletion stuck.
            data = await _expect_ok(session, "list_shapes", {"file_path": batch_path})
            remaining_ids = [s["id"] for s in data["shapes"]]
            if shape_ids[6] in remaining_ids:
                raise SystemExit(f"FAIL: deleted shape {shape_ids[6]} still in list_shapes")

            await _expect_ok(session, "close_document", {
                "file_path": batch_path, "save_changes": True,
            })

            await _expect_ok(session, "close_document", {
                "file_path": test_path, "save_changes": True,
            })

    print(f"\nSUCCESS - all tools exercised. Outputs:")
    print(f"  .vsdx (singles):    {test_path}")
    print(f"  .png:               {png_out}")
    print(f"  .pdf:               {pdf_out}")
    print(f"  .vsdx (batch built): {batch_path}")
    print("")
    print("Open the single-tool .vsdx and verify visually:")
    print(f"  - Page '{page1_name}': rectangle 'Start' (LIGHT BLUE fill, BOLD navy text, centered),")
    print("    circle 'End' (NAVY outline at 2.5pt), with a STRAIGHT connector between them.")
    print(f"  - Page 'Page 2': rectangle 'On Page 2'.")
    print("  - Only 2 pages should remain (the duplicate was deleted).")
    print("")
    print("Open the batch-built .vsdx and verify visually:")
    print("  - 6 blue shapes with bold white centered labels A-F (G was deleted).")
    print("  - A is wider than the rest (transform_shapes resize).")
    print("  - Connectors form a tree: A-B-C across the top, A-D-E-F down/right.")


if __name__ == "__main__":
    asyncio.run(run())
