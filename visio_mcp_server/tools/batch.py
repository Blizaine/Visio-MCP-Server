"""
Batch tools — the recommended path for any operation affecting more than
one shape.

Each batch tool:
- Makes one MCP round-trip instead of N (one model decision, one COM session).
- Disables `Application.ScreenUpdating` for the duration of the batch so
  Visio doesn't repaint between operations.
- Wraps the whole batch in a single `undo_scope` — partial failures roll
  back cleanly.
- Does NOT save automatically; call `save_document` or `close_document`.

Compared to looping single-shape tools, batches are typically 5-20x faster
end-to-end because the dominant cost is the per-call model round-trip,
not the COM work itself.
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import Optional

from ..com.app import get_visio_app
from ..com.document import ensure_document_open, find_shape_on_page
from ..com.undo import undo_scope
from ..errors import InvalidArgument, ShapeNotFound, envelope
from ..server_instance import mcp
from .styling import apply_fill, apply_line, apply_text_format


@contextmanager
def _batch_context(scope_name: str):
    """Disable Visio's screen redraws and wrap the work in one undo scope.

    Restoring `ScreenUpdating` on exit is essential — if we left it false,
    the UI would stay frozen for the rest of the session.
    """
    app = get_visio_app()
    prev = app.ScreenUpdating
    app.ScreenUpdating = False
    try:
        with undo_scope(scope_name):
            yield
    finally:
        app.ScreenUpdating = prev


def _draw_shape(page, shape_type: str, x: float, y: float,
                width: float, height: float):
    kind = (shape_type or "Rectangle").lower()
    if kind == "rectangle":
        return page.DrawRectangle(x, y, x + width, y + height)
    if kind in ("circle", "ellipse"):
        return page.DrawOval(x, y, x + width, y + height)
    if kind == "line":
        return page.DrawLine(x, y, x + width, y + height)
    return page.DrawRectangle(x, y, x + width, y + height)


# ---------------------------------------------------------------------------
# add_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
@envelope("add_shapes")
async def add_shapes(file_path: str, shapes: list, page_name: Optional[str] = None) -> dict:
    """Add many shapes in a single batch. Strongly preferred over multiple
    `add_shape` calls — typically 5-20x faster end-to-end.

    Each entry in `shapes` is a dict:
        {
          "shape_type": "Rectangle" | "Circle" | "Ellipse" | "Line",
          "x": float, "y": float,             # bottom-left corner, in inches
          "width": float, "height": float,    # in inches, default 1.0
          "text": str                          # optional, sets shape.Text
        }

    Creates the file if it doesn't exist (same as `add_shape`).
    Does not auto-save; call `save_document` or `close_document`.

    Args:
        file_path: Path to the Visio file.
        shapes: List of shape specs (see above).
        page_name: Target page. Defaults to the active page.

    Returns:
        {"page_name": str, "count": int,
         "shapes": [{"shape_id", "shape_type", "x", "y", "width", "height", "text"}]}
    """
    if not shapes:
        raise InvalidArgument("add_shapes called with empty shapes list")

    handle = ensure_document_open(file_path, create_if_missing=True)
    page = handle.get_page(page_name)

    results = []
    with _batch_context(f"Add {len(shapes)} shapes"):
        for i, spec in enumerate(shapes):
            try:
                shape_type = spec.get("shape_type", "Rectangle")
                x = float(spec["x"])
                y = float(spec["y"])
                width = float(spec.get("width", 1.0))
                height = float(spec.get("height", 1.0))
            except (KeyError, TypeError, ValueError) as e:
                raise InvalidArgument(
                    f"shapes[{i}] is missing required fields or has bad types: {e}",
                    details={"index": i, "spec": spec},
                )

            shape = _draw_shape(page, shape_type, x, y, width, height)
            text = spec.get("text")
            if text is not None:
                shape.Text = str(text)
            else:
                shape.Text = shape_type  # match add_shape's default behavior

            results.append({
                "shape_id": int(shape.ID),
                "shape_type": shape_type,
                "x": x, "y": y, "width": width, "height": height,
                "text": text if text is not None else shape_type,
            })

    return {"page_name": page.Name, "count": len(results), "shapes": results}


# ---------------------------------------------------------------------------
# connect_shapes_bulk
# ---------------------------------------------------------------------------

@mcp.tool()
@envelope("connect_shapes_bulk")
async def connect_shapes_bulk(file_path: str, connections: list,
                              page_name: Optional[str] = None) -> dict:
    """Create many connectors in a single batch. Preferred over multiple
    `connect_shapes` calls.

    Each entry in `connections`:
        {
          "shape1_id": int, "shape2_id": int,
          "connector_type": "Dynamic" | "Straight"   # optional, default "Dynamic"
        }

    Does not auto-save.

    Args:
        file_path: Path to the Visio file.
        connections: List of connection specs.
        page_name: Page the shapes live on. Defaults to the active page.

    Returns:
        {"page_name": str, "count": int,
         "connectors": [{"connector_id", "shape1_id", "shape2_id", "connector_type"}]}
    """
    if not connections:
        raise InvalidArgument("connect_shapes_bulk called with empty connections list")

    handle = ensure_document_open(file_path)
    app = get_visio_app()
    page = handle.get_page(page_name)

    results = []
    with _batch_context(f"Connect {len(connections)} pairs"):
        for i, conn in enumerate(connections):
            try:
                s1_id = int(conn["shape1_id"])
                s2_id = int(conn["shape2_id"])
            except (KeyError, TypeError, ValueError) as e:
                raise InvalidArgument(
                    f"connections[{i}] missing shape1_id/shape2_id: {e}",
                    details={"index": i, "spec": conn},
                )
            ctype = conn.get("connector_type", "Dynamic")

            s1 = find_shape_on_page(page, s1_id)
            s2 = find_shape_on_page(page, s2_id)
            if s1 is None or s2 is None:
                raise ShapeNotFound(
                    f"connections[{i}]: shape IDs {s1_id}/{s2_id} not on page '{page.Name}'",
                    details={"index": i, "shape1_id": s1_id, "shape2_id": s2_id, "page_name": page.Name},
                )

            connector = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)
            if (ctype or "").lower() == "straight":
                try:
                    connector.Cells("ShapeRouteStyle").Formula = "2"
                except Exception:
                    pass
            connector.Cells("BeginX").GlueTo(s1.Cells("PinX"))
            connector.Cells("EndX").GlueTo(s2.Cells("PinX"))

            results.append({
                "connector_id": int(connector.ID),
                "shape1_id": s1_id,
                "shape2_id": s2_id,
                "connector_type": ctype or "Dynamic",
            })

    return {"page_name": page.Name, "count": len(results), "connectors": results}


# ---------------------------------------------------------------------------
# style_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
@envelope("style_shapes")
async def style_shapes(file_path: str, updates: list,
                       page_name: Optional[str] = None) -> dict:
    """Apply fill / line / text-format / text changes to many shapes at once.

    Strongly preferred over looping `set_shape_fill`, `set_shape_line`,
    `set_shape_text_format`, or `add_text` — typically 5-20x faster.

    Each entry in `updates` targets one shape and bundles whatever you want
    to change. All sub-specs are optional; include only the keys you want.

        {
          "shape_id": int,
          "fill": {"color": str, "pattern"?: int},
          "line": {"color"?: str, "weight"?: float (pt), "pattern"?: int},
          "text": str,                             # sets shape.Text
          "text_format": {
              "font"?: str, "size"?: float (pt), "color"?: str,
              "bold"?: bool, "italic"?: bool, "underline"?: bool,
              "align"?: "left" | "center" | "right" | "justify"
          }
        }

    Color values accept #RRGGBB, RGB(r,g,b), or named colors. `bold`/`italic`/
    `underline` toggle individual bits — pass only the ones you want to change.

    Does not auto-save; call `save_document` or `close_document`.

    Args:
        file_path: Path to the Visio file.
        updates: List of per-shape update specs.
        page_name: Page the shapes live on. Defaults to the active page.

    Returns:
        {"page_name": str, "count": int,
         "results": [{"shape_id": int, "applied": {...}}, ...]}
    """
    if not updates:
        raise InvalidArgument("style_shapes called with empty updates list")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    doc = handle.com_doc

    results = []
    with _batch_context(f"Style {len(updates)} shapes"):
        for i, upd in enumerate(updates):
            try:
                shape_id = int(upd["shape_id"])
            except (KeyError, TypeError, ValueError) as e:
                raise InvalidArgument(
                    f"updates[{i}] missing shape_id: {e}",
                    details={"index": i, "spec": upd},
                )

            shape = find_shape_on_page(page, shape_id)
            if shape is None:
                raise ShapeNotFound(
                    f"updates[{i}]: shape ID {shape_id} not on page '{page.Name}'",
                    details={"index": i, "shape_id": shape_id, "page_name": page.Name},
                )

            applied: dict = {}
            if upd.get("fill"):
                applied["fill"] = apply_fill(shape, upd["fill"])
            if upd.get("line"):
                applied["line"] = apply_line(shape, upd["line"])
            if upd.get("text") is not None:
                shape.Text = str(upd["text"])
                applied["text"] = str(upd["text"])
            if upd.get("text_format"):
                applied["text_format"] = apply_text_format(shape, doc, upd["text_format"])

            if not applied:
                raise InvalidArgument(
                    f"updates[{i}]: no changes specified (need fill, line, text, or text_format)",
                    details={"index": i, "shape_id": shape_id},
                )

            results.append({"shape_id": shape_id, "applied": applied})

    return {"page_name": page.Name, "count": len(results), "results": results}


# ---------------------------------------------------------------------------
# delete_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
@envelope("delete_shapes")
async def delete_shapes(file_path: str, shape_ids: list,
                        page_name: Optional[str] = None) -> dict:
    """Delete one or more shapes from a page.

    The IDs do not need to be in any particular order. Connectors glued to
    deleted shapes are kept (they become floating connectors); delete those
    explicitly if you want them gone.

    Does not auto-save.

    Args:
        file_path: Path to the Visio file.
        shape_ids: List of shape IDs to delete.
        page_name: Page the shapes live on. Defaults to the active page.

    Returns:
        {"page_name": str, "deleted_ids": [int, ...], "count": int}
    """
    if not shape_ids:
        raise InvalidArgument("delete_shapes called with empty shape_ids list")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    deleted = []
    with _batch_context(f"Delete {len(shape_ids)} shapes"):
        # Resolve all targets BEFORE deleting any. If any miss, we abort the
        # batch (the undo scope rolls back) so the caller gets an all-or-nothing
        # operation.
        targets = []
        for sid in shape_ids:
            try:
                sid_int = int(sid)
            except (TypeError, ValueError):
                raise InvalidArgument(f"shape_ids entry {sid!r} is not an int")
            shape = find_shape_on_page(page, sid_int)
            if shape is None:
                raise ShapeNotFound(
                    f"shape ID {sid_int} not on page '{page.Name}'",
                    details={"shape_id": sid_int, "page_name": page.Name},
                )
            targets.append((sid_int, shape))

        for sid_int, shape in targets:
            shape.Delete()
            deleted.append(sid_int)

    return {"page_name": page.Name, "deleted_ids": deleted, "count": len(deleted)}


# ---------------------------------------------------------------------------
# transform_shapes
# ---------------------------------------------------------------------------

@mcp.tool()
@envelope("transform_shapes")
async def transform_shapes(file_path: str, updates: list,
                           page_name: Optional[str] = None) -> dict:
    """Move, resize, and/or rotate many shapes in a single batch.

    Each entry in `updates`:
        {
          "shape_id": int,
          "x"?: float,             # new bottom-left X, in inches
          "y"?: float,             # new bottom-left Y, in inches
          "width"?: float,         # new width, in inches
          "height"?: float,        # new height, in inches
          "angle_degrees"?: float  # new rotation, in degrees
        }

    All position/size fields are optional — pass only what you want to change.
    Positions assume the shape's pin is at its center (the default for
    shapes created via this server). For shapes with custom LocPin offsets
    the position may be slightly off.

    Does not auto-save.

    Args:
        file_path: Path to the Visio file.
        updates: List of per-shape transform specs.
        page_name: Page the shapes live on. Defaults to the active page.

    Returns:
        {"page_name": str, "count": int,
         "results": [{"shape_id": int, "applied": {...}}, ...]}
    """
    if not updates:
        raise InvalidArgument("transform_shapes called with empty updates list")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    results = []
    with _batch_context(f"Transform {len(updates)} shapes"):
        for i, upd in enumerate(updates):
            try:
                shape_id = int(upd["shape_id"])
            except (KeyError, TypeError, ValueError) as e:
                raise InvalidArgument(
                    f"updates[{i}] missing shape_id: {e}",
                    details={"index": i, "spec": upd},
                )

            shape = find_shape_on_page(page, shape_id)
            if shape is None:
                raise ShapeNotFound(
                    f"updates[{i}]: shape ID {shape_id} not on page '{page.Name}'",
                    details={"index": i, "shape_id": shape_id, "page_name": page.Name},
                )

            keys = ("x", "y", "width", "height", "angle_degrees")
            if all(upd.get(k) is None for k in keys):
                raise InvalidArgument(
                    f"updates[{i}]: no changes specified (need at least one of {keys})",
                    details={"index": i, "shape_id": shape_id},
                )

            # Read current size so a partial update (e.g. only x) still
            # produces a sensible new pin position.
            current_w = float(shape.Cells("Width").Result("in"))
            current_h = float(shape.Cells("Height").Result("in"))
            new_w = float(upd["width"]) if upd.get("width") is not None else current_w
            new_h = float(upd["height"]) if upd.get("height") is not None else current_h

            applied: dict = {}
            if upd.get("width") is not None:
                shape.Cells("Width").Formula = f"{new_w} in"
                applied["width"] = new_w
            if upd.get("height") is not None:
                shape.Cells("Height").Formula = f"{new_h} in"
                applied["height"] = new_h
            if upd.get("x") is not None:
                # Pin is at the center for our default shapes.
                shape.Cells("PinX").Formula = f"{float(upd['x']) + new_w / 2.0} in"
                applied["x"] = float(upd["x"])
            if upd.get("y") is not None:
                shape.Cells("PinY").Formula = f"{float(upd['y']) + new_h / 2.0} in"
                applied["y"] = float(upd["y"])
            if upd.get("angle_degrees") is not None:
                shape.Cells("Angle").Formula = f"{float(upd['angle_degrees'])} deg"
                applied["angle_degrees"] = float(upd["angle_degrees"])

            results.append({"shape_id": shape_id, "applied": applied})

    return {"page_name": page.Name, "count": len(results), "results": results}
