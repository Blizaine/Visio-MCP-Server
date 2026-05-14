"""
Shape-level tools: add, connect, text, list.
"""

from __future__ import annotations

from typing import Optional

from ..com.app import get_visio_app
from ..com.document import ensure_document_open
from ..com.undo import undo_scope
from ..errors import ShapeNotFound, envelope
from ..server_instance import mcp


def _find_shape(page, shape_id: int):
    for shape in page.Shapes:
        if shape.ID == shape_id:
            return shape
    return None


@mcp.tool()
@envelope("add_shape")
async def add_shape(file_path: str, shape_type: str, x: float, y: float,
                    width: Optional[float] = 1.0, height: Optional[float] = 1.0,
                    page_name: Optional[str] = None) -> dict:
    """Add a shape to a Visio document. Creates the file if it doesn't exist.

    Args:
        file_path: Path to the Visio file.
        shape_type: "Rectangle" (default), "Circle"/"Ellipse", or "Line".
                    Unrecognized values fall back to Rectangle.
        x, y: Lower-left corner of the shape's bounding box (page units).
        width, height: Shape size (page units, default 1.0).
        page_name: Name of the target page. Defaults to the active page
                   (or the document's first page if a different document
                   is currently active).

    Returns:
        {"shape_id": int, "shape_type": str, "x": float, "y": float,
         "width": float, "height": float, "page_name": str}
    """
    handle = ensure_document_open(file_path, create_if_missing=True)
    page = handle.get_page(page_name)

    kind = shape_type.lower()
    with undo_scope(f"Add {shape_type}"):
        if kind == "rectangle":
            shape = page.DrawRectangle(x, y, x + width, y + height)
        elif kind in ("circle", "ellipse"):
            shape = page.DrawOval(x, y, x + width, y + height)
        elif kind == "line":
            shape = page.DrawLine(x, y, x + width, y + height)
        else:
            shape = page.DrawRectangle(x, y, x + width, y + height)
        if shape is not None:
            shape.Text = shape_type

    handle.save()
    return {
        "shape_id": int(shape.ID),
        "shape_type": shape_type,
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "page_name": page.Name,
    }


@mcp.tool()
@envelope("connect_shapes")
async def connect_shapes(file_path: str, shape1_id: int, shape2_id: int,
                         connector_type: Optional[str] = "Dynamic",
                         page_name: Optional[str] = None) -> dict:
    """Connect two shapes with a connector.

    Args:
        file_path: Path to the Visio file.
        shape1_id, shape2_id: IDs of the shapes to connect. Both must live
                              on the same page.
        connector_type: "Dynamic" (default) or "Straight". "Curved" is
                       accepted but currently routes as Dynamic until real
                       curved-connector support lands in a later phase.
        page_name: Page the shapes live on. Defaults to the active page.

    Returns:
        {"connector_id": int, "shape1_id": int, "shape2_id": int,
         "connector_type": str, "page_name": str}
    """
    handle = ensure_document_open(file_path)
    app = get_visio_app()
    page = handle.get_page(page_name)

    shape1 = _find_shape(page, shape1_id)
    shape2 = _find_shape(page, shape2_id)
    if shape1 is None or shape2 is None:
        raise ShapeNotFound(
            f"Could not find shapes with IDs {shape1_id} and {shape2_id}",
            details={"shape1_id": shape1_id, "shape2_id": shape2_id},
        )

    with undo_scope("Connect shapes"):
        connector = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)

        # ShapeRouteStyle: 2 = visLORouteStraight; default leaves Visio's
        # dynamic right-angle routing.
        if (connector_type or "").lower() == "straight":
            try:
                connector.Cells("ShapeRouteStyle").Formula = "2"
            except Exception:
                pass  # cell unavailable on some connector masters

        # Dynamic shape-glue: gluing an endpoint to PinX tells Visio to pick
        # the best connection point automatically as the shapes move.
        connector.Cells("BeginX").GlueTo(shape1.Cells("PinX"))
        connector.Cells("EndX").GlueTo(shape2.Cells("PinX"))

    handle.save()
    return {
        "connector_id": int(connector.ID),
        "shape1_id": shape1_id,
        "shape2_id": shape2_id,
        "connector_type": connector_type or "Dynamic",
        "page_name": page.Name,
    }


@mcp.tool()
@envelope("add_text")
async def add_text(file_path: str, shape_id: int, text: str,
                   page_name: Optional[str] = None) -> dict:
    """Set the text of a shape.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape.
        text: Text to set.
        page_name: Page the shape lives on. Defaults to the active page.

    Returns:
        {"shape_id": int, "page_name": str}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    target = _find_shape(page, shape_id)
    if target is None:
        raise ShapeNotFound(
            f"Could not find shape with ID {shape_id} on page '{page.Name}'",
            details={"shape_id": shape_id, "page_name": page.Name},
        )

    with undo_scope("Set shape text"):
        target.Text = text

    handle.save()
    return {"shape_id": shape_id, "page_name": page.Name}


@mcp.tool()
@envelope("list_shapes")
async def list_shapes(file_path: str, page_name: Optional[str] = None) -> dict:
    """List all shapes on a page.

    Args:
        file_path: Path to the Visio file.
        page_name: Name of the page to list. Defaults to the active page.

    Returns:
        {"page_name": str,
         "shapes": [{"id", "name", "text", "type",
                     "position": {"x", "y"}, "size": {"width", "height"}}, ...]}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    shapes_info = []
    for shape in page.Shapes:
        shapes_info.append({
            "id": int(shape.ID),
            "name": shape.Name,
            "text": shape.Text,
            "type": shape.Type,
            "position": {
                "x": float(shape.Cells("PinX").Result("in")),
                "y": float(shape.Cells("PinY").Result("in")),
            },
            "size": {
                "width": float(shape.Cells("Width").Result("in")),
                "height": float(shape.Cells("Height").Result("in")),
            },
        })
    return {"page_name": page.Name, "shapes": shapes_info}
