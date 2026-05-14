"""
Visio MCP Server

This MCP server provides tools for creating and editing Visio files.
It uses the Microsoft.Office.Interop.Visio API via Python's win32com interface.
"""

import os
import sys
import json
import atexit
import time
import winreg
from typing import Optional

import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("visio-server")

DEFAULT_SAVE_PATH = os.path.expandvars(r"%USERPROFILE%\Documents")

visio_app = None
open_documents: dict = {}  # keyed by _normalize_path(file_path)


# ----------------------------- helpers ---------------------------------------

def _normalize_path(p: str) -> str:
    """Return a canonical form so 'C:\\foo.vsdx' and 'c:/foo.vsdx' map to one key."""
    return os.path.normcase(os.path.abspath(p))


def _ensure_com_initialized() -> None:
    """Call CoInitialize on the current thread. Safe to call repeatedly."""
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass  # already initialized on this thread


def check_visio_installed() -> bool:
    """Check if Visio is installed without launching it."""
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        return True
    except OSError:
        return False


def get_visio_app():
    """Initialize or return the Visio Application object.

    Tries Dispatch -> dynamic.Dispatch -> DispatchEx in order; some Visio
    installations need the fallbacks because of how the COM type library is
    registered. Records every failure so we can report them all if we give up.
    """
    global visio_app
    if visio_app is not None:
        return visio_app

    attempts = (
        ("Dispatch", win32com.client.Dispatch),
        ("dynamic.Dispatch", win32com.client.dynamic.Dispatch),
        ("DispatchEx", win32com.client.DispatchEx),
    )
    errors = []
    for name, fn in attempts:
        try:
            visio_app = fn("Visio.Application")
            time.sleep(1)
            visio_app.Visible = True
            return visio_app
        except Exception as e:
            errors.append(f"{name}: {e}")
    raise RuntimeError(
        "Failed to initialize Visio.Application after multiple attempts: "
        + "; ".join(errors)
    )


def close_visio_app():
    """Properly close the Visio Application."""
    global visio_app, open_documents

    for _path, doc in list(open_documents.items()):
        try:
            doc.Close()
        except Exception:
            pass
    open_documents = {}

    if visio_app:
        try:
            visio_app.Quit()
        except Exception:
            pass
        visio_app = None


# Document lifecycle helpers (raise on failure; tools convert to strings).

def _create_document(template_path: Optional[str], save_path: str):
    app = get_visio_app()
    if template_path and os.path.exists(template_path):
        doc = app.Documents.Add(template_path)
    else:
        doc = app.Documents.Add("")
    time.sleep(1)
    doc.SaveAs(save_path)
    open_documents[_normalize_path(save_path)] = doc
    return doc


def _open_document(file_path: str):
    app = get_visio_app()
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Visio file does not exist: {file_path}")
    doc = app.Documents.Open(file_path)
    open_documents[_normalize_path(file_path)] = doc
    return doc


def _ensure_document_open(file_path: str, create_if_missing: bool = False):
    """Return the live COM Document for file_path, opening or creating as needed.

    Validates that any cached handle is still alive before reusing it.
    """
    key = _normalize_path(file_path)
    cached = open_documents.get(key)
    if cached is not None:
        try:
            _ = cached.Name  # liveness probe
            return cached
        except Exception:
            del open_documents[key]

    if os.path.exists(file_path):
        return _open_document(file_path)
    if create_if_missing:
        return _create_document(None, file_path)
    raise FileNotFoundError(f"Visio file does not exist: {file_path}")


def _find_shape(page, shape_id: int):
    for shape in page.Shapes:
        if shape.ID == shape_id:
            return shape
    return None


# ----------------------------- tools -----------------------------------------

@mcp.tool()
async def create_visio_file(template_path: Optional[str] = None, save_path: Optional[str] = None) -> str:
    """Create a new Visio file.

    Args:
        template_path: Path to the Visio template file (.vstx, .vst, etc.) to use.
                      If not provided, a default template will be used.
        save_path: Path where the file should be saved. If not provided,
                  it will be saved in the user's Documents folder with a default name.

    Returns:
        The path to the created Visio file.
    """
    _ensure_com_initialized()
    try:
        if not save_path:
            filename = f"New_Diagram_{int(time.time())}.vsdx"
            save_path = os.path.join(DEFAULT_SAVE_PATH, filename)
        elif os.path.dirname(save_path) == "":
            save_path = os.path.join(DEFAULT_SAVE_PATH, save_path)

        os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
        _create_document(template_path, save_path)
        return f"Visio file created successfully at: {save_path}"
    except Exception as e:
        return f"Error creating Visio file: {e}"


@mcp.tool()
async def open_visio_file(file_path: str) -> str:
    """Open an existing Visio file.

    Args:
        file_path: Path to the Visio file to open.

    Returns:
        Result message indicating success or failure.
    """
    _ensure_com_initialized()
    try:
        key = _normalize_path(file_path)
        cached = open_documents.get(key)
        if cached is not None:
            try:
                _ = cached.Name
                return f"Visio file is already open: {file_path}"
            except Exception:
                del open_documents[key]
        _open_document(file_path)
        return f"Visio file opened successfully: {file_path}"
    except FileNotFoundError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error opening Visio file: {e}"


@mcp.tool()
async def add_shape(file_path: str, shape_type: str, x: float, y: float,
                    width: Optional[float] = 1.0, height: Optional[float] = 1.0) -> str:
    """Add a shape to an existing Visio document.

    Args:
        file_path: Path to the Visio file.
        shape_type: Type of shape to add (e.g., "Rectangle", "Circle", "Line", etc.).
        x: X-coordinate for the shape.
        y: Y-coordinate for the shape.
        width: Width of the shape (default: 1.0).
        height: Height of the shape (default: 1.0).

    Returns:
        Result message indicating success or failure.
    """
    _ensure_com_initialized()
    try:
        app = get_visio_app()
        doc = _ensure_document_open(file_path, create_if_missing=True)
        page = app.ActivePage

        kind = shape_type.lower()
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
        doc.Save()
        return f"Shape '{shape_type}' added to the Visio file at ({x}, {y}) with ID {shape.ID}"
    except Exception as e:
        return f"Error adding shape to Visio file: {e}"


@mcp.tool()
async def connect_shapes(file_path: str, shape1_id: int, shape2_id: int,
                         connector_type: Optional[str] = "Dynamic") -> str:
    """Connect two shapes in a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape1_id: ID of the first shape.
        shape2_id: ID of the second shape.
        connector_type: "Dynamic" (default) or "Straight". "Curved" is accepted
                       but currently routes as Dynamic; true curved routing is
                       planned for a later phase.

    Returns:
        Result message indicating success or failure.
    """
    _ensure_com_initialized()
    try:
        app = get_visio_app()
        doc = _ensure_document_open(file_path)
        page = app.ActivePage

        shape1 = _find_shape(page, shape1_id)
        shape2 = _find_shape(page, shape2_id)
        if shape1 is None or shape2 is None:
            return f"Error: Could not find shapes with IDs {shape1_id} and {shape2_id}"

        connector = page.Drop(app.ConnectorToolDataObject, 0.0, 0.0)

        # ShapeRouteStyle on the connector controls layout routing.
        # 2 = visLORouteStraight; default leaves Visio's dynamic right-angle routing.
        if (connector_type or "").lower() == "straight":
            try:
                connector.Cells("ShapeRouteStyle").Formula = "2"
            except Exception:
                pass  # ShapeSheet cell unavailable on some connector masters

        # Dynamic shape-glue: gluing an endpoint to PinX tells Visio to pick the
        # best connection point automatically as the shapes move.
        connector.Cells("BeginX").GlueTo(shape1.Cells("PinX"))
        connector.Cells("EndX").GlueTo(shape2.Cells("PinX"))

        doc.Save()
        return f"Shapes {shape1_id} and {shape2_id} connected successfully with {connector_type} connector"
    except FileNotFoundError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error connecting shapes: {e}"


@mcp.tool()
async def add_text(file_path: str, shape_id: int, text: str) -> str:
    """Add text to a shape in a Visio document.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to add text to.
        text: Text to add to the shape.

    Returns:
        Result message indicating success or failure.
    """
    _ensure_com_initialized()
    try:
        app = get_visio_app()
        doc = _ensure_document_open(file_path)
        page = app.ActivePage

        target = _find_shape(page, shape_id)
        if target is None:
            return f"Error: Could not find shape with ID {shape_id}"

        target.Text = text
        doc.Save()
        return f"Text added to shape {shape_id} successfully"
    except FileNotFoundError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error adding text to shape: {e}"


@mcp.tool()
async def list_shapes(file_path: str) -> str:
    """List all shapes in a Visio document.

    Args:
        file_path: Path to the Visio file.

    Returns:
        JSON string containing information about all shapes in the document.
    """
    _ensure_com_initialized()
    try:
        app = get_visio_app()
        _ = _ensure_document_open(file_path)
        page = app.ActivePage

        shapes_info = []
        for shape in page.Shapes:
            shapes_info.append({
                "ID": shape.ID,
                "Name": shape.Name,
                "Text": shape.Text,
                "Type": shape.Type,
                "Position": {
                    "X": shape.Cells("PinX").Result(""),
                    "Y": shape.Cells("PinY").Result(""),
                },
                "Size": {
                    "Width": shape.Cells("Width").Result(""),
                    "Height": shape.Cells("Height").Result(""),
                },
            })
        return json.dumps(shapes_info, indent=2)
    except FileNotFoundError as e:
        return f"Error: {e}"
    except Exception as e:
        return f"Error listing shapes: {e}"


@mcp.tool()
async def close_document(file_path: str, save_changes: Optional[bool] = True) -> str:
    """Close a Visio document.

    Args:
        file_path: Path to the Visio file.
        save_changes: Whether to save changes before closing (default: True).

    Returns:
        Result message indicating success or failure.
    """
    _ensure_com_initialized()
    try:
        key = _normalize_path(file_path)
        if key not in open_documents:
            return f"Document {file_path} is not currently open"

        doc = open_documents[key]
        if save_changes:
            doc.Save()
        doc.Close()
        del open_documents[key]
        return f"Document {file_path} closed successfully"
    except Exception as e:
        return f"Error closing document: {e}"


atexit.register(close_visio_app)


def main():
    """Entry point for the MCP server."""
    if not check_visio_installed():
        sys.stderr.write("Microsoft Visio is not installed. This MCP server requires Visio to function.\n")
        sys.exit(1)

    try:
        _ = get_visio_app()
        mcp.run(transport="stdio")
    except Exception as e:
        sys.stderr.write(f"Error initializing Visio MCP Server: {e}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
