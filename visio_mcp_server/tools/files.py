"""
File-level tools: create, open, close.

Each tool body raises on failure and returns structured data on success;
the `@envelope` decorator converts that to the standard JSON shape.
"""

from __future__ import annotations

import os
import time
from typing import Optional

from ..com.document import (
    create_document,
    ensure_document_open,
    forget_document,
    get_open_document,
    open_document,
)
from ..com.undo import undo_scope
from ..errors import envelope
from ..server_instance import mcp

DEFAULT_SAVE_PATH = os.path.expandvars(r"%USERPROFILE%\Documents")


@mcp.tool()
@envelope("create_visio_file")
async def create_visio_file(template_path: Optional[str] = None, save_path: Optional[str] = None) -> dict:
    """Create a new Visio file.

    Args:
        template_path: Path to a Visio template (.vstx, .vst). Optional.
        save_path: Where to save the file. If a bare filename, saved under
                  the user's Documents folder. If omitted, an auto-named file
                  is saved there.

    Returns:
        {"path": str} — the path the file was saved to.
    """
    if not save_path:
        save_path = os.path.join(DEFAULT_SAVE_PATH, f"New_Diagram_{int(time.time())}.vsdx")
    elif os.path.dirname(save_path) == "":
        save_path = os.path.join(DEFAULT_SAVE_PATH, save_path)

    os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)

    with undo_scope("Create Visio document"):
        create_document(template_path, save_path)
    return {"path": save_path}


@mcp.tool()
@envelope("open_visio_file")
async def open_visio_file(file_path: str) -> dict:
    """Open an existing Visio file.

    Args:
        file_path: Path to the Visio file to open.

    Returns:
        {"path": str, "already_open": bool}
    """
    existing = get_open_document(file_path)
    if existing is not None:
        return {"path": file_path, "already_open": True}

    open_document(file_path)
    return {"path": file_path, "already_open": False}


@mcp.tool()
@envelope("close_document")
async def close_document(file_path: str, save_changes: Optional[bool] = True) -> dict:
    """Close a Visio document.

    Args:
        file_path: Path to the Visio file.
        save_changes: Whether to save changes before closing (default: True).

    Returns:
        {"path": str, "was_open": bool}
    """
    handle = get_open_document(file_path)
    if handle is None:
        return {"path": file_path, "was_open": False}

    handle.close(save_changes=bool(save_changes))
    forget_document(file_path)
    return {"path": file_path, "was_open": True}


@mcp.tool()
@envelope("save_document")
async def save_document(file_path: str) -> dict:
    """Persist any in-memory changes for an open Visio document.

    As of v2.3.0, mutating tools (add_shape, set_shape_fill, transform_shapes,
    etc.) no longer save automatically — that was a per-call latency cost
    and the dominant disk-write driver. Instead, call this when you want
    the on-disk file in sync with what's in Visio, or let `close_document`
    save on close.

    Args:
        file_path: Path to the Visio file. Must already be open.

    Returns:
        {"path": str}
    """
    handle = ensure_document_open(file_path)
    handle.save()
    return {"path": file_path}
