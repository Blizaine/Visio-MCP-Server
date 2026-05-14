"""
Page-level tools: list / add / delete / set_active / duplicate.
"""

from __future__ import annotations

from typing import Optional

from ..com.app import get_visio_app
from ..com.document import ensure_document_open
from ..com.undo import undo_scope
from ..errors import InvalidArgument, PageNotFound, envelope
from ..server_instance import mcp


@mcp.tool()
@envelope("list_pages")
async def list_pages(file_path: str) -> dict:
    """List all pages in a Visio document.

    Args:
        file_path: Path to the Visio file.

    Returns:
        {"pages": [{"index": int, "name": str, "width": float, "height": float,
                    "background": bool}, ...]}
    """
    handle = ensure_document_open(file_path)
    pages = []
    # Pages.Item uses 1-based indexing.
    for i in range(1, handle.com_doc.Pages.Count + 1):
        page = handle.com_doc.Pages.Item(i)
        # PageSheet cells expose page dimensions in the document's units.
        # Result("in") forces inches so callers get a stable unit.
        width = float(page.PageSheet.Cells("PageWidth").Result("in"))
        height = float(page.PageSheet.Cells("PageHeight").Result("in"))
        pages.append({
            "index": i,
            "name": page.Name,
            "width": width,
            "height": height,
            "background": bool(page.Background),
        })
    return {"pages": pages}


@mcp.tool()
@envelope("add_page")
async def add_page(file_path: str, name: Optional[str] = None,
                   width: Optional[float] = None, height: Optional[float] = None,
                   background: Optional[bool] = False) -> dict:
    """Add a new page to a Visio document.

    Args:
        file_path: Path to the Visio file.
        name: Name for the new page. If None, Visio auto-names it
              (typically "Page-N" where N is the next index).
        width, height: Page size in inches. If either is omitted, the
                      document's default page size is used.
        background: True for a background page (used as a backdrop for
                   other pages), False (default) for a normal foreground
                   page.

    Returns:
        {"name": str, "index": int}
    """
    handle = ensure_document_open(file_path)

    with undo_scope(f"Add page{f' {name}' if name else ''}"):
        page = handle.com_doc.Pages.Add()
        if background:
            page.Background = True
        if name is not None:
            page.Name = name
        if width is not None:
            page.PageSheet.Cells("PageWidth").Formula = f"{float(width)} in"
        if height is not None:
            page.PageSheet.Cells("PageHeight").Formula = f"{float(height)} in"

    handle.save()
    return {"name": page.Name, "index": int(page.Index)}


@mcp.tool()
@envelope("delete_page")
async def delete_page(file_path: str, page_name: str) -> dict:
    """Delete a page from a Visio document.

    Refuses to delete the last remaining page (Visio requires at least one).

    Args:
        file_path: Path to the Visio file.
        page_name: Name of the page to delete.

    Returns:
        {"deleted_name": str}
    """
    handle = ensure_document_open(file_path)

    if handle.com_doc.Pages.Count <= 1:
        raise InvalidArgument(
            "Cannot delete the last remaining page; a Visio document must have at least one page",
            details={"document": handle.path, "page_count": int(handle.com_doc.Pages.Count)},
        )

    page = handle.get_page(page_name)
    deleted = page.Name  # capture before delete

    with undo_scope(f"Delete page {deleted}"):
        # Page.Delete(int): 0 = delete shapes too; 1 = delete just the page (default behavior).
        page.Delete(0)

    handle.save()
    return {"deleted_name": deleted}


@mcp.tool()
@envelope("set_active_page")
async def set_active_page(file_path: str, page_name: str) -> dict:
    """Make a specific page the active one in Visio's window.

    This affects which page subsequent default-targeted tool calls
    (page_name=None) operate on, and which page the user sees.

    Args:
        file_path: Path to the Visio file.
        page_name: Name of the page to activate.

    Returns:
        {"active_page_name": str}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    app = get_visio_app()

    # Find the window showing our document, activate it, then switch its page.
    target_window = None
    for i in range(1, app.Windows.Count + 1):
        win = app.Windows.Item(i)
        try:
            win_doc = win.Document
        except Exception:
            continue
        if win_doc is not None and win_doc.FullName == handle.com_doc.FullName:
            target_window = win
            break

    if target_window is None:
        # Fall back to the application's active window; Visio will switch the
        # underlying document context as needed.
        target_window = app.ActiveWindow

    target_window.Activate()
    target_window.Page = page
    return {"active_page_name": page.Name}


@mcp.tool()
@envelope("duplicate_page")
async def duplicate_page(file_path: str, source_page_name: str,
                         new_name: Optional[str] = None) -> dict:
    """Duplicate an existing page (shapes, dimensions, properties).

    Args:
        file_path: Path to the Visio file.
        source_page_name: Name of the page to copy.
        new_name: Optional name for the duplicate. If omitted, Visio
                  auto-names it (e.g. "Page-N").

    Returns:
        {"source_name": str, "new_name": str, "new_index": int}
    """
    handle = ensure_document_open(file_path)
    source = handle.get_page(source_page_name)

    with undo_scope(f"Duplicate {source_page_name}"):
        copy = source.Duplicate()
        if new_name is not None:
            copy.Name = new_name

    handle.save()
    return {
        "source_name": source.Name,
        "new_name": copy.Name,
        "new_index": int(copy.Index),
    }
