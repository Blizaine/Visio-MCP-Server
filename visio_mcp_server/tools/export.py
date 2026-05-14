"""
Export tools: single-page image export, multi-page PDF export.

Visio exposes two distinct export paths:
- `Page.Export(filename)` writes a single page as a raster or vector image.
  Format is inferred from the filename extension. Visio recognises .png,
  .jpg/.jpeg, .gif, .bmp, .tif/.tiff, .svg, .emf, and .wmf.
- `Document.ExportAsFixedFormat(format, ...)` writes one or more pages as
  PDF or XPS.
"""

from __future__ import annotations

import os
from typing import Optional

from ..com.document import ensure_document_open
from ..errors import InvalidArgument, envelope
from ..server_instance import mcp

_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff",
               ".svg", ".emf", ".wmf"}

# VisFixedFormatTypes / VisDocExIntent / VisPrintOutRange constants
_FIXED_FORMAT_PDF = 1
_INTENT_PRINT = 1     # higher fidelity; the screen variant (2) compresses harder
_RANGE_ALL = 0
_RANGE_CURRENT = 1
_RANGE_FROMTO = 2


@mcp.tool()
@envelope("export_page")
async def export_page(file_path: str, output_path: str,
                      page_name: Optional[str] = None) -> dict:
    """Export a single Visio page to an image file.

    Format is inferred from the output filename extension. Supported:
    .png, .jpg/.jpeg, .gif, .bmp, .tif/.tiff, .svg, .emf, .wmf.

    Args:
        file_path: Path to the source Visio file.
        output_path: Destination path. Extension drives the format.
        page_name: Page to export. Defaults to the active page.

    Returns:
        {"output_path": str, "page_name": str, "format": str, "bytes": int}
    """
    ext = os.path.splitext(output_path)[1].lower()
    if ext not in _IMAGE_EXTS:
        raise InvalidArgument(
            f"Unsupported image extension {ext!r}; expected one of {sorted(_IMAGE_EXTS)}",
            details={"output_path": output_path, "supported": sorted(_IMAGE_EXTS)},
        )

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    page.Export(output_path)

    size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
    return {
        "output_path": output_path,
        "page_name": page.Name,
        "format": ext.lstrip("."),
        "bytes": size,
    }


@mcp.tool()
@envelope("export_pdf")
async def export_pdf(file_path: str, output_path: str,
                     page_range: Optional[str] = None) -> dict:
    """Export pages of a Visio document to PDF.

    Args:
        file_path: Path to the source Visio file.
        output_path: Destination PDF path.
        page_range: Which pages to include. Accepts:
                   - None or "all" (default): all pages
                   - "current": just the currently-active page
                   - "N": just page N (1-based)
                   - "N-M": pages N through M inclusive

    Returns:
        {"output_path": str, "page_range": str, "bytes": int}
    """
    if not output_path.lower().endswith(".pdf"):
        raise InvalidArgument(
            f"Output path must end in .pdf; got {output_path!r}",
            details={"output_path": output_path},
        )

    handle = ensure_document_open(file_path)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    rng = (page_range or "all").strip().lower()
    if rng == "all":
        handle.com_doc.ExportAsFixedFormat(
            _FIXED_FORMAT_PDF, output_path, _INTENT_PRINT, _RANGE_ALL,
        )
        normalized = "all"
    elif rng == "current":
        handle.com_doc.ExportAsFixedFormat(
            _FIXED_FORMAT_PDF, output_path, _INTENT_PRINT, _RANGE_CURRENT,
        )
        normalized = "current"
    elif "-" in rng:
        try:
            start_s, end_s = rng.split("-", 1)
            start, end = int(start_s), int(end_s)
        except ValueError:
            raise InvalidArgument(f"Could not parse page_range {page_range!r}; expected 'N-M'")
        if start < 1 or end < start:
            raise InvalidArgument(f"Invalid page range {page_range!r}; expected 1 <= start <= end")
        handle.com_doc.ExportAsFixedFormat(
            _FIXED_FORMAT_PDF, output_path, _INTENT_PRINT, _RANGE_FROMTO, start, end,
        )
        normalized = f"{start}-{end}"
    else:
        try:
            page_num = int(rng)
        except ValueError:
            raise InvalidArgument(f"Could not parse page_range {page_range!r}")
        if page_num < 1:
            raise InvalidArgument(f"Page numbers are 1-based; got {page_num}")
        handle.com_doc.ExportAsFixedFormat(
            _FIXED_FORMAT_PDF, output_path, _INTENT_PRINT, _RANGE_FROMTO, page_num, page_num,
        )
        normalized = str(page_num)

    size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
    return {"output_path": output_path, "page_range": normalized, "bytes": size}
