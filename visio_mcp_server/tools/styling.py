"""
Shape styling tools: fill, line, text format.

Every Visio shape exposes a ShapeSheet — a grid of named cells that drive
appearance. These tools write to the relevant cells directly. Cell names
referenced here:

- Fill section:  FillForegnd, FillPattern
- Line section:  LineColor, LineWeight, LinePattern
- Character section (text style): Char.Font, Char.Size, Char.Color, Char.Style
- Paragraph section (text alignment): Para.HorzAlign

Char.Style is a bit-mask: 1=bold, 2=italic, 4=underline. We read the
current value and toggle only the bits the caller specified, so partial
updates don't clobber other styles.
"""

from __future__ import annotations

from typing import Optional

from ..com.document import ensure_document_open, find_shape_on_page
from ..com.undo import undo_scope
from ..errors import InvalidArgument, ShapeNotFound, envelope
from ..server_instance import mcp

_ALIGN_MAP = {"left": 0, "center": 1, "right": 2, "justify": 3}

_CHAR_STYLE_BOLD = 1
_CHAR_STYLE_ITALIC = 2
_CHAR_STYLE_UNDERLINE = 4

_NAMED_COLORS = {
    "black": "RGB(0,0,0)",
    "white": "RGB(255,255,255)",
    "red": "RGB(255,0,0)",
    "green": "RGB(0,128,0)",
    "blue": "RGB(0,0,255)",
    "yellow": "RGB(255,255,0)",
    "cyan": "RGB(0,255,255)",
    "magenta": "RGB(255,0,255)",
    "gray": "RGB(128,128,128)",
    "grey": "RGB(128,128,128)",
    "orange": "RGB(255,165,0)",
    "purple": "RGB(128,0,128)",
    "pink": "RGB(255,192,203)",
    "brown": "RGB(165,42,42)",
}


def _parse_color(value: str) -> str:
    """Convert a user-supplied color into a Visio `RGB(r,g,b)` formula string.

    Accepts:
      - `#RRGGBB` or `#RGB`
      - `rgb(r,g,b)` / `RGB(r,g,b)`
      - Named colors from `_NAMED_COLORS`
    """
    v = value.strip()
    upper = v.upper()
    if upper.startswith("RGB(") and upper.endswith(")"):
        return upper

    if v.startswith("#"):
        hex_part = v[1:]
        if len(hex_part) == 3:
            r, g, b = (int(c * 2, 16) for c in hex_part)
        elif len(hex_part) == 6:
            r = int(hex_part[0:2], 16)
            g = int(hex_part[2:4], 16)
            b = int(hex_part[4:6], 16)
        else:
            raise InvalidArgument(f"Invalid hex color {value!r}; expected #RGB or #RRGGBB")
        return f"RGB({r},{g},{b})"

    if v.lower() in _NAMED_COLORS:
        return _NAMED_COLORS[v.lower()]

    raise InvalidArgument(
        f"Could not parse color {value!r}; use #RRGGBB, RGB(r,g,b), or a named color",
        details={"supported_names": sorted(_NAMED_COLORS.keys())},
    )


def _resolve_target_shape(file_path: str, shape_id: int, page_name: Optional[str]):
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    shape = find_shape_on_page(page, shape_id)
    if shape is None:
        raise ShapeNotFound(
            f"Could not find shape with ID {shape_id} on page '{page.Name}'",
            details={"shape_id": shape_id, "page_name": page.Name},
        )
    return handle, page, shape


@mcp.tool()
@envelope("set_shape_fill")
async def set_shape_fill(file_path: str, shape_id: int, color: str,
                         pattern: Optional[int] = 1,
                         page_name: Optional[str] = None) -> dict:
    """Set a shape's fill color and pattern.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to style.
        color: Fill color. Accepts `#RRGGBB`, `RGB(r,g,b)`, or a named color
              (red, blue, green, ...).
        pattern: Fill pattern code. 0 = none, 1 = solid (default),
                2-40 = various hatches/gradients.
        page_name: Page the shape lives on. Defaults to the active page.

    Returns:
        {"shape_id": int, "color": str, "pattern": int, "page_name": str}
    """
    handle, page, shape = _resolve_target_shape(file_path, shape_id, page_name)
    color_formula = _parse_color(color)

    with undo_scope("Set shape fill"):
        shape.Cells("FillForegnd").Formula = color_formula
        if pattern is not None:
            shape.Cells("FillPattern").Formula = str(int(pattern))

    handle.save()
    return {
        "shape_id": shape_id,
        "color": color_formula,
        "pattern": int(pattern) if pattern is not None else None,
        "page_name": page.Name,
    }


@mcp.tool()
@envelope("set_shape_line")
async def set_shape_line(file_path: str, shape_id: int,
                         color: Optional[str] = None,
                         weight: Optional[float] = None,
                         pattern: Optional[int] = None,
                         page_name: Optional[str] = None) -> dict:
    """Set a shape's outline color, weight, and dash pattern.

    All parameters are optional; pass only the ones you want to change.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape.
        color: Outline color (`#RRGGBB`, `RGB(r,g,b)`, or named).
        weight: Outline thickness in points.
        pattern: Dash pattern code. 0 = no line, 1 = solid,
                2-23 = various dashes.
        page_name: Page the shape lives on. Defaults to the active page.

    Returns:
        {"shape_id": int, "color": str|null, "weight_pt": float|null,
         "pattern": int|null, "page_name": str}
    """
    if color is None and weight is None and pattern is None:
        raise InvalidArgument(
            "set_shape_line called with no changes; pass at least one of color, weight, pattern"
        )

    handle, page, shape = _resolve_target_shape(file_path, shape_id, page_name)
    color_formula = _parse_color(color) if color is not None else None

    with undo_scope("Set shape line"):
        if color_formula is not None:
            shape.Cells("LineColor").Formula = color_formula
        if weight is not None:
            shape.Cells("LineWeight").Formula = f"{float(weight)} pt"
        if pattern is not None:
            shape.Cells("LinePattern").Formula = str(int(pattern))

    handle.save()
    return {
        "shape_id": shape_id,
        "color": color_formula,
        "weight_pt": float(weight) if weight is not None else None,
        "pattern": int(pattern) if pattern is not None else None,
        "page_name": page.Name,
    }


@mcp.tool()
@envelope("set_shape_text_format")
async def set_shape_text_format(file_path: str, shape_id: int,
                                font: Optional[str] = None,
                                size: Optional[float] = None,
                                color: Optional[str] = None,
                                bold: Optional[bool] = None,
                                italic: Optional[bool] = None,
                                underline: Optional[bool] = None,
                                align: Optional[str] = None,
                                page_name: Optional[str] = None) -> dict:
    """Set text formatting on a shape.

    All parameters are optional; only the specified attributes change.
    Bold/italic/underline are bit flags on Char.Style — we read the current
    value and toggle only the bits the caller asked about.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape.
        font: Font family name (e.g. "Arial", "Calibri"). Must already be
              available in the document's Fonts collection.
        size: Font size in points.
        color: Text color (`#RRGGBB`, `RGB(r,g,b)`, or named).
        bold, italic, underline: True to enable, False to disable.
        align: Horizontal alignment — "left", "center", "right", or "justify".
        page_name: Page the shape lives on. Defaults to the active page.

    Returns:
        {"shape_id": int, "applied": {...}, "page_name": str}
    """
    if all(v is None for v in (font, size, color, bold, italic, underline, align)):
        raise InvalidArgument(
            "set_shape_text_format called with no changes; pass at least one attribute"
        )

    handle, page, shape = _resolve_target_shape(file_path, shape_id, page_name)
    applied: dict = {}

    with undo_scope("Set shape text format"):
        if font is not None:
            try:
                font_id = int(handle.com_doc.Fonts.ItemU(font).ID)
            except Exception:
                raise InvalidArgument(
                    f"Font {font!r} not available in this document",
                    details={"font": font},
                )
            shape.Cells("Char.Font").Formula = str(font_id)
            applied["font"] = font

        if size is not None:
            shape.Cells("Char.Size").Formula = f"{float(size)} pt"
            applied["size_pt"] = float(size)

        if color is not None:
            formula = _parse_color(color)
            shape.Cells("Char.Color").Formula = formula
            applied["color"] = formula

        if any(b is not None for b in (bold, italic, underline)):
            current = int(shape.Cells("Char.Style").Result(""))
            new_style = current
            if bold is not None:
                new_style = (new_style | _CHAR_STYLE_BOLD) if bold else (new_style & ~_CHAR_STYLE_BOLD)
                applied["bold"] = bool(bold)
            if italic is not None:
                new_style = (new_style | _CHAR_STYLE_ITALIC) if italic else (new_style & ~_CHAR_STYLE_ITALIC)
                applied["italic"] = bool(italic)
            if underline is not None:
                new_style = (new_style | _CHAR_STYLE_UNDERLINE) if underline else (new_style & ~_CHAR_STYLE_UNDERLINE)
                applied["underline"] = bool(underline)
            shape.Cells("Char.Style").Formula = str(new_style)

        if align is not None:
            key = align.lower().strip()
            if key not in _ALIGN_MAP:
                raise InvalidArgument(
                    f"Unknown align {align!r}; expected one of {list(_ALIGN_MAP.keys())}",
                )
            shape.Cells("Para.HorzAlign").Formula = str(_ALIGN_MAP[key])
            applied["align"] = key

    handle.save()
    return {"shape_id": shape_id, "applied": applied, "page_name": page.Name}
