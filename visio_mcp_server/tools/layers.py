"""
Layer tools — segregate shapes onto named layers per page.

In AV system drawings, layers commonly carry signal types: "Audio
Routing", "Video Routing", "Control Wiring", "Network". Engineers
toggle layer visibility / print / lock during review.

Visio's Layer model:
- Each Page has its own Layers collection (not document-wide).
- A shape can belong to zero, one, or many layers on its page.
- Layer properties (Visible, Print, Active, Lock, Color) live in the
  page's ShapeSheet under section visSectionLayer (5) and are accessed
  via `Layer.CellsU(<name>)`.
"""

from __future__ import annotations

from typing import Optional

from ..com.document import ensure_document_open, find_shape_on_page
from ..com.undo import undo_scope
from ..errors import InvalidArgument, ShapeNotFound, envelope
from ..server_instance import mcp
from .batch import _batch_context
from .styling import parse_color

# Layer cells live in section 5 (visSectionLayer) of the page's ShapeSheet.
# Each layer occupies one row, addressed by `layer.Index - 1`. The columns
# below are the documented offsets within that section. Going through
# `Page.PageSheet.CellsSRC(section, row, column)` is the lowest-level
# (and most portable) way to touch them — some Visio builds don't expose
# `Layer.CellsU(name)` consistently.
_VIS_SECTION_LAYER = 5
_LAYER_COL_VISIBLE = 1
_LAYER_COL_PRINT = 2
_LAYER_COL_ACTIVE = 3
_LAYER_COL_LOCK = 4
_LAYER_COL_COLOR = 8


def _layer_cell(layer, column: int):
    """Get the Cell at (visSectionLayer, layer.row, column) via the
    page's ShapeSheet. `layer.Index` is 1-based; the row is 0-based."""
    page = layer.Page
    return page.PageSheet.CellsSRC(_VIS_SECTION_LAYER, int(layer.Index) - 1, column)


def _read_layer(layer) -> dict:
    """Snapshot of a Layer's user-visible properties."""
    def _bool_cell(column, default=False):
        try:
            return bool(_layer_cell(layer, column).ResultIU)
        except Exception:
            return default

    color_formula = ""
    try:
        color_formula = str(_layer_cell(layer, _LAYER_COL_COLOR).FormulaU)
    except Exception:
        pass

    return {
        "name": str(layer.Name),
        "index": int(layer.Index),
        "visible": _bool_cell(_LAYER_COL_VISIBLE, True),
        "print": _bool_cell(_LAYER_COL_PRINT, True),
        "active": _bool_cell(_LAYER_COL_ACTIVE, False),
        "locked": _bool_cell(_LAYER_COL_LOCK, False),
        "color_formula": color_formula,
    }


def _find_layer(page, name: str):
    """Locate a layer on `page` by case-insensitive name. Returns the
    Layer COM object or None."""
    needle = name.lower()
    for i in range(1, page.Layers.Count + 1):
        layer = page.Layers.Item(i)
        if str(layer.Name).lower() == needle:
            return layer
    return None


def _apply_layer_properties(layer, *, visible=None, printable=None,
                            locked=None, color=None) -> dict:
    """Write the subset of properties the caller passed. Returns applied."""
    applied: dict = {}
    if visible is not None:
        _layer_cell(layer, _LAYER_COL_VISIBLE).FormulaU = "1" if visible else "0"
        applied["visible"] = bool(visible)
    if printable is not None:
        _layer_cell(layer, _LAYER_COL_PRINT).FormulaU = "1" if printable else "0"
        applied["print"] = bool(printable)
    if locked is not None:
        _layer_cell(layer, _LAYER_COL_LOCK).FormulaU = "1" if locked else "0"
        applied["locked"] = bool(locked)
    if color is not None:
        formula = parse_color(color)
        _layer_cell(layer, _LAYER_COL_COLOR).FormulaU = formula
        applied["color"] = formula
    return applied


# --------------------------------------------------------------------- tools

@mcp.tool()
@envelope("list_layers")
async def list_layers(file_path: str, page_name: Optional[str] = None) -> dict:
    """List the layers on a page, with their visible/print/lock state.

    Args:
        file_path: Path to the Visio file.
        page_name: Page to inspect. Defaults to active page.

    Returns:
        {"page_name": str, "count": int,
         "layers": [{"name", "index", "visible", "print", "active",
                     "locked", "color_formula"}]}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    layers = []
    for i in range(1, page.Layers.Count + 1):
        layers.append(_read_layer(page.Layers.Item(i)))
    return {"page_name": page.Name, "count": len(layers), "layers": layers}


@mcp.tool()
@envelope("add_layer")
async def add_layer(file_path: str, name: str,
                    page_name: Optional[str] = None,
                    visible: Optional[bool] = None,
                    printable: Optional[bool] = None,
                    locked: Optional[bool] = None,
                    color: Optional[str] = None) -> dict:
    """Add a layer to a page.

    Visible / printable / locked default to Visio's defaults if omitted
    (typically visible=True, printable=True, locked=False).

    Args:
        file_path: Path to the Visio file.
        name: Layer name (must be unique on the page).
        page_name: Target page. Defaults to active page.
        visible, printable, locked: Initial state for the toggles.
        color: Initial layer color (`#RRGGBB`, `RGB(...)`, or named).

    Returns:
        {"page_name": str, "layer": {...full layer snapshot...}}
    """
    if not name or not name.strip():
        raise InvalidArgument("add_layer: name must be non-empty")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    if _find_layer(page, name) is not None:
        raise InvalidArgument(
            f"Layer {name!r} already exists on page '{page.Name}'",
            details={"name": name, "page_name": page.Name},
        )

    with undo_scope(f"Add layer {name}"):
        layer = page.Layers.Add(name)
        _apply_layer_properties(layer, visible=visible, printable=printable,
                                locked=locked, color=color)

    return {"page_name": page.Name, "layer": _read_layer(layer)}


@mcp.tool()
@envelope("delete_layer")
async def delete_layer(file_path: str, name: str,
                       page_name: Optional[str] = None,
                       delete_shapes: bool = False) -> dict:
    """Delete a layer.

    By default, shapes on the layer are kept (they just lose their
    layer membership). Set `delete_shapes=True` to remove the shapes
    too — useful when a whole signal-flow layer is being dropped.

    Args:
        file_path: Path to the Visio file.
        name: Layer name (case-insensitive).
        page_name: Target page. Defaults to active page.
        delete_shapes: If True, also delete every shape on the layer.

    Returns:
        {"page_name": str, "deleted_name": str, "deleted_shapes": bool}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    layer = _find_layer(page, name)
    if layer is None:
        raise InvalidArgument(
            f"Layer {name!r} not found on page '{page.Name}'",
            details={"name": name, "page_name": page.Name},
        )

    deleted_name = str(layer.Name)
    with undo_scope(f"Delete layer {deleted_name}"):
        layer.Delete(1 if delete_shapes else 0)

    return {
        "page_name": page.Name,
        "deleted_name": deleted_name,
        "deleted_shapes": bool(delete_shapes),
    }


@mcp.tool()
@envelope("set_layer_properties")
async def set_layer_properties(file_path: str, name: str,
                               page_name: Optional[str] = None,
                               visible: Optional[bool] = None,
                               printable: Optional[bool] = None,
                               locked: Optional[bool] = None,
                               color: Optional[str] = None) -> dict:
    """Toggle visibility / print / lock / color on an existing layer.

    All four toggles are optional; pass only what you want to change.
    Pattern matches `set_shape_line` etc. — partial updates supported.

    Args:
        file_path: Path to the Visio file.
        name: Layer name (case-insensitive).
        page_name: Target page. Defaults to active page.
        visible: Toggle visibility.
        printable: Toggle print.
        locked: Toggle lock (locked shapes can't be selected/edited).
        color: Layer color (`#RRGGBB`, `RGB(...)`, or named).

    Returns:
        {"page_name": str, "name": str, "applied": {...}}
    """
    if all(v is None for v in (visible, printable, locked, color)):
        raise InvalidArgument(
            "set_layer_properties: pass at least one of visible, printable, locked, color"
        )

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    layer = _find_layer(page, name)
    if layer is None:
        raise InvalidArgument(
            f"Layer {name!r} not found on page '{page.Name}'",
            details={"name": name, "page_name": page.Name},
        )

    with undo_scope(f"Set layer {name} properties"):
        applied = _apply_layer_properties(
            layer, visible=visible, printable=printable,
            locked=locked, color=color,
        )

    return {"page_name": page.Name, "name": str(layer.Name), "applied": applied}


@mcp.tool()
@envelope("set_shapes_layers")
async def set_shapes_layers(file_path: str, assignments: list,
                            page_name: Optional[str] = None) -> dict:
    """Assign shapes to layers in a single batch.

    Each entry in `assignments`:
        {"shape_id": int, "layers": ["Layer A", "Layer B", ...]}

    REPLACE semantics: the shape's layer membership is replaced with
    exactly the listed layers. Pass an empty list to remove the shape
    from all layers.

    Pre-validates that every named layer exists on the page; if any
    layer name (across any assignment) is missing, the batch aborts
    before any writes.

    Args:
        file_path: Path to the Visio file.
        assignments: List of `{shape_id, layers}` assignments.
        page_name: Page the shapes live on. Defaults to active page.

    Returns:
        {"page_name": str, "count": int,
         "results": [{"shape_id", "layers": [str, ...]}]}
    """
    if not assignments:
        raise InvalidArgument("set_shapes_layers called with empty assignments list")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    # Build a {lower(name): Layer} map up front so we fail fast on a
    # missing layer name. We use Layer.Add(shape, 0) / Layer.Remove(shape, 0)
    # rather than Shape.AddLayer / Shape.RemoveLayer because the latter
    # don't dispatch on every Visio build.
    layer_by_name: dict = {}
    all_layers: list = []
    for i in range(1, page.Layers.Count + 1):
        layer = page.Layers.Item(i)
        layer_by_name[str(layer.Name).lower()] = layer
        all_layers.append(layer)

    # Resolve every assignment before mutating anything.
    resolved = []
    for j, a in enumerate(assignments):
        try:
            shape_id = int(a["shape_id"])
        except (KeyError, TypeError, ValueError):
            raise InvalidArgument(
                f"assignments[{j}]: missing or invalid shape_id",
                details={"index": j, "spec": a},
            )
        layers = a.get("layers", [])
        if not isinstance(layers, list):
            raise InvalidArgument(
                f"assignments[{j}]: 'layers' must be a list of names (got {type(layers).__name__})",
                details={"index": j, "shape_id": shape_id},
            )

        shape = find_shape_on_page(page, shape_id)
        if shape is None:
            raise ShapeNotFound(
                f"assignments[{j}]: shape ID {shape_id} not on page '{page.Name}'",
                details={"index": j, "shape_id": shape_id, "page_name": page.Name},
            )

        target_layers = []
        for layer_name in layers:
            key = str(layer_name).lower()
            if key not in layer_by_name:
                raise InvalidArgument(
                    f"assignments[{j}]: layer {layer_name!r} not on page '{page.Name}'",
                    details={"index": j, "shape_id": shape_id,
                             "missing_layer": layer_name, "page_name": page.Name,
                             "available": [str(l.Name) for l in all_layers]},
                )
            target_layers.append(layer_by_name[key])

        resolved.append((shape_id, shape, target_layers, list(layers)))

    results = []
    with _batch_context(f"Assign layers on {len(resolved)} shapes"):
        for shape_id, shape, target_layers, layer_names in resolved:
            # Remove from every layer on the page first. Layer.Remove is a
            # no-op when the shape isn't on the layer, but some Visio
            # builds raise — swallow that.
            for layer in all_layers:
                try:
                    layer.Remove(shape, 0)  # 0 = don't preserve formatting
                except Exception:
                    pass
            for layer in target_layers:
                layer.Add(shape, 0)
            results.append({"shape_id": shape_id, "layers": layer_names})

    return {"page_name": page.Name, "count": len(results), "results": results}
