"""
Drop master shapes from stencils onto pages.

This is the payoff for the Phase 7 stencil index. The typical AV workflow:

    find_masters("Crestron DM-NVX-360")     # picks the right master
    drop_masters(items=[                    # places it on the page
        {"stencil": "Crestron", "master": "DM-NVX-360",
         "x": 2.0, "y": 5.0, "text": "ENC-01"},
        ...
    ])

Stencils are cached in their open Visio Document form for the lifetime of
the server, so dropping a dozen masters from one stencil costs one open
(~0.5s) plus N drops (~50ms each).

Property writes (the optional `data` field per item) target ShapeSheet
`Prop.<Name>` cells inherited from the master. Unknown property names are
silently skipped — for richer shape-data handling (discovery of available
fields, validation, find-by-data), see Phase 9 tools.
"""

from __future__ import annotations

from typing import Optional

from ..com.document import ensure_document_open
from ..com.stencils import get_or_open_stencil
from ..com.undo import undo_scope
from ..errors import InvalidArgument, envelope
from ..server_instance import mcp
from .batch import _batch_context


def _format_prop_value(value) -> Optional[str]:
    """Format a Python value as a Visio ShapeSheet formula. None means skip."""
    if value is None:
        return None
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return str(value)
    s = str(value).replace('"', '""')   # Visio escapes embedded quotes by doubling
    return f'"{s}"'


def _apply_data(shape, data: dict) -> dict:
    """Write `Prop.<name>` cells. Returns {name: applied_value} for what stuck."""
    applied: dict = {}
    for name, value in data.items():
        formula = _format_prop_value(value)
        if formula is None:
            continue
        try:
            shape.Cells(f"Prop.{name}").FormulaU = formula
        except Exception:
            # Property doesn't exist on this master, or its name is malformed.
            # Phase 9's set_shape_data will validate and report; here we
            # silently skip so a partial drop still completes.
            continue
        applied[name] = value
    return applied


def _drop_one(page, stencil_doc, spec: dict, index_for_errors: int):
    """Drop a single master per `spec`. Returns the resulting Shape."""
    master_name = spec.get("master")
    if not master_name:
        raise InvalidArgument(
            f"items[{index_for_errors}]: missing 'master' name",
            details={"index": index_for_errors, "spec": spec},
        )
    try:
        master = stencil_doc.Masters.ItemU(master_name)
    except Exception:
        raise InvalidArgument(
            f"items[{index_for_errors}]: master {master_name!r} not found in stencil "
            f"{spec.get('stencil')!r}. Use `list_masters` to see what's available.",
            details={"index": index_for_errors,
                     "stencil": spec.get("stencil"),
                     "master": master_name},
        )

    try:
        x = float(spec["x"])
        y = float(spec["y"])
    except (KeyError, TypeError, ValueError):
        raise InvalidArgument(
            f"items[{index_for_errors}]: 'x' and 'y' are required numbers (inches)",
            details={"index": index_for_errors, "spec": spec},
        )

    shape = page.Drop(master, x, y)

    width = spec.get("width")
    height = spec.get("height")
    if width is not None:
        shape.Cells("Width").Formula = f"{float(width)} in"
    if height is not None:
        shape.Cells("Height").Formula = f"{float(height)} in"

    text = spec.get("text")
    if text is not None:
        shape.Text = str(text)

    return shape


@mcp.tool()
@envelope("drop_master")
async def drop_master(
    file_path: str,
    stencil: str,
    master: str,
    x: float,
    y: float,
    page_name: Optional[str] = None,
    width: Optional[float] = None,
    height: Optional[float] = None,
    text: Optional[str] = None,
    data: Optional[dict] = None,
) -> dict:
    """Drop a single master shape from a stencil onto a page.

    For real Visio diagrams (AV system design, network topologies, org
    charts), this is the right tool — `add_shape` only draws generic
    rectangles/circles/lines. Use `find_masters` first to pick the right
    stencil + master if you don't already know.

    Prefer `drop_masters` (plural) when placing more than one shape;
    it opens each stencil once and disables screen redraws for the batch.

    Args:
        file_path: Path to the Visio file. Created if it doesn't exist.
        stencil: Stencil name (filename without extension) or full path.
                Must be in the stencil index (`reindex_stencils` if not).
        master: Master shape name within that stencil. Case-sensitive on
               Visio's side.
        x, y: Drop position in inches (where the master's pin lands).
        page_name: Target page. Defaults to the active page (or the
                  document's first page in multi-doc scenarios).
        width, height: Optional override of the master's default size.
        text: Optional initial text label on the dropped shape.
        data: Optional dict of `{prop_name: value}` to write to
              `Prop.<name>` cells inherited from the master. Unknown
              property names are silently skipped — use Phase 9 tools
              for property discovery and validation.

    Returns:
        {"shape_id": int, "stencil": str, "master": str, "x": float,
         "y": float, "page_name": str, "applied_data": dict|null}
    """
    handle = ensure_document_open(file_path, create_if_missing=True)
    page = handle.get_page(page_name)
    stencil_doc = get_or_open_stencil(stencil)

    spec = {"stencil": stencil, "master": master, "x": x, "y": y}
    if width is not None:
        spec["width"] = width
    if height is not None:
        spec["height"] = height
    if text is not None:
        spec["text"] = text

    with undo_scope(f"Drop {master}"):
        shape = _drop_one(page, stencil_doc, spec, 0)
        applied = _apply_data(shape, data) if data else {}

    return {
        "shape_id": int(shape.ID),
        "stencil": stencil,
        "master": master,
        "x": float(x),
        "y": float(y),
        "page_name": page.Name,
        "applied_data": applied if applied else None,
    }


@mcp.tool()
@envelope("drop_masters")
async def drop_masters(file_path: str, items: list,
                       page_name: Optional[str] = None) -> dict:
    """Drop multiple master shapes onto a page in a single batch.

    Strongly preferred over multiple `drop_master` calls — typically
    5-20x faster end-to-end because there's one model round-trip
    instead of N, each stencil opens at most once, and Visio's screen
    redraw is suspended for the duration.

    Each entry in `items`:

        {
          "stencil": str,                     # short name or full path
          "master": str,                      # master shape name
          "x": float, "y": float,             # drop position in inches
          "width"?: float, "height"?: float,  # optional size override
          "text"?: str,                       # optional shape text
          "data"?: {prop_name: value, ...}    # optional Prop.<name> writes
        }

    Args:
        file_path: Path to the Visio file. Created if it doesn't exist.
        items: List of drop specs (see above).
        page_name: Target page. Defaults to the active page.

    Returns:
        {"page_name": str, "count": int,
         "shapes": [{"shape_id", "stencil", "master", "x", "y",
                     "applied_data"}, ...]}
    """
    if not items:
        raise InvalidArgument("drop_masters called with empty items list")

    handle = ensure_document_open(file_path, create_if_missing=True)
    page = handle.get_page(page_name)

    # Pre-open every distinct stencil before any drops so a typo'd stencil
    # name fails fast instead of half-way through the batch.
    stencil_docs: dict = {}
    for i, item in enumerate(items):
        name = item.get("stencil")
        if not name:
            raise InvalidArgument(
                f"items[{i}]: missing 'stencil' name",
                details={"index": i, "spec": item},
            )
        if name not in stencil_docs:
            stencil_docs[name] = get_or_open_stencil(name)

    results = []
    with _batch_context(f"Drop {len(items)} masters"):
        for i, spec in enumerate(items):
            stencil_doc = stencil_docs[spec["stencil"]]
            shape = _drop_one(page, stencil_doc, spec, i)
            applied = _apply_data(shape, spec["data"]) if spec.get("data") else {}
            results.append({
                "shape_id": int(shape.ID),
                "stencil": spec["stencil"],
                "master": spec["master"],
                "x": float(spec["x"]),
                "y": float(spec["y"]),
                "applied_data": applied if applied else None,
            })

    return {"page_name": page.Name, "count": len(results), "shapes": results}
