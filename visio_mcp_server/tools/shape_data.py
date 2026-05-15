"""
Shape data tools: read, write, and search custom property fields.

Most stencil-derived shapes carry a "Prop." (custom property) section in
their ShapeSheet — a Crestron switcher has Model/Inputs/Outputs, a Cisco
mic has SerialNumber/Channel, an AWS instance has Region/Type. These
fields are what turn a placed master from a decorative icon into an
information-bearing diagram element.

This phase exposes:

- `get_shape_data` — full property dump for one shape.
- `set_shape_data` / `set_shapes_data` — write to existing Prop.* cells.
  All-or-nothing: if any property doesn't exist, the whole update aborts
  with the missing names in `error.details.missing`. The undo scope
  rolls back any partial writes.
- `find_shapes_by_data` — search a page for shapes whose properties
  match a query dict.

For inheritance: Prop.* cells are inherited from the master. Setting a
value on an instance overrides the master's default for that instance
only. We can't create *new* property rows on instances (that requires
master-level edits), so writes only target rows that already exist.
"""

from __future__ import annotations

from typing import Optional

from ..com.document import (
    ensure_document_open,
    find_shape_on_page,
    format_prop_value,
    read_shape_data,
)
from ..com.undo import undo_scope
from ..errors import InvalidArgument, ShapeNotFound, envelope
from ..server_instance import mcp
from .batch import _batch_context


def _matches_query(props: dict, query: dict) -> bool:
    """AND-match. For each (name, expected) in query, the shape must have
    that property AND the value must match.

    - String expectations: case-insensitive substring search on the
      property's display value (matches the way humans actually look
      stuff up — "find shapes with Manufacturer containing 'cisco'").
    - Numeric / boolean expectations: exact equality on the typed value.
    """
    for name, expected in query.items():
        if name not in props:
            return False
        actual = props[name].get("value")
        if isinstance(expected, str):
            if not isinstance(actual, str):
                actual = str(actual) if actual is not None else ""
            if expected.lower() not in actual.lower():
                return False
        else:
            if actual != expected:
                return False
    return True


def _write_props(shape, data: dict) -> tuple[dict, list]:
    """Write Prop.<name> cells. Returns (applied, missing)."""
    applied: dict = {}
    missing: list = []
    for name, value in data.items():
        formula = format_prop_value(value)
        if formula is None:
            continue
        try:
            shape.Cells(f"Prop.{name}").FormulaU = formula
        except Exception:
            missing.append(name)
            continue
        applied[name] = value
    return applied, missing


# ----------------------------------------------------------------- tools

@mcp.tool()
@envelope("get_shape_data")
async def get_shape_data(file_path: str, shape_id: int,
                         page_name: Optional[str] = None) -> dict:
    """Read every custom property (Prop.* cell) on a shape.

    Use this to discover what fields a stencil master exposes before
    calling `set_shape_data`/`set_shapes_data`. Returns the property's
    name, current value (typed), display label, prompt/help text, and
    type. Returns an empty dict if the shape has no custom properties
    (most generic primitives don't).

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape to inspect.
        page_name: Page the shape lives on. Defaults to active page.

    Returns:
        {"shape_id": int, "page_name": str,
         "data": {
            "<PropName>": {
              "value": str|float|bool,
              "label": str,         # display label, e.g. "Serial Number"
              "prompt": str,        # tooltip/help
              "type": str,          # "String"|"Number"|"Boolean"|...
              "type_id": int,
              "formula": str        # raw ShapeSheet formula
            }, ...}}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    shape = find_shape_on_page(page, shape_id)
    if shape is None:
        raise ShapeNotFound(
            f"Could not find shape with ID {shape_id} on page '{page.Name}'",
            details={"shape_id": shape_id, "page_name": page.Name},
        )
    return {
        "shape_id": shape_id,
        "page_name": page.Name,
        "data": read_shape_data(shape),
    }


@mcp.tool()
@envelope("set_shape_data")
async def set_shape_data(file_path: str, shape_id: int, data: dict,
                         page_name: Optional[str] = None) -> dict:
    """Write custom property values on one shape. All-or-nothing.

    Only writes to property rows that already exist on the shape
    (typically inherited from its master). If any name in `data` doesn't
    exist on the shape, the whole update aborts with `PROPERTY_NOT_FOUND`
    and the missing names listed in `error.details.missing` — the undo
    scope rolls back any partial writes that landed before the miss.

    For batch writes across many shapes, use `set_shapes_data`.

    Args:
        file_path: Path to the Visio file.
        shape_id: ID of the shape.
        data: `{prop_name: value, ...}`. Values may be string/int/float/bool.
              None values are skipped (no write attempted).
        page_name: Page the shape lives on. Defaults to active page.

    Returns:
        {"shape_id": int, "page_name": str, "applied": {prop: value, ...}}
    """
    if not data:
        raise InvalidArgument("set_shape_data called with empty data dict")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)
    shape = find_shape_on_page(page, shape_id)
    if shape is None:
        raise ShapeNotFound(
            f"Could not find shape with ID {shape_id} on page '{page.Name}'",
            details={"shape_id": shape_id, "page_name": page.Name},
        )

    with undo_scope(f"Set data on shape {shape_id}"):
        applied, missing = _write_props(shape, data)
        if missing:
            raise InvalidArgument(
                f"Properties not found on shape {shape_id}: {missing}. "
                f"Use `get_shape_data` to list available properties.",
                details={"shape_id": shape_id, "missing": missing,
                         "would_apply": applied},
            )

    return {"shape_id": shape_id, "page_name": page.Name, "applied": applied}


@mcp.tool()
@envelope("set_shapes_data")
async def set_shapes_data(file_path: str, updates: list,
                          page_name: Optional[str] = None) -> dict:
    """Write custom property values across many shapes in one batch.

    Each entry in `updates`:
        {"shape_id": int, "data": {prop_name: value, ...}}

    All-or-nothing: a missing property on ANY shape aborts the whole
    batch. The undo scope rolls back every write in progress.

    Args:
        file_path: Path to the Visio file.
        updates: List of `{shape_id, data}` updates.
        page_name: Page the shapes live on. Defaults to active page.

    Returns:
        {"page_name": str, "count": int,
         "results": [{"shape_id", "applied"}, ...]}
    """
    if not updates:
        raise InvalidArgument("set_shapes_data called with empty updates list")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    # Resolve every target shape up front so a bad ID fails fast.
    targets = []
    for i, upd in enumerate(updates):
        try:
            shape_id = int(upd["shape_id"])
        except (KeyError, TypeError, ValueError):
            raise InvalidArgument(
                f"updates[{i}]: missing or invalid shape_id",
                details={"index": i, "spec": upd},
            )
        data = upd.get("data") or {}
        if not data:
            raise InvalidArgument(
                f"updates[{i}]: missing or empty data dict",
                details={"index": i, "shape_id": shape_id},
            )
        shape = find_shape_on_page(page, shape_id)
        if shape is None:
            raise ShapeNotFound(
                f"updates[{i}]: shape ID {shape_id} not on page '{page.Name}'",
                details={"index": i, "shape_id": shape_id, "page_name": page.Name},
            )
        targets.append((i, shape_id, shape, data))

    results = []
    aggregated_missing: list = []
    with _batch_context(f"Set data on {len(targets)} shapes"):
        for i, shape_id, shape, data in targets:
            applied, missing = _write_props(shape, data)
            if missing:
                aggregated_missing.append({"index": i, "shape_id": shape_id, "missing": missing})
                # Raise inside the batch so the undo scope rolls everything back.
                raise InvalidArgument(
                    f"updates[{i}]: properties not found on shape {shape_id}: {missing}",
                    details={
                        "first_missing": aggregated_missing[0],
                        "results_so_far": results,
                    },
                )
            results.append({"shape_id": shape_id, "applied": applied})

    return {"page_name": page.Name, "count": len(results), "results": results}


@mcp.tool()
@envelope("find_shapes_by_data")
async def find_shapes_by_data(file_path: str, query: dict,
                              page_name: Optional[str] = None,
                              limit: int = 50) -> dict:
    """Search a page for shapes whose custom properties match a query.

    AND-semantics: a shape must match every property in `query`. Strings
    are matched case-insensitively as substrings ("cisco" matches
    "Cisco Systems"). Numbers and booleans are matched exactly.

    Useful for: "find every device whose Manufacturer is Crestron",
    "find all shapes with VLAN=10", "find shapes missing a SerialNumber"
    (run get_shape_data first to confirm field names).

    Args:
        file_path: Path to the Visio file.
        query: `{prop_name: expected_value, ...}`.
        page_name: Page to search. Defaults to active page.
        limit: Max results returned (default 50).

    Returns:
        {"page_name": str, "query": dict, "count": int,
         "shapes": [{"shape_id", "name", "text", "data"}, ...]}
    """
    if not query:
        raise InvalidArgument("find_shapes_by_data called with empty query")

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    matches = []
    for shape in page.Shapes:
        try:
            props = read_shape_data(shape)
        except Exception:
            continue
        if _matches_query(props, query):
            matches.append({
                "shape_id": int(shape.ID),
                "name": str(shape.Name),
                "text": str(shape.Text),
                "data": props,
            })
            if len(matches) >= max(1, limit):
                break

    return {
        "page_name": page.Name,
        "query": query,
        "count": len(matches),
        "shapes": matches,
    }
