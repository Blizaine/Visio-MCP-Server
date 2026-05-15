"""
Group tools — combine shapes that should be treated as a single unit.

Common AV uses: a rack of devices, a single-room subsystem, a logical
"DSP block" wrapping mic + matrix + amp shapes.

Visio's grouping works via Selection objects: create an empty selection,
add the target shapes to it, call `.Group()`. The result is itself a
Shape (with its own ID) whose `.Shapes` collection holds the members.
This means existing tools like `transform_shapes`, `style_shapes`, and
`delete_shapes` already work on groups — no per-group versions needed.
"""

from __future__ import annotations

from typing import Optional

from ..com.document import ensure_document_open, find_shape_on_page
from ..com.undo import undo_scope
from ..errors import InvalidArgument, ShapeNotFound, envelope
from ..server_instance import mcp

# Visio enum values used here. VisActionCodes governs Selection.Select():
#   visDeselect = 0, visSelect = 1, visSubSelect = 2, ...
# (Easy to mix up — using visSubSelect by accident causes Group() to grab
# entire connected components instead of just the listed shapes.)
_VIS_SEL_TYPE_EMPTY = 1       # visSelTypeEmpty
_VIS_SEL_MODE_SKIP_SUPER = 2  # visSelModeSkipSuper
_VIS_SELECT = 1               # visSelect (NOT visSubSelect=2)


@mcp.tool()
@envelope("group_shapes")
async def group_shapes(file_path: str, shape_ids: list,
                       page_name: Optional[str] = None) -> dict:
    """Combine multiple shapes into a single group.

    The resulting group is itself a Shape with its own ID; you can pass
    that ID to `transform_shapes` to move/resize the whole group,
    `style_shapes` to style it, `delete_shapes` to remove it, etc.

    Args:
        file_path: Path to the Visio file.
        shape_ids: IDs of the shapes to group (at least 2; grouping one
                  shape is a no-op error in Visio).
        page_name: Page the shapes live on. Defaults to active page.

    Returns:
        {"page_name": str, "group_id": int, "member_ids": [int],
         "member_count": int}
    """
    if not shape_ids or len(shape_ids) < 2:
        raise InvalidArgument(
            "group_shapes: need at least 2 shape_ids to form a group",
            details={"shape_ids": list(shape_ids)},
        )

    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    # Resolve every target up front; bad IDs fail before we touch a thing.
    resolved = []
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
        resolved.append((sid_int, shape))

    with undo_scope(f"Group {len(resolved)} shapes"):
        selection = page.CreateSelection(_VIS_SEL_TYPE_EMPTY, _VIS_SEL_MODE_SKIP_SUPER)
        for _sid, shape in resolved:
            selection.Select(shape, _VIS_SELECT)
        group_shape = selection.Group()

    return {
        "page_name": page.Name,
        "group_id": int(group_shape.ID),
        "member_ids": [sid for sid, _ in resolved],
        "member_count": len(resolved),
    }


@mcp.tool()
@envelope("ungroup_shape")
async def ungroup_shape(file_path: str, group_id: int,
                        page_name: Optional[str] = None) -> dict:
    """Break a group back into its constituent shapes.

    Member shapes keep their existing IDs and stay on the page; only
    the group wrapper goes away. Use `list_shapes` to enumerate the
    page state afterwards — different Visio builds disagree about how
    `group_shape.Shapes` and `page.Shapes` interact while a group
    exists, so this tool deliberately doesn't try to enumerate member
    IDs itself.

    Args:
        file_path: Path to the Visio file.
        group_id: ID of the group shape to ungroup.
        page_name: Page the group lives on. Defaults to active page.

    Returns:
        {"page_name": str, "ungrouped_id": int}
    """
    handle = ensure_document_open(file_path)
    page = handle.get_page(page_name)

    group_shape = find_shape_on_page(page, int(group_id))
    if group_shape is None:
        raise ShapeNotFound(
            f"shape ID {group_id} not on page '{page.Name}'",
            details={"group_id": group_id, "page_name": page.Name},
        )

    with undo_scope(f"Ungroup shape {group_id}"):
        group_shape.Ungroup()

    return {
        "page_name": page.Name,
        "ungrouped_id": int(group_id),
    }
