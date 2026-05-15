"""
Template discovery tools — find branded starter documents to base new
diagrams on. Pair with `create_visio_file(template="...")` to actually
use one.

Configure paths via `CTI_VISIO_TEMPLATE_PATHS` (semicolon-separated).
The server also picks up Visio's own `Application.TemplatePaths`.
"""

from __future__ import annotations

from typing import Optional

from ..com import templates as _templates
from ..errors import InvalidArgument, envelope
from ..server_instance import mcp


@mcp.tool()
@envelope("list_templates")
async def list_templates(category: Optional[str] = None,
                         limit: int = 50, offset: int = 0) -> dict:
    """List indexed templates, optionally filtered by category.

    Args:
        category: Optional substring filter on the category (derived from
                 the immediate parent directory name relative to the
                 configured root). Case-insensitive.
        limit: Max templates to return (default 50).
        offset: Pagination offset (default 0).

    Returns:
        {"total": int, "returned": int,
         "templates": [{"name", "path", "category", "format"}, ...]}
    """
    _templates.ensure_indexed()
    all_entries = _templates.all_templates()
    if category:
        needle = category.lower()
        all_entries = [e for e in all_entries if needle in e.category.lower()]
    all_entries.sort(key=lambda e: (e.category.lower(), e.name.lower()))
    total = len(all_entries)
    sliced = all_entries[offset:offset + max(1, limit)]
    return {
        "total": total,
        "returned": len(sliced),
        "templates": [
            {
                "name": e.name,
                "path": e.path,
                "category": e.category,
                "format": e.format,
            }
            for e in sliced
        ],
    }


@mcp.tool()
@envelope("find_templates")
async def find_templates(query: str, limit: int = 10) -> dict:
    """Rank-search templates by name. Exact > substring > fuzzy.

    Args:
        query: Substring or full name (case-insensitive).
        limit: Max results (default 10).

    Returns:
        {"query": str, "count": int,
         "results": [{"name", "path", "category", "format", "score"}, ...]}
    """
    if not query or not query.strip():
        raise InvalidArgument("find_templates called with empty query")
    results = _templates.find_templates(query, limit=limit)
    return {"query": query, "count": len(results), "results": results}


@mcp.tool()
@envelope("template_index_status")
async def template_index_status() -> dict:
    """Report what's in the template index.

    Returns:
        {"built_at": str|null, "template_count": int,
         "configured_paths": [str], "paths_scanned": [str],
         "cache_path": str, "errors": [...]}
    """
    return _templates.status_snapshot()


@mcp.tool()
@envelope("reindex_templates")
async def reindex_templates(force: bool = False) -> dict:
    """Rebuild the template index. Cheap — pure file-system walk, no
    Visio opens needed (unlike stencils).

    Args:
        force: If True, ignore the per-file mtime cache and rebuild
              every entry. Defaults to False.

    Returns:
        {"template_count": int, "paths_scanned": [str], "errors": [...],
         "built_at": str}
    """
    status = _templates.build_index(force=bool(force))
    return {
        "template_count": status.template_count,
        "paths_scanned": list(status.paths_scanned),
        "errors": list(status.errors),
        "built_at": status.built_at,
    }
