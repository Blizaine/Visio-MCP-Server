"""
Stencil discovery and search tools.

Use these to find master shapes BEFORE dropping them onto a page. The
typical flow for an AV diagram is:

    find_masters("Crestron NVX-360") -> picks the right master/stencil
    drop_master(...)                  -> place it on the page (Phase 8)

Run `reindex_stencils()` once when you first install (or change your
stencil directory) to populate the on-disk index. Indexing 200+ stencils
can take several minutes — Visio has to open each file briefly to
enumerate its masters. Subsequent calls are fast (cached).
"""

from __future__ import annotations

from typing import Optional

from ..com import stencils as _stencils
from ..errors import InvalidArgument, envelope
from ..server_instance import mcp


@mcp.tool()
@envelope("list_stencils")
async def list_stencils(manufacturer: Optional[str] = None,
                        limit: int = 50, offset: int = 0) -> dict:
    """List indexed stencils, optionally filtered by manufacturer.

    Args:
        manufacturer: Optional substring filter on the manufacturer name
                     (derived from the stencil filename — e.g. "Crestron NVX.vssx"
                     -> manufacturer "Crestron"). Case-insensitive.
        limit: Max stencils to return (default 50).
        offset: Pagination offset (default 0).

    Returns:
        {"total": int, "returned": int, "stencils": [{"name", "path",
         "manufacturer", "format", "master_count"}]}
    """
    _stencils.ensure_indexed()
    all_entries = _stencils.all_stencils()
    if manufacturer:
        needle = manufacturer.lower()
        all_entries = [e for e in all_entries if needle in e.manufacturer.lower()]
    all_entries.sort(key=lambda e: (e.manufacturer.lower(), e.name.lower()))
    total = len(all_entries)
    sliced = all_entries[offset:offset + max(1, limit)]
    return {
        "total": total,
        "returned": len(sliced),
        "stencils": [
            {
                "name": e.name,
                "path": e.path,
                "manufacturer": e.manufacturer,
                "format": e.format,
                "master_count": len(e.masters),
            }
            for e in sliced
        ],
    }


@mcp.tool()
@envelope("list_masters")
async def list_masters(stencil: str, query: Optional[str] = None,
                       limit: int = 50) -> dict:
    """List master shapes in a specific stencil.

    Args:
        stencil: Stencil name (e.g. "Crestron NVX") or full path. The
                short name lookup is case-insensitive and matches the
                file's basename without extension.
        query: Optional substring filter on master names (case-insensitive).
        limit: Max masters to return (default 50).

    Returns:
        {"stencil": str, "path": str, "manufacturer": str, "total": int,
         "returned": int, "masters": [{"name", "prompt", "prop_names"}]}
    """
    entry = _stencils.get_stencil(stencil)
    if entry is None:
        raise InvalidArgument(
            f"Stencil {stencil!r} not in index. Try `list_stencils` to see what's available, "
            "or `reindex_stencils` if you added new ones.",
            details={"stencil": stencil},
        )
    masters = entry.masters
    if query:
        q = query.lower()
        masters = [m for m in masters if q in m["name"].lower()]
    total = len(masters)
    sliced = masters[:max(1, limit)]
    return {
        "stencil": entry.name,
        "path": entry.path,
        "manufacturer": entry.manufacturer,
        "total": total,
        "returned": len(sliced),
        "masters": [
            {
                "name": m["name"],
                "prompt": m.get("prompt", ""),
                "prop_names": m.get("prop_names", []),
            }
            for m in sliced
        ],
    }


@mcp.tool()
@envelope("find_masters")
async def find_masters(query: str, manufacturer: Optional[str] = None,
                       limit: int = 10) -> dict:
    """Rank-search master shapes across ALL indexed stencils.

    This is the workhorse for "I need a Crestron DM-NVX-360" — the model
    types the model number and gets back the best matches with their
    stencil and manufacturer context.

    Ranking favors: exact name match > substring match > token overlap
    with name + manufacturer > fuzzy similarity. Manufacturer matches are
    boosted because stencil filenames are organized by manufacturer.

    Args:
        query: What to search for. Tokens separated by whitespace.
        manufacturer: Restrict to one manufacturer (exact, case-insensitive).
                     Useful when you know the brand and want to filter noise.
        limit: Max results (default 10).

    Returns:
        {"query": str, "count": int,
         "results": [{"master_name", "stencil", "stencil_path",
                      "manufacturer", "prompt", "prop_names", "score"}, ...]}
    """
    if not query or not query.strip():
        raise InvalidArgument("find_masters called with empty query")
    results = _stencils.find_masters(query, manufacturer=manufacturer, limit=limit)
    return {"query": query, "count": len(results), "results": results}


@mcp.tool()
@envelope("stencil_index_status")
async def stencil_index_status() -> dict:
    """Report what's in the stencil index right now.

    Useful for diagnosing "find_masters returns nothing" — usually means
    the index is empty (call `reindex_stencils`), points at the wrong
    directories (set `CTI_VISIO_STENCIL_PATHS`), or a stencil failed to
    open (see `errors`).

    Returns:
        {"built_at": str|null, "stencil_count": int, "master_count": int,
         "configured_paths": [str], "paths_scanned": [str],
         "cache_path": str, "in_progress": bool, "progress": str|null,
         "current_file": str|null, "errors": [{"path", "error"}]}
    """
    return _stencils.status_snapshot()


@mcp.tool()
@envelope("reindex_stencils")
async def reindex_stencils(force: bool = False) -> dict:
    """Rebuild the stencil index from disk. Run this after adding,
    removing, or modifying stencil files.

    The build process opens every stencil in Visio briefly (hidden,
    read-only, macros disabled). For ~200 stencils this can take several
    minutes the first time; subsequent runs reuse cached entries for
    files whose mtime hasn't changed.

    Args:
        force: If True, ignore the per-file mtime cache and rebuild
              every entry from scratch. Defaults to False.

    Returns:
        {"stencil_count": int, "master_count": int, "errors": [...],
         "paths_scanned": [str], "built_at": str}
    """
    status = _stencils.build_index(force=bool(force))
    return {
        "stencil_count": status.stencil_count,
        "master_count": status.master_count,
        "errors": list(status.errors),
        "paths_scanned": list(status.paths_scanned),
        "built_at": status.built_at,
    }
