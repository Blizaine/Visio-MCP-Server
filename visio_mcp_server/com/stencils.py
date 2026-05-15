"""
Stencil discovery, indexing, and search.

A Visio stencil is a `.vss`, `.vssx`, or `.vssm` file containing master
shapes. Real diagrams are built by dropping masters from stencils onto
pages — generic primitives (`page.DrawRectangle` etc.) only get you so far.

This module:
1. Discovers stencil paths (env var, Visio's StencilPaths, Windows defaults).
2. Builds an index of all stencils + masters reachable from those paths.
3. Caches the index to disk so subsequent server starts are fast.
4. Provides ranked search (`find_masters`) across the index.

Indexing requires opening every stencil in Visio briefly to enumerate its
`Masters` collection. This is expensive — minutes for a few hundred
stencils — but only happens on first run or when a stencil's mtime changes.

Macro safety: `.vssm` files contain VBA macros. Before opening any
stencil, we set `Application.AutomationSecurity = 3`
(`msoAutomationSecurityForceDisable`) so macros never run during the
indexed open. Restored to the prior value after the indexing pass.
"""

from __future__ import annotations

import json
import logging
import os
import time
from dataclasses import dataclass, field, asdict
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable, Optional

from ..errors import ComError
from .app import get_visio_app

logger = logging.getLogger("visio_mcp.com.stencils")

_STENCIL_EXTENSIONS = {".vss", ".vssx", ".vssm"}

# Visio enum values (from the VisOpenSaveArgs / VisAutomationSecurity).
# Hard-coded so we don't depend on `win32com.client.constants` being populated.
_VIS_OPEN_HIDDEN = 0x40        # visOpenHidden
_VIS_OPEN_RO = 0x2             # visOpenRO
_VIS_OPEN_NO_WORKSPACE = 0x100 # visOpenNoWorkspace — don't change UI workspace
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# The on-disk cache file. Per-user so multiple repos sharing a host don't
# stomp on each other.
_CACHE_DIR = Path(os.environ.get("USERPROFILE", "")) / ".cti-visio-mcp"
_CACHE_PATH = _CACHE_DIR / "stencil_index.json"
_INDEX_VERSION = 1


@dataclass
class MasterEntry:
    name: str
    base_id: str
    prompt: str = ""
    prop_names: list = field(default_factory=list)


@dataclass
class StencilEntry:
    path: str
    name: str            # filename without extension
    manufacturer: str    # heuristic: first token of name before space/underscore
    format: str          # "vss", "vssx", "vssm"
    mtime: float         # for cache invalidation
    masters: list = field(default_factory=list)  # list[MasterEntry as dict]


@dataclass
class IndexStatus:
    """In-memory state used by the status tool."""
    built_at: Optional[str] = None
    stencil_count: int = 0
    master_count: int = 0
    paths_scanned: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    in_progress: bool = False
    current_file: Optional[str] = None
    progress: Optional[str] = None  # e.g. "47/200"


# Singleton index state. Loaded lazily from disk on first access.
_status = IndexStatus()
_index: dict[str, StencilEntry] = {}   # keyed by absolute path
_loaded_from_disk = False


# --------------------------------------------------------------------- paths

def _from_env() -> list[Path]:
    raw = os.environ.get("CTI_VISIO_STENCIL_PATHS", "")
    if not raw:
        return []
    return [Path(p.strip()) for p in raw.split(";") if p.strip()]


def _from_visio() -> list[Path]:
    """Visio's own StencilPaths property — semicolon-separated."""
    try:
        app = get_visio_app()
        raw = str(app.StencilPaths or "")
    except Exception:
        return []
    return [Path(p.strip()) for p in raw.split(";") if p.strip()]


def _windows_defaults() -> list[Path]:
    home = os.environ.get("USERPROFILE", "")
    if not home:
        return []
    return [Path(home) / "Documents" / "My Shapes"]


def discover_stencil_paths() -> list[Path]:
    """Return the deduplicated, existing directories we'll scan."""
    seen: set[str] = set()
    out: list[Path] = []
    for src in (_from_env, _from_visio, _windows_defaults):
        for p in src():
            key = str(p.resolve()).lower() if p.exists() else str(p).lower()
            if key in seen:
                continue
            seen.add(key)
            if p.exists() and p.is_dir():
                out.append(p)
    return out


def _walk_stencils(root: Path) -> Iterable[Path]:
    """Yield .vss/.vssx/.vssm files under `root`, recursively. Skips
    Visio's transient ~$lock files."""
    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            if name.startswith("~$"):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in _STENCIL_EXTENSIONS:
                yield Path(dirpath) / name


# ----------------------------------------------------------- index build

def _manufacturer_from(stem: str) -> str:
    """Heuristic: take the first whitespace- or underscore-separated token.
    `Crestron NVX` -> `Crestron`; `Cisco_Network` -> `Cisco`."""
    for sep in (" ", "_"):
        if sep in stem:
            return stem.split(sep, 1)[0]
    return stem


def _open_stencil_for_indexing(app, path: str):
    """Open a stencil hidden + read-only + don't change the workspace."""
    flags = _VIS_OPEN_HIDDEN | _VIS_OPEN_RO | _VIS_OPEN_NO_WORKSPACE
    return app.Documents.OpenEx(path, flags)


def _harvest_masters(doc) -> list[dict]:
    out = []
    for i in range(1, doc.Masters.Count + 1):
        m = doc.Masters.Item(i)
        try:
            base_id = str(m.BaseID)
        except Exception:
            base_id = ""
        try:
            prompt = str(m.Prompt or "")
        except Exception:
            prompt = ""
        prop_names = _harvest_prop_names(m)
        out.append(asdict(MasterEntry(
            name=str(m.Name),
            base_id=base_id,
            prompt=prompt,
            prop_names=prop_names,
        )))
    return out


def _harvest_prop_names(master) -> list[str]:
    """Read the master's Shape -> User-defined Cells custom property labels.

    Visio stores custom shape data in the 'Prop' section of the ShapeSheet.
    Reading these without prompting the user requires touching the master's
    `Shapes.ItemFromID(0)` (the master's shape itself). Best-effort — many
    masters either don't expose this metadata cleanly or use generic names.
    """
    try:
        shape = master.Shapes.ItemFromID(0)
    except Exception:
        return []
    names = []
    try:
        for i in range(1, shape.RowCount(243) + 1):  # 243 = visSectionProp
            try:
                row_name = shape.CellsSRC(243, i - 1, 0).RowName
                if row_name:
                    names.append(str(row_name))
            except Exception:
                continue
    except Exception:
        return []
    return names


def _build_one_stencil(app, path: Path, prior_mtime: Optional[float]) -> Optional[StencilEntry]:
    """Index a single stencil file. Returns None on failure (logged)."""
    try:
        mtime = path.stat().st_mtime
    except OSError as e:
        logger.warning("stencil unreadable %s: %s", path, e)
        _status.errors.append({"path": str(path), "error": str(e)})
        return None

    abs_path = str(path.resolve())

    # Cache hit?
    if prior_mtime is not None and abs(prior_mtime - mtime) < 1e-6:
        # Caller is responsible for keeping the existing entry from cache.
        return None

    ext = path.suffix.lower().lstrip(".")
    stem = path.stem
    entry = StencilEntry(
        path=abs_path,
        name=stem,
        manufacturer=_manufacturer_from(stem),
        format=ext,
        mtime=mtime,
    )

    doc = None
    try:
        doc = _open_stencil_for_indexing(app, abs_path)
        entry.masters = _harvest_masters(doc)
    except Exception as e:
        logger.warning("stencil open failed %s: %s", abs_path, e)
        _status.errors.append({"path": abs_path, "error": str(e)})
        entry.masters = []
    finally:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass

    return entry


def build_index(force: bool = False, paths: Optional[list[Path]] = None) -> IndexStatus:
    """Walk discovered paths and (re)build the index.

    Reuses cached entries whose `mtime` matches the file on disk unless
    `force=True`. Streams progress to stderr.
    """
    global _index, _loaded_from_disk

    _ensure_loaded_from_disk()

    if paths is None:
        paths = discover_stencil_paths()

    _status.paths_scanned = [str(p) for p in paths]
    _status.errors = []
    _status.in_progress = True
    _status.current_file = None

    # Enumerate every stencil file we'll touch.
    all_files = []
    for root in paths:
        all_files.extend(_walk_stencils(root))
    total = len(all_files)
    logger.info("stencil index: scanning %d files across %d roots", total, len(paths))

    app = get_visio_app()
    prior_security = None
    try:
        prior_security = app.AutomationSecurity
    except Exception:
        prior_security = None
    try:
        app.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception as e:
        logger.warning("could not lower AutomationSecurity (macros may run): %s", e)

    try:
        seen_paths: set[str] = set()
        for i, path in enumerate(all_files, 1):
            _status.current_file = str(path)
            _status.progress = f"{i}/{total}"
            abs_path = str(path.resolve())
            seen_paths.add(abs_path)

            prior = _index.get(abs_path)
            prior_mtime = prior.mtime if prior and not force else None
            new_entry = _build_one_stencil(app, path, prior_mtime)
            if new_entry is not None:
                _index[abs_path] = new_entry
                logger.info("indexed %s (%d masters)", path.name, len(new_entry.masters))
            elif prior is not None:
                # Cache hit — keep the prior entry as-is.
                logger.debug("cached %s", path.name)

        # Drop entries for files that no longer exist on disk.
        for stale in list(_index.keys()):
            if stale not in seen_paths:
                del _index[stale]
                logger.info("removed stale index entry %s", stale)

    finally:
        if prior_security is not None:
            try:
                app.AutomationSecurity = prior_security
            except Exception:
                pass
        _status.in_progress = False
        _status.current_file = None
        _status.progress = None

    _status.built_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    _status.stencil_count = len(_index)
    _status.master_count = sum(len(e.masters) for e in _index.values())
    _save_cache()
    return _status


# ---------------------------------------------------------------- cache I/O

def _ensure_loaded_from_disk() -> None:
    global _loaded_from_disk
    if _loaded_from_disk:
        return
    _loaded_from_disk = True
    if not _CACHE_PATH.exists():
        return
    try:
        with _CACHE_PATH.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception as e:
        logger.warning("could not read stencil cache %s: %s", _CACHE_PATH, e)
        return
    if data.get("version") != _INDEX_VERSION:
        logger.info("stencil cache version mismatch; will rebuild")
        return
    for raw in data.get("stencils", []):
        try:
            entry = StencilEntry(**{k: raw[k] for k in (
                "path", "name", "manufacturer", "format", "mtime", "masters"
            )})
            _index[entry.path] = entry
        except Exception as e:
            logger.warning("dropping bad cache entry: %s", e)
    _status.built_at = data.get("built_at")
    _status.stencil_count = len(_index)
    _status.master_count = sum(len(e.masters) for e in _index.values())


def _save_cache() -> None:
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": _INDEX_VERSION,
            "built_at": _status.built_at,
            "stencils": [asdict(e) for e in _index.values()],
        }
        with _CACHE_PATH.open("w", encoding="utf-8", newline="\n") as fh:
            json.dump(payload, fh, indent=2)
    except Exception as e:
        logger.warning("could not write stencil cache %s: %s", _CACHE_PATH, e)


# -------------------------------------------------------------- public API

def ensure_indexed() -> None:
    """Build the index if we have nothing on disk and nothing in memory.
    Used by `find_masters` etc. to bootstrap lazily on first use."""
    _ensure_loaded_from_disk()
    if _index:
        return
    build_index()


def all_stencils() -> list[StencilEntry]:
    _ensure_loaded_from_disk()
    return list(_index.values())


def get_stencil(name_or_path: str) -> Optional[StencilEntry]:
    """Find a stencil by absolute path OR by basename (with or without ext)."""
    _ensure_loaded_from_disk()
    if name_or_path in _index:
        return _index[name_or_path]
    target = os.path.basename(name_or_path).lower()
    target_stem = os.path.splitext(target)[0]
    for entry in _index.values():
        if entry.name.lower() == target_stem or os.path.basename(entry.path).lower() == target:
            return entry
    return None


def status_snapshot() -> dict:
    _ensure_loaded_from_disk()
    return {
        "built_at": _status.built_at,
        "stencil_count": _status.stencil_count,
        "master_count": _status.master_count,
        "paths_scanned": _status.paths_scanned,
        "configured_paths": [str(p) for p in discover_stencil_paths()],
        "cache_path": str(_CACHE_PATH),
        "in_progress": _status.in_progress,
        "current_file": _status.current_file,
        "progress": _status.progress,
        "errors": list(_status.errors),
    }


# ------------------------------------------------------------------ search

def _score_master(query: str, master_name: str, manufacturer: str, prompt: str) -> float:
    """Return a search score in [0, 1000]. Higher = better match.

    Designed for human-typed queries like "Crestron DM-NVX-360" or "NVX-360".
    """
    q = query.lower().strip()
    if not q:
        return 0.0
    name = master_name.lower()
    mfr = manufacturer.lower()
    prom = prompt.lower()

    # Exact name match.
    if name == q:
        return 1000.0
    # Substring matches.
    if q in name:
        # Earlier-in-name is better.
        pos = name.find(q)
        return 800.0 - min(pos, 50) * 2.0
    # Manufacturer + remainder pattern: "Crestron NVX-360" → split and score
    # name with the non-manufacturer tokens.
    tokens = q.split()
    name_hits = sum(1 for t in tokens if t in name)
    mfr_hits = sum(1 for t in tokens if t in mfr)
    prom_hits = sum(1 for t in tokens if t in prom)
    if name_hits or mfr_hits:
        base = 200.0 + 100.0 * name_hits + 80.0 * mfr_hits + 20.0 * prom_hits
    else:
        base = 0.0

    # Fuzzy fallback on the master name.
    fuzz = SequenceMatcher(None, q, name).ratio()
    if fuzz > 0.5:
        base = max(base, 100.0 + (fuzz - 0.5) * 400.0)

    return base


# ---------------------------------------------------------- runtime cache

# Cache of open stencil documents, keyed by absolute path. Opened lazily on
# the first drop from a given stencil and reused for the lifetime of the
# server process. Opening a stencil is ~0.5s; reuse matters when an AV
# diagram drops a dozen masters in a single batch.
_open_stencil_docs: dict = {}


def get_or_open_stencil(name_or_path: str):
    """Return a live COM Document for the named stencil, opening it hidden
    + read-only if it isn't already cached.

    Looks the stencil up by short name (filename without extension) or full
    absolute path. Returns the same Document object on repeated calls.
    Raises ComError on COM failure, ValueError if the stencil isn't in the
    index — callers should let those propagate up to the envelope.
    """
    entry = get_stencil(name_or_path)
    if entry is None:
        raise ValueError(
            f"Stencil {name_or_path!r} not in index. "
            "Run `reindex_stencils` if you recently added it; "
            "check `stencil_index_status` for configured paths."
        )

    cached = _open_stencil_docs.get(entry.path)
    if cached is not None:
        try:
            _ = cached.Name  # liveness probe
            return cached
        except Exception:
            _open_stencil_docs.pop(entry.path, None)

    app = get_visio_app()
    flags = _VIS_OPEN_HIDDEN | _VIS_OPEN_RO
    try:
        doc = app.Documents.OpenEx(entry.path, flags)
    except Exception as e:
        raise ComError(f"Could not open stencil {entry.path}: {e}") from e
    _open_stencil_docs[entry.path] = doc
    logger.info("stencil opened hidden for drops: %s", entry.path)
    return doc


def close_all_stencils() -> None:
    """Close every cached stencil. Called from `close_visio_app` at process
    exit. Best-effort — COM cleanup during interpreter shutdown is flaky."""
    for path, doc in list(_open_stencil_docs.items()):
        try:
            doc.Close()
        except Exception:
            pass
    _open_stencil_docs.clear()


# ------------------------------------------------------------------ search

def find_masters(query: str, manufacturer: Optional[str] = None,
                 limit: int = 10) -> list[dict]:
    """Rank-search masters across the entire index."""
    ensure_indexed()
    mfr_filter = manufacturer.lower() if manufacturer else None
    results: list[tuple[float, dict]] = []
    for stencil in _index.values():
        if mfr_filter and stencil.manufacturer.lower() != mfr_filter:
            continue
        for m in stencil.masters:
            score = _score_master(query, m["name"], stencil.manufacturer, m.get("prompt", ""))
            if score <= 0:
                continue
            results.append((score, {
                "master_name": m["name"],
                "stencil": stencil.name,
                "stencil_path": stencil.path,
                "manufacturer": stencil.manufacturer,
                "prompt": m.get("prompt", ""),
                "prop_names": m.get("prop_names", []),
                "score": round(score, 1),
            }))
    results.sort(key=lambda t: t[0], reverse=True)
    return [r for _, r in results[:limit]]
