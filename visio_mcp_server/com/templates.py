"""
Template discovery and indexing.

Visio templates (.vstx / .vst / .vstm) — and "starter" .vsdx files used
the same way — encode page geometry, theme, default stencils, and
sometimes pre-placed boilerplate. CTI's `CTI Visio Signal Flow
Template.vsdx` is an example: opening it as a new document gives every
AV drawing the same titleblock and grid setup.

Unlike stencils, templates don't need to be opened in Visio to be
indexed — we just want their filenames and paths. That makes the
template index cheap: pure file-system walk, sub-second on hundreds of
files.

Path discovery: env var `CTI_VISIO_TEMPLATE_PATHS` (semicolon-separated)
+ `Application.TemplatePaths` + Windows defaults. Same shape as the
stencil paths so the configuration story is consistent.

Cache: ``%USERPROFILE%/.cti-visio-mcp/template_index.json`` (per-file
mtime invalidation).
"""

from __future__ import annotations

import json
import logging
import os
import time
from dataclasses import asdict, dataclass, field
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable, Optional

from .app import get_visio_app

logger = logging.getLogger("visio_mcp.com.templates")

_TEMPLATE_EXTENSIONS = {".vstx", ".vst", ".vstm", ".vsdx", ".vsd", ".vsdm"}

_CACHE_DIR = Path(os.environ.get("USERPROFILE", "")) / ".cti-visio-mcp"
_CACHE_PATH = _CACHE_DIR / "template_index.json"
_INDEX_VERSION = 1


@dataclass
class TemplateEntry:
    path: str
    name: str       # filename without extension
    category: str   # immediate parent dir name relative to the configured root
    format: str     # "vstx" / "vsdx" / etc.
    mtime: float


@dataclass
class IndexStatus:
    built_at: Optional[str] = None
    template_count: int = 0
    paths_scanned: list = field(default_factory=list)
    errors: list = field(default_factory=list)


_status = IndexStatus()
_index: dict[str, TemplateEntry] = {}
_loaded_from_disk = False


# --------------------------------------------------------------------- paths

def _from_env() -> list[Path]:
    raw = os.environ.get("CTI_VISIO_TEMPLATE_PATHS", "")
    if not raw:
        return []
    return [Path(p.strip()) for p in raw.split(";") if p.strip()]


def _from_visio() -> list[Path]:
    try:
        app = get_visio_app()
        raw = str(app.TemplatePaths or "")
    except Exception:
        return []
    return [Path(p.strip()) for p in raw.split(";") if p.strip()]


def _windows_defaults() -> list[Path]:
    home = os.environ.get("USERPROFILE", "")
    if not home:
        return []
    return [Path(home) / "Documents" / "My Shapes"]  # rare, but possible


def discover_template_paths() -> list[Path]:
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


# ----------------------------------------------------------- index build

def _walk_templates(root: Path) -> Iterable[tuple[Path, Path]]:
    """Yield (file_path, root) pairs for templates under `root`."""
    for dirpath, _dirs, files in os.walk(root):
        for name in files:
            if name.startswith("~$"):
                continue
            ext = os.path.splitext(name)[1].lower()
            if ext in _TEMPLATE_EXTENSIONS:
                yield Path(dirpath) / name, root


def _category_for(file_path: Path, root: Path) -> str:
    """Use the file's immediate parent name as category, unless that's
    the root itself."""
    try:
        parent = file_path.parent.resolve()
        root_res = root.resolve()
    except OSError:
        return ""
    if parent == root_res:
        return ""
    return parent.name


def build_index(force: bool = False, paths: Optional[list[Path]] = None) -> IndexStatus:
    global _loaded_from_disk

    _ensure_loaded_from_disk()

    if paths is None:
        paths = discover_template_paths()

    _status.paths_scanned = [str(p) for p in paths]
    _status.errors = []

    all_pairs = []
    for root in paths:
        all_pairs.extend(_walk_templates(root))

    logger.info("template index: scanning %d files across %d roots",
                len(all_pairs), len(paths))

    seen: set[str] = set()
    for file_path, root in all_pairs:
        try:
            mtime = file_path.stat().st_mtime
        except OSError as e:
            _status.errors.append({"path": str(file_path), "error": str(e)})
            continue

        abs_path = str(file_path.resolve())
        seen.add(abs_path)

        prior = _index.get(abs_path)
        if prior is not None and not force and abs(prior.mtime - mtime) < 1e-6:
            continue  # cache hit — keep prior entry

        _index[abs_path] = TemplateEntry(
            path=abs_path,
            name=file_path.stem,
            category=_category_for(file_path, root),
            format=file_path.suffix.lower().lstrip("."),
            mtime=mtime,
        )

    # Drop entries for files that no longer exist.
    for stale in list(_index.keys()):
        if stale not in seen:
            del _index[stale]

    _status.built_at = time.strftime("%Y-%m-%dT%H:%M:%S")
    _status.template_count = len(_index)
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
        logger.warning("could not read template cache: %s", e)
        return
    if data.get("version") != _INDEX_VERSION:
        return
    for raw in data.get("templates", []):
        try:
            entry = TemplateEntry(**{k: raw[k] for k in (
                "path", "name", "category", "format", "mtime"
            )})
            _index[entry.path] = entry
        except Exception:
            continue
    _status.built_at = data.get("built_at")
    _status.template_count = len(_index)


def _save_cache() -> None:
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": _INDEX_VERSION,
            "built_at": _status.built_at,
            "templates": [asdict(e) for e in _index.values()],
        }
        with _CACHE_PATH.open("w", encoding="utf-8", newline="\n") as fh:
            json.dump(payload, fh, indent=2)
    except Exception as e:
        logger.warning("could not write template cache: %s", e)


# -------------------------------------------------------------- public API

def ensure_indexed() -> None:
    _ensure_loaded_from_disk()
    if _index:
        return
    build_index()


def all_templates() -> list[TemplateEntry]:
    _ensure_loaded_from_disk()
    return list(_index.values())


def get_template(name_or_path: str) -> Optional[TemplateEntry]:
    """Look up by absolute path or by name (filename without extension,
    case-insensitive)."""
    _ensure_loaded_from_disk()
    if name_or_path in _index:
        return _index[name_or_path]
    target = name_or_path.lower()
    target_base = os.path.splitext(os.path.basename(name_or_path))[0].lower()
    for entry in _index.values():
        if entry.name.lower() == target_base or entry.path.lower() == target:
            return entry
    return None


def status_snapshot() -> dict:
    _ensure_loaded_from_disk()
    return {
        "built_at": _status.built_at,
        "template_count": _status.template_count,
        "paths_scanned": _status.paths_scanned,
        "configured_paths": [str(p) for p in discover_template_paths()],
        "cache_path": str(_CACHE_PATH),
        "errors": list(_status.errors),
    }


def find_templates(query: str, limit: int = 10) -> list[dict]:
    """Rank-search by template name. Exact > substring > fuzzy."""
    ensure_indexed()
    q = query.lower().strip()
    if not q:
        return []
    results: list[tuple[float, dict]] = []
    for entry in _index.values():
        name = entry.name.lower()
        if name == q:
            score = 1000.0
        elif q in name:
            pos = name.find(q)
            score = 800.0 - min(pos, 50) * 2.0
        else:
            fuzz = SequenceMatcher(None, q, name).ratio()
            if fuzz < 0.5:
                continue
            score = 100.0 + (fuzz - 0.5) * 400.0
        results.append((score, {
            "name": entry.name,
            "path": entry.path,
            "category": entry.category,
            "format": entry.format,
            "score": round(score, 1),
        }))
    results.sort(key=lambda t: t[0], reverse=True)
    return [r for _, r in results[:limit]]
