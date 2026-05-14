"""
Document handle wrapper plus the in-process registry of open documents.

Each tool call resolves a file path to a `DocumentHandle` via
`ensure_document_open(path)`. The handle wraps the live Visio COM Document
object and the canonical, normalized path key used for caching.

The registry is keyed on `normalize_path(p) = normcase(abspath(p))` so that
`C:\\foo.vsdx`, `c:/foo.vsdx`, and `./foo.vsdx` all map to one entry.
"""

from __future__ import annotations

import logging
import os
import time
from typing import Optional

from ..errors import ComError, PageNotFound, VisioFileNotFound
from .app import get_visio_app

logger = logging.getLogger("visio_mcp.com.document")

_open_documents: dict[str, "DocumentHandle"] = {}


def normalize_path(p: str) -> str:
    return os.path.normcase(os.path.abspath(p))


class DocumentHandle:
    """Wraps a Visio COM Document + its normalized path."""

    def __init__(self, com_doc, normalized_path: str) -> None:
        self.com_doc = com_doc
        self.path = normalized_path

    @property
    def name(self) -> str:
        return self.com_doc.Name

    def is_alive(self) -> bool:
        try:
            _ = self.com_doc.Name
            return True
        except Exception:
            return False

    def save(self) -> None:
        self.com_doc.Save()

    def close(self, save_changes: bool = True) -> None:
        if save_changes:
            try:
                self.com_doc.Save()
            except Exception as e:
                logger.warning("save before close failed for %s: %s", self.path, e)
        self.com_doc.Close()

    def get_page(self, page_name: Optional[str] = None):
        """Resolve a page on this document.

        - `page_name=None`: returns Application.ActivePage if it belongs to
          this document, otherwise falls back to this document's first page.
          The fallback fixes a latent bug where a multi-doc workflow could
          edit the wrong document's page.
        - `page_name="Foo"`: looks up by name (case-sensitive on Visio's side);
          raises PageNotFound if not present.
        """
        if page_name is not None:
            try:
                return self.com_doc.Pages.ItemU(page_name)
            except Exception:
                raise PageNotFound(
                    f"Page '{page_name}' not found in document",
                    details={"document": self.path, "page_name": page_name},
                )

        app = get_visio_app()
        try:
            active = app.ActivePage
            if active is not None and active.Document.FullName == self.com_doc.FullName:
                return active
        except Exception:
            pass
        return self.com_doc.Pages.Item(1)


def create_document(template_path: Optional[str], save_path: str) -> DocumentHandle:
    app = get_visio_app()
    try:
        if template_path and os.path.exists(template_path):
            com_doc = app.Documents.Add(template_path)
        else:
            com_doc = app.Documents.Add("")
        time.sleep(1)
        com_doc.SaveAs(save_path)
    except Exception as e:
        raise ComError(f"Could not create Visio document at {save_path}: {e}") from e

    handle = DocumentHandle(com_doc, normalize_path(save_path))
    _open_documents[handle.path] = handle
    logger.info("created document path=%s", save_path)
    return handle


def open_document(file_path: str) -> DocumentHandle:
    if not os.path.exists(file_path):
        raise VisioFileNotFound(f"Visio file does not exist: {file_path}")

    app = get_visio_app()
    try:
        com_doc = app.Documents.Open(file_path)
    except Exception as e:
        raise ComError(f"Could not open Visio document {file_path}: {e}") from e

    handle = DocumentHandle(com_doc, normalize_path(file_path))
    _open_documents[handle.path] = handle
    logger.info("opened document path=%s", file_path)
    return handle


def ensure_document_open(file_path: str, create_if_missing: bool = False) -> DocumentHandle:
    """Return a live `DocumentHandle` for `file_path`, opening or creating as needed.

    Stale cache entries (handle whose underlying COM object is dead) are
    evicted automatically.
    """
    key = normalize_path(file_path)
    cached = _open_documents.get(key)
    if cached is not None:
        if cached.is_alive():
            return cached
        del _open_documents[key]

    if os.path.exists(file_path):
        return open_document(file_path)
    if create_if_missing:
        return create_document(None, file_path)
    raise VisioFileNotFound(f"Visio file does not exist: {file_path}")


def get_open_document(file_path: str) -> Optional[DocumentHandle]:
    """Return the cached handle for `file_path` if present and live, else None."""
    key = normalize_path(file_path)
    cached = _open_documents.get(key)
    if cached is None:
        return None
    if not cached.is_alive():
        del _open_documents[key]
        return None
    return cached


def forget_document(file_path: str) -> None:
    """Drop a handle from the registry. Caller is responsible for closing it."""
    key = normalize_path(file_path)
    _open_documents.pop(key, None)


def close_all_documents() -> None:
    """Used by the atexit hook in com.app — best-effort close-everything."""
    for path, handle in list(_open_documents.items()):
        try:
            handle.com_doc.Close()
        except Exception:
            pass
        del _open_documents[path]
