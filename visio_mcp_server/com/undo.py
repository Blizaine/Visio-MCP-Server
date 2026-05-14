"""
Undo-scope context manager.

Wraps a block of Visio mutations in an application-level undo scope so that
partial failures cancel cleanly: if any exception escapes the `with` block,
EndUndoScope is called with commit=False, rolling the changes back. A clean
exit commits.

Tools should wrap any multi-step mutation in this scope. Single-shape
mutations don't strictly need it, but the cost is one extra COM call and
keeps behavior consistent.
"""

from __future__ import annotations

import logging
from contextlib import contextmanager

from .app import get_visio_app

logger = logging.getLogger("visio_mcp.com.undo")


@contextmanager
def undo_scope(name: str):
    app = get_visio_app()
    try:
        scope_id = app.BeginUndoScope(name)
    except Exception as e:
        # If Visio refuses to begin a scope, fall through without one — better
        # than failing the whole tool. The user just loses undo granularity.
        logger.warning("BeginUndoScope(%r) failed: %s", name, e)
        yield
        return

    try:
        yield
    except Exception:
        try:
            app.EndUndoScope(scope_id, False)
        except Exception as e:
            logger.warning("EndUndoScope(commit=False) failed: %s", e)
        raise
    else:
        try:
            app.EndUndoScope(scope_id, True)
        except Exception as e:
            logger.warning("EndUndoScope(commit=True) failed: %s", e)
