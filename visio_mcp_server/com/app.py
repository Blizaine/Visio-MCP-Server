"""
Visio Application lifecycle and COM-thread initialization.

The Application object is a process-global singleton — there's only ever one
Visio.Application per host process, and Visio enforces that. The lazy
`get_visio_app()` returns it, launching Visio on first call.

`ensure_com_initialized()` exists because FastMCP may dispatch tool calls on
worker threads that haven't yet called CoInitialize; calling it at every
tool entry is cheap and idempotent on a thread that already has it.
"""

from __future__ import annotations

import atexit
import logging
import time
import winreg

import pythoncom
import win32com.client

from ..errors import VisioUnavailable

logger = logging.getLogger("visio_mcp.com.app")

_visio_app = None


def ensure_com_initialized() -> None:
    try:
        pythoncom.CoInitialize()
    except pythoncom.com_error:
        pass  # already initialized on this thread; refcount bumped


def check_visio_installed() -> bool:
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Visio.Application")
        return True
    except OSError:
        return False


def get_visio_app():
    global _visio_app
    if _visio_app is not None:
        return _visio_app

    attempts = (
        ("Dispatch", win32com.client.Dispatch),
        ("dynamic.Dispatch", win32com.client.dynamic.Dispatch),
        ("DispatchEx", win32com.client.DispatchEx),
    )
    errors = []
    for name, fn in attempts:
        try:
            _visio_app = fn("Visio.Application")
            time.sleep(1)
            _visio_app.Visible = True
            logger.info("Visio.Application acquired via %s", name)
            return _visio_app
        except Exception as e:
            errors.append(f"{name}: {e}")
            logger.debug("Visio acquisition via %s failed: %s", name, e)

    raise VisioUnavailable(
        "Failed to initialize Visio.Application after multiple attempts",
        details={"attempts": errors},
    )


def close_visio_app() -> None:
    """Best-effort shutdown registered with atexit.

    Closes any tracked documents first, then quits the application. We
    suppress errors aggressively because the interpreter is on its way
    out and COM cleanup at this stage is notoriously unreliable.
    """
    global _visio_app

    # Late imports to avoid circular dependencies at module load time.
    from .document import close_all_documents
    from .stencils import close_all_stencils

    close_all_documents()
    close_all_stencils()

    if _visio_app is not None:
        try:
            _visio_app.Quit()
        except Exception:
            pass
        _visio_app = None


atexit.register(close_visio_app)
