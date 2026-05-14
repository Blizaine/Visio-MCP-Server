"""
Exception hierarchy and the envelope decorator used by every MCP tool.

Every tool function in `visio_mcp_server.tools.*` returns a structured dict on
success or raises a subclass of `VisioMCPError` on failure. The `@envelope`
decorator catches those, plus a few well-known builtins, and serializes the
result into the standard JSON shape MCP clients consume:

    {"ok": true,  "data": <tool-specific>, "error": null}
    {"ok": false, "data": null,            "error": {"code": str, "message": str, "details": dict|null}}

Tools never build that JSON themselves — they raise or return data, and the
decorator handles the rest. This keeps the tool bodies focused on Visio work.
"""

from __future__ import annotations

import json
import logging
from functools import wraps
from typing import Any, Awaitable, Callable, Optional

logger = logging.getLogger("visio_mcp.tools")


class VisioMCPError(Exception):
    """Base class for all errors surfaced through the MCP envelope."""

    code: str = "INTERNAL"

    def __init__(self, message: str, details: Optional[dict] = None) -> None:
        super().__init__(message)
        self.details = details

    def to_envelope_error(self) -> dict:
        return {"code": self.code, "message": str(self), "details": self.details}


class VisioFileNotFound(VisioMCPError):
    code = "FILE_NOT_FOUND"


class ShapeNotFound(VisioMCPError):
    code = "SHAPE_NOT_FOUND"


class PageNotFound(VisioMCPError):
    code = "PAGE_NOT_FOUND"


class VisioUnavailable(VisioMCPError):
    code = "VISIO_UNAVAILABLE"


class ComError(VisioMCPError):
    code = "COM_ERROR"


class InvalidArgument(VisioMCPError):
    code = "INVALID_ARGUMENT"


def _success(data: Any) -> str:
    return json.dumps({"ok": True, "data": data, "error": None})


def _failure(code: str, message: str, details: Optional[dict] = None) -> str:
    return json.dumps({"ok": False, "data": None, "error": {"code": code, "message": message, "details": details}})


def envelope(tool_name: str) -> Callable[[Callable[..., Awaitable[Any]]], Callable[..., Awaitable[str]]]:
    """Wrap a tool coroutine so it returns the standard JSON envelope string.

    Initializes COM on the calling thread (some FastMCP transports dispatch
    onto a worker thread that hasn't called CoInitialize), logs entry/exit,
    and maps both VisioMCPError subclasses and a few builtins to error codes.
    """

    def decorator(fn: Callable[..., Awaitable[Any]]) -> Callable[..., Awaitable[str]]:
        @wraps(fn)
        async def wrapped(*args, **kwargs) -> str:
            # Lazy import: errors.py is loaded before com.app, so importing at
            # module scope would create a circular dependency.
            from .com.app import ensure_com_initialized
            ensure_com_initialized()
            logger.info("tool=%s call kwargs=%s", tool_name, _redact(kwargs))
            try:
                data = await fn(*args, **kwargs)
                logger.info("tool=%s ok", tool_name)
                return _success(data)
            except VisioMCPError as e:
                logger.warning("tool=%s code=%s message=%s", tool_name, e.code, e)
                return _failure(e.code, str(e), e.details)
            except FileNotFoundError as e:
                logger.warning("tool=%s code=FILE_NOT_FOUND message=%s", tool_name, e)
                return _failure("FILE_NOT_FOUND", str(e))
            except Exception as e:
                logger.exception("tool=%s code=INTERNAL", tool_name)
                return _failure("INTERNAL", f"{type(e).__name__}: {e}")

        # The inner tool body is annotated `-> dict` for readability, but the
        # wrapper actually returns the JSON envelope as a string. Patch the
        # annotation so FastMCP's schema introspection sees what it really gets.
        wrapped.__annotations__ = dict(getattr(fn, "__annotations__", {}))
        wrapped.__annotations__["return"] = str
        return wrapped

    return decorator


def _redact(kwargs: dict) -> dict:
    """Cap any oversized values so logs stay readable."""
    out = {}
    for k, v in kwargs.items():
        if isinstance(v, str) and len(v) > 200:
            out[k] = v[:200] + "..."
        else:
            out[k] = v
    return out
