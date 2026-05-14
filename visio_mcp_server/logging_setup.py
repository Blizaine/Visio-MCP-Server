"""
Configure logging to stderr.

Stdout is reserved for the MCP JSON-RPC protocol on stdio transports, so we
must never log there. The log level is controlled by the VISIO_MCP_LOG_LEVEL
environment variable (default INFO).
"""

import logging
import os
import sys


def configure_logging() -> None:
    level_name = os.environ.get("VISIO_MCP_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    root = logging.getLogger("visio_mcp")
    if root.handlers:
        return  # idempotent: harmless to call twice

    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%dT%H:%M:%S",
    ))
    root.addHandler(handler)
    root.setLevel(level)
    root.propagate = False
