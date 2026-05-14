"""
Visio MCP Server — entry point.

The actual tools live in `visio_mcp_server.tools.*` and register themselves
with the shared FastMCP instance on import. This module wires logging,
verifies Visio is installed, and starts the stdio transport.

Keeping the module name `visio_server` preserves the installed-script entry
point declared in pyproject.toml (`visio_mcp_server.visio_server:main`) and
the `python -m visio_mcp_server.visio_server` invocation in the README.
"""

import logging
import sys

from .com.app import check_visio_installed, get_visio_app
from .logging_setup import configure_logging
from .server_instance import mcp
from . import tools  # noqa: F401 — side-effect import registers @mcp.tool() functions

logger = logging.getLogger("visio_mcp.server")


def main() -> None:
    configure_logging()

    if not check_visio_installed():
        sys.stderr.write(
            "Microsoft Visio is not installed. This MCP server requires Visio to function.\n"
        )
        sys.exit(1)

    try:
        get_visio_app()
        logger.info("starting MCP stdio transport")
        mcp.run(transport="stdio")
    except Exception as e:
        logger.exception("fatal startup error")
        sys.stderr.write(f"Error initializing Visio MCP Server: {e}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
