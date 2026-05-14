"""
The single FastMCP instance the package shares.

Lives in its own module so the tools submodules can `from ..server_instance
import mcp` and register themselves with `@mcp.tool()` without creating a
circular import via `visio_server.py`.
"""

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("visio-server")
