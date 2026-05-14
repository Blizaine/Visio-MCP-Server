"""
Importing this package triggers tool registration as a side effect.

Each submodule's `@mcp.tool()` decorators run at import time, registering
their tools with the shared FastMCP instance in `server_instance`.
"""

from . import files  # noqa: F401 — side-effect imports register the tools
from . import shapes  # noqa: F401
