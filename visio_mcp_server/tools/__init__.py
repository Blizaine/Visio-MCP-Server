"""
Importing this package triggers tool registration as a side effect.

Each submodule's `@mcp.tool()` decorators run at import time, registering
their tools with the shared FastMCP instance in `server_instance`.
"""

from . import export  # noqa: F401 — side-effect imports register the tools
from . import files  # noqa: F401
from . import pages  # noqa: F401
from . import shapes  # noqa: F401
from . import styling  # noqa: F401
from . import batch  # noqa: F401
