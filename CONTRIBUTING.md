# Contributing

Thanks for your interest. This is a CTI fork of [GongRzhe/Office-Visio-MCP-Server](https://github.com/GongRzhe/Office-Visio-MCP-Server) tailored for AV system design workflows. We accept fixes and small features; large changes should start with an issue.

## Setup

```powershell
# One-time: uv (Astral) for Python tool/venv management
powershell -ExecutionPolicy Bypass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Per-clone
git clone https://github.com/Blizaine/Visio-MCP-Server.git
cd Visio-MCP-Server
uv venv
uv pip install -e .
```

## Running the end-to-end smoke test

The repo ships with `scripts/smoke_test.py` which spawns the server as a
subprocess over the MCP stdio transport and exercises every tool against
real Visio. **Requires Visio installed on your machine.**

```powershell
.\.venv\Scripts\python.exe scripts\smoke_test.py
```

Pass before opening a PR. The test will pop Visio up; expect ~5 seconds
of activity. Output files land in `%TEMP%`.

## Code layout

```
visio_mcp_server/
├── visio_server.py          # entry point: configure logging, start stdio
├── server_instance.py       # shared FastMCP instance
├── errors.py                # exception hierarchy + @envelope decorator
├── logging_setup.py
├── com/
│   ├── app.py               # Visio.Application singleton + CoInitialize
│   ├── document.py          # DocumentHandle + open-doc registry
│   └── undo.py              # undo_scope context manager
└── tools/
    ├── files.py             # create/open/close/save
    ├── pages.py             # page CRUD
    ├── shapes.py            # single-shape: add/connect/text/list
    ├── styling.py           # single-shape: fill/line/text_format + apply helpers
    ├── export.py            # export_page, export_pdf
    └── batch.py             # add_shapes, connect_shapes_bulk, style_shapes,
                             # delete_shapes, transform_shapes
```

Every tool returns the standard JSON envelope:
```json
{"ok": true,  "data": {...}, "error": null}
{"ok": false, "data": null,  "error": {"code": "...", "message": "...", "details": null}}
```

The `@envelope("tool_name")` decorator (in `errors.py`) handles
serialization, exception mapping, COM thread initialization, and
structured logging. Tool bodies just raise on failure and return raw data
on success.

## PR checklist

- [ ] `scripts/smoke_test.py` passes against real Visio on your machine
- [ ] If you added a tool, the smoke test exercises it
- [ ] Docstrings are clear (the model reads them; quality directly affects how
      the LLM picks the right tool)
- [ ] CHANGELOG.md has an entry under `[Unreleased]`
- [ ] If the public API changed, bumped the version in `pyproject.toml`

## Tool-design conventions

- Single-shape tools (`add_shape`, `set_shape_fill`, etc.) exist for
  incremental edits. Batch tools (`add_shapes`, `style_shapes`, etc.) are
  the recommended path for multi-shape operations.
- Mutating tools do NOT auto-save the document. Callers persist via
  `save_document` or rely on `close_document` (saves by default).
- Every mutating tool wraps its work in `undo_scope("...")` so partial
  failures roll back cleanly.
- Page targeting: every shape-level tool accepts an optional `page_name`
  parameter; `None` defaults to the active page on the relevant document.

## Reporting issues

Open one at <https://github.com/Blizaine/Visio-MCP-Server/issues>. Include:
- Visio version (Help > About in Visio)
- Python version (`python --version`)
- Package version (`uv tool list` if installed via uv)
- The full envelope you got back, especially the `error` block
- The MCP client you're using (Claude Desktop, Claude Code, Cline, etc.)
- A minimal reproduction
