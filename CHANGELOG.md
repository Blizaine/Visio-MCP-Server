# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added (Phase 3 — page support)
- `list_pages(file_path)` — returns each page's index, name, dimensions
  (inches), and background flag.
- `add_page(file_path, name?, width?, height?, background?)` — appends a
  page. Optional name, dimensions in inches, and a `background` flag for
  backdrop pages.
- `delete_page(file_path, page_name)` — refuses to delete the last
  remaining page (Visio requires at least one).
- `set_active_page(file_path, page_name)` — makes a page the active one in
  Visio's window; locates the right window for the document first so this
  works in multi-document workflows.
- `duplicate_page(file_path, source_page_name, new_name?)` — copies a page
  including shapes and properties; optionally renames the duplicate.
- Every existing shape tool (`add_shape`, `connect_shapes`, `add_text`,
  `list_shapes`) now accepts an optional `page_name`. When omitted the
  active page is used.

### Changed (Phase 3)
- `DocumentHandle.active_page` (property) replaced with
  `DocumentHandle.get_page(page_name=None)`. Default behavior also fixes a
  latent bug: previously the tools always used `Application.ActivePage`,
  which is global — in a multi-document workflow this could edit the
  wrong document's page. The new default returns the active page only
  when it belongs to *this* document, falling back to the document's
  first page otherwise.
- All shape-tool responses now include a `page_name` field so callers can
  confirm which page their operation targeted.
- Package version 2.0.0 → 2.1.0 (additive — new tools + new optional
  parameters; no breaking changes).
- New error code `PAGE_NOT_FOUND` for `page_name` lookups that miss.

### Added (sharing prep)
- `BOOTSTRAP.md` — colleague-facing install guide. Covers prerequisites,
  the uv-based automated install, manual step-by-step fallback, MCP client
  configuration snippets for Claude Desktop and Claude Code, install
  sources (git/PyPI/wheel/local), troubleshooting, and uninstall.
- `scripts/bootstrap.ps1` — automated installer. Takes a `-Source`
  parameter and runs end-to-end: verifies Visio, installs uv if missing,
  installs CPython 3.12 via uv, `uv tool install`s the package, and prints
  the MCP client config snippet. Does not modify the MCP client config
  file directly (paste-it-yourself is safer than auto-merging JSON).
- `scripts/smoke_test.py` (committed earlier in bc7ccf0) — end-to-end test
  that drives every tool against a real Visio. Doubles as a colleague
  verification step.

### Changed
- README's Installation section now recommends the uv-based path and
  points at `BOOTSTRAP.md` for the colleague-facing rollout. The old
  `pip install pywin32` flow is removed.

### Changed — BREAKING (Phase 2 — architecture refactor)
- **Response format**: every tool now returns a JSON-encoded envelope
  `{"ok": bool, "data": ..., "error": {"code", "message", "details"} | null}`.
  Previously tools returned ad-hoc plain strings like `"Visio file created
  successfully at: ..."`. Clients that parsed those strings must switch to
  reading the envelope. See the new "Response Format" section in the README.
- **Package version**: 1.0.1 → 2.0.0 to reflect the breaking response shape.
- **Module layout**: the 465-line `visio_server.py` monolith is split into
  focused modules:
  - `visio_mcp_server.server_instance` — the shared FastMCP instance
  - `visio_mcp_server.errors` — exception hierarchy + `@envelope` decorator
  - `visio_mcp_server.logging_setup` — stderr logging configuration
  - `visio_mcp_server.com.app` — Visio application lifecycle + CoInitialize
  - `visio_mcp_server.com.document` — `DocumentHandle` + open-doc registry
  - `visio_mcp_server.com.undo` — `undo_scope` context manager
  - `visio_mcp_server.tools.files` — create/open/close
  - `visio_mcp_server.tools.shapes` — add/connect/text/list
  `visio_mcp_server.visio_server:main` is preserved as the entry point so
  the installed-script command and `python -m` invocation don't change.

### Added (Phase 2)
- `undo_scope(name)` context manager wraps every mutating tool. Partial
  failures inside the block trigger `EndUndoScope(commit=False)`, rolling
  the operation back. Clean exits commit normally.
- Structured logging to stderr via the `visio_mcp` logger hierarchy. Level
  is configurable through the `VISIO_MCP_LOG_LEVEL` environment variable
  (default `INFO`). Stdout is never written to so the MCP stdio protocol
  is unaffected.
- `list_shapes` now returns shape positions/sizes in inches explicitly
  (`Cell.Result("in")` instead of `Result("")`) and includes them as
  numeric floats rather than strings.

### Fixed (Phase 1 — correctness)
- `add_shape`, `connect_shapes`, `add_text`, and `list_shapes` no longer call
  each other through the `@mcp.tool()`-decorated wrappers. Business logic is
  now in plain private helpers (`_create_document`, `_open_document`,
  `_ensure_document_open`, `_find_shape`) and tools call helpers directly.
  The previous pattern relied on `await`-ing decorated tool functions and
  string-matching `"Error" in result`, which was fragile and could break
  depending on the MCP framework version.
- `connect_shapes` no longer sets `LinePattern = "0"` when "Straight" or
  "Curved" is selected — that formula makes the connector line invisible.
  "Straight" now writes `ShapeRouteStyle = 2` (visLORouteStraight) instead.
  "Curved" is accepted but currently routes as Dynamic; true curved-connector
  support is deferred to Phase 4 (styling).
- `open_documents` dict is now keyed on a canonicalized path
  (`os.path.normcase(os.path.abspath(p))`) so `C:\foo.vsdx` and `c:/foo.vsdx`
  share one entry. Cached document handles are liveness-probed before reuse.
- Bare `except:` blocks replaced with `except Exception:` / `except OSError:`.
  Errors now surface real messages instead of generic "Failed" strings.
- Every tool calls `pythoncom.CoInitialize()` on entry so the COM apartment is
  set up on whatever thread FastMCP dispatched the call on. Safe to call
  repeatedly on the same thread.

### Changed (Phase 0 — hygiene)
- Raised minimum Python version to 3.10 (was inconsistently 3.8 / 3.12).
- README now points at the correct module entry point (`python -m visio_mcp_server.visio_server`) and uses `visio-server` as the MCP server key (was incorrectly `ppt`).

### Removed (Phase 0 — hygiene)
- Stale `mcp-config.json` tracked from upstream (contained the original author's machine paths). The file is now generated by `setup_mcp.py` and gitignored.
- Unused imports (`glob`, `tempfile`, `typing.Dict`, `typing.Any`, `typing.List`) from `visio_mcp_server/visio_server.py`.
- "Export diagrams as images" claim from the README's Features section — that tool does not exist yet and remains on the roadmap.

### Notes
- Tool signatures and return-string formats are preserved. No API surface change.
- Phase 1 verification is currently structural-only; the host machine does not
  have Python installed yet, so `py_compile` and a real Visio smoke test have
  not been run. Recommended to do both after `pip install -e .`.
