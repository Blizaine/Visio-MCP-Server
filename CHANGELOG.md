# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added (Phase 5 — batch tools and modify-existing CRUD)
- `add_shapes(file_path, shapes=[...], page_name?)` — bulk shape creation.
  Each item: `{shape_type, x, y, width, height, text?}`.
- `connect_shapes_bulk(file_path, connections=[...], page_name?)` — bulk
  connectors. Each item: `{shape1_id, shape2_id, connector_type?}`.
- `style_shapes(file_path, updates=[...], page_name?)` — bulk fill / line /
  text-format / text changes. Each item bundles whatever you want to
  change on one shape: `{shape_id, fill?, line?, text?, text_format?}`.
- `delete_shapes(file_path, shape_ids=[...], page_name?)` — bulk delete.
  All-or-nothing: if any target is missing, the batch aborts and the undo
  scope rolls back.
- `transform_shapes(file_path, updates=[...], page_name?)` — bulk move /
  resize / rotate. Each item: `{shape_id, x?, y?, width?, height?,
  angle_degrees?}`. Partial updates supported.
- `save_document(file_path)` — explicit save, replacing the auto-save that
  used to fire after every mutating tool.

### Changed (Phase 5)
- **Mutating tools no longer auto-save** (`add_shape`, `connect_shapes`,
  `add_text`, `add_page`, `delete_page`, `duplicate_page`, `set_shape_fill`,
  `set_shape_line`, `set_shape_text_format`). Saving the .vsdx (re-zipping
  XML and writing to disk) was ~100-300ms per call; for multi-step diagrams
  that added significant latency. Call `save_document` explicitly when you
  want persistence, or rely on `close_document` (still saves by default).
  This is a semantic change but doesn't affect the response envelope, so
  it's a minor bump.
- Each batch tool wraps its work in a single `undo_scope` and disables
  `Application.ScreenUpdating` for the duration — Visio doesn't repaint
  between operations, dropping per-op latency further.
- Docstrings on all single-shape tools (`add_shape`, `connect_shapes`,
  `add_text`, `set_shape_fill`, `set_shape_line`, `set_shape_text_format`)
  now explicitly direct callers toward the batch variants when handling
  more than one shape.
- Styling apply logic factored into reusable helpers (`apply_fill`,
  `apply_line`, `apply_text_format`) in `tools/styling.py` so the single-
  shape tools and the `style_shapes` batch share one source of truth.
- Package version 2.2.0 → 2.3.0.

### Performance
Smoke-test measurement: building 7 shapes + 6 connectors + 7 style updates
(20 operations) via batch tools takes ~0.34s of COM-side work. Previously
the equivalent ~20 single-tool calls would have spent ~5 minutes on model
round-trips alone (in addition to the COM work).

### Added (Phase 4 — export and styling)
- `export_page(file_path, output_path, page_name?)` — single-page export.
  Format inferred from the output extension: `.png`, `.jpg`/`.jpeg`,
  `.gif`, `.bmp`, `.tif`/`.tiff`, `.svg`, `.emf`, `.wmf`. Returns
  `{output_path, page_name, format, bytes}`.
- `export_pdf(file_path, output_path, page_range?)` — multi-page PDF
  export via `Document.ExportAsFixedFormat`. Range accepts `None`/`"all"`,
  `"current"`, `"N"`, or `"N-M"`. Returns `{output_path, page_range, bytes}`.
- `set_shape_fill(file_path, shape_id, color, pattern?, page_name?)` —
  writes `FillForegnd` and `FillPattern` cells.
- `set_shape_line(file_path, shape_id, color?, weight?, pattern?, page_name?)`
  — partial updates; weight is in points. Writes `LineColor`, `LineWeight`,
  `LinePattern`.
- `set_shape_text_format(file_path, shape_id, font?, size?, color?, bold?,
  italic?, underline?, align?, page_name?)` — partial updates. `bold`,
  `italic`, `underline` toggle only their bits on `Char.Style` so other
  styles aren't clobbered. Writes `Char.Font`, `Char.Size`, `Char.Color`,
  `Char.Style`, and `Para.HorzAlign`.
- Color parser accepts `#RGB`, `#RRGGBB`, `RGB(r,g,b)`, `rgb(r,g,b)`, and a
  small set of named colors (red/blue/green/.../pink/brown). All are
  normalized to the `RGB(r,g,b)` formula written into the ShapeSheet.

### Changed (Phase 4)
- `_find_shape` (private to `tools/shapes.py`) replaced with a shared
  `find_shape_on_page` helper in `com/document.py`. Now uses
  `Shapes.ItemFromID(id)` — O(1) via Visio's internal index — instead of
  iterating `page.Shapes` linearly.
- Package version 2.1.0 → 2.2.0 (additive; no breaking changes).

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
