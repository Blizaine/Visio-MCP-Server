# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added (Phase 12 — layers and groups, v3.6.0)
AV system drawings routinely use both: separate "Audio Routing",
"Video Routing", "Control" layers so engineers can toggle visibility
during review; groups for racks / single-room subsystems so a whole
subsystem moves and styles as a unit.

Layers (`visio_mcp_server/tools/layers.py`):
- `list_layers(file_path, page_name?)` — name, index, visibility,
  print, active, locked, color formula for each layer on the page.
- `add_layer(file_path, name, page_name?, visible?, printable?,
  locked?, color?)` — name must be unique on the page.
- `delete_layer(file_path, name, page_name?, delete_shapes?)` —
  `delete_shapes=False` (default) keeps the shapes on the page; True
  removes them too.
- `set_layer_properties(file_path, name, page_name?, visible?,
  printable?, locked?, color?)` — partial update; pass only what you
  want to change.
- `set_shapes_layers(file_path, assignments=[{shape_id, layers}], page_name?)`
  — batch REPLACE semantics. The shape's layer membership is replaced
  with exactly the listed layers (empty list removes the shape from
  every layer). Pre-validates layer names; aborts the whole batch with
  `INVALID_ARGUMENT` if any name is missing.

Groups (`visio_mcp_server/tools/groups.py`):
- `group_shapes(file_path, shape_ids, page_name?)` — needs at least
  2 shape IDs; returns `{group_id, member_ids, member_count}`.
  Existing tools (`transform_shapes`, `style_shapes`,
  `delete_shapes`) operate on the group's ID since a group IS a shape.
- `ungroup_shape(file_path, group_id, page_name?)` — breaks a group
  back into its constituents. Returns just `{ungrouped_id}`; use
  `list_shapes` to enumerate the page afterwards.

### Tools learned along the way

Two Visio-COM nuances that bit us during Phase 12 development:
- **Layer cells via `Layer.CellsU(name)` aren't reliable across builds.**
  Some Visio installations don't dispatch the named accessor on Layer
  objects. We route through `Page.PageSheet.CellsSRC(visSectionLayer=5,
  row=layer.Index-1, column=<int>)` instead, which always works.
- **`Selection.Select(shape, visSelect)` uses VisActionCodes**, where
  `visSelect = 1` (NOT 2 — that's `visSubSelect`). Using the wrong
  value causes `Group()` to grab the entire connected component (any
  shape reachable via gluing) instead of just the listed shapes.
  Calling out the trap so it's not re-introduced.
- **`Shape.AddLayer` / `Shape.RemoveLayer` don't dispatch in all
  Visio builds.** We use `Layer.Add(shape, 0)` / `Layer.Remove(shape, 0)`
  via the Layer object instead.

### Changed (Phase 12)
- Package version 3.5.0 → 3.6.0 (additive; no breaking changes).
- Tool count 38 → 45.

### Verification
Smoke test extended: 3 layers (Video Routing red, Audio Routing
orange, Network blue) added on the batch page, 6 shapes assigned
across them (one shape on 2 layers), Network layer hidden via
`set_layer_properties`, bogus layer name correctly rejected. 3 shapes
grouped (group_id=14), ungrouped, original members verified back on
page via `list_shapes`.

### Changed (Phase 11 — AV-grade connectors, v3.5.0)
Signal-typing for AV diagrams: connectors now carry visual semantics.
Yellow lines for audio, red for video, green for control, blue for
network — whatever your team standardizes on — encoded per-connection
in the same batch call that creates them.

- `connect_shapes` gains 5 new optional parameters:
  - `label` — text the connector displays midway along its route.
  - `color` — line color (same parser as `set_shape_line`).
  - `weight` — line thickness in points.
  - `pattern` — dash pattern (0=no line, 1=solid, 2-23=dashes).
  - `connector_master` — `{"stencil": str, "master": str}` to use a
    stencil-defined connector master (branded cable masters etc.)
    instead of the default dynamic connector.
  Response gains `applied_style` showing what was applied.
- `connect_shapes_bulk` — each entry in `connections` accepts the same
  five fields, so an entire AV signal flow's mixed signal types land
  in a single batch call. Pre-opens every distinct `connector_master`
  stencil before any drops so typos fail fast.
- Internal: `drop_connector(page, connector_master?, stencil_doc_cache?)`
  helper in `tools/shapes.py` centralizes connector creation (default
  vs. stencil-master). `apply_connector_style()` in `tools/styling.py`
  reuses `apply_line` so connector styling is identical to shape
  styling. Both helpers shared by single + batch tools.
- Package version 3.4.0 → 3.5.0 (additive; no breaking changes).
- Tool count unchanged at 38 — both new capabilities ride on existing
  tools.

### Verification
Smoke test extended to drive a 6-edge AV-style signal flow:
- 2 video runs (red, 1.5pt, label="HDMI")
- 1 control run (green, dashed, label="Cresnet")
- 2 audio runs (orange, 1pt, label="Audio")
- 1 network run (blue, 2pt, label="1Gb Ethernet")
All 6 connectors reported applied_style; labels visible in `list_shapes`
output. Batch wall-clock for 7 shapes + 6 styled connectors + 7 style
applies remains 0.42s.

### Added (Phase 10 — template discovery, v3.4.0)
Closes the third gap in the discovery story: stencils (Phase 7),
masters (Phase 8), shape data (Phase 9), and now branded templates.
AV diagrams typically start from a company-standard titleblock template;
this phase makes that template addressable by name instead of by
machine-specific absolute path.

- `visio_mcp_server/com/templates.py` — file-system walker + on-disk
  index. Unlike stencils (which need a real Visio open to enumerate
  masters), templates only need file metadata, so the index is cheap —
  sub-second on hundreds of files. Same cache pattern as stencils:
  `%USERPROFILE%/.cti-visio-mcp/template_index.json` with per-file
  mtime invalidation.
  - Path discovery: env var `CTI_VISIO_TEMPLATE_PATHS` (semicolon-separated)
    + Visio's `Application.TemplatePaths` + Windows defaults.
  - Recognized extensions: `.vstx`, `.vst`, `.vstm`, `.vsdx`, `.vsd`, `.vsdm`.
    Reality is that a `.vsdx` "starter file" is often used as a template
    — `Documents.Add(path)` does the right thing for either format.
  - Category = immediate parent directory name (for grouping by, e.g.,
    `Signal Flow/`, `Org Chart/`).
- `visio_mcp_server/tools/templates.py` — 4 new tools:
  - `list_templates(category?, limit?, offset?)`
  - `find_templates(query, limit?)` — exact > substring > fuzzy
  - `template_index_status()`
  - `reindex_templates(force?)`

### Changed (Phase 10)
- `create_visio_file` now accepts an optional `template: str` argument
  that resolves a template name against the index. Mutually exclusive
  with the existing `template_path`. Both omitted = blank document
  (unchanged). Response gains `template_used: str|null` so callers can
  confirm which template seeded the new file.
- Package version 3.3.0 → 3.4.0 (additive; no breaking changes).
- Tool count: 34 → 38.

### Verification
Smoke test pass against the test template directory:
`CTI Visio Signal Flow Template.vsdx` indexed, found by
`find_templates("CTI")` with score 800, and used as the seed for a new
document via `create_visio_file(template=...)`.

### Added (Phase 9 — shape data tools, v3.3.0)
The third piece of the stencil story: read, write, and search the
custom-property fields ("Prop.*" cells in the ShapeSheet) that turn a
placed master from a decorative icon into an information-bearing
diagram element. A Cisco mic master has fields like Manufacturer,
DeviceName, SerialNumber — those are how a network/AV diagram becomes
self-documenting.

- `get_shape_data(file_path, shape_id, page_name?)` — full property
  dump. Each property returns `{value, label, prompt, type, type_id,
  formula}` so the caller knows both the user-visible label and the
  underlying typed value. Number/Currency come back as floats, Boolean
  as Python bool, strings/lists/dates as formatted display strings.
- `set_shape_data(file_path, shape_id, data={prop: value, ...}, page_name?)`
  — all-or-nothing write. If any property doesn't exist on the shape,
  the call fails with `INVALID_ARGUMENT`, `error.details.missing` lists
  the names, and the undo scope rolls back any partial writes that
  landed first.
- `set_shapes_data(file_path, updates=[{shape_id, data}, ...], page_name?)`
  — batch version. Same all-or-nothing semantics across the whole batch.
- `find_shapes_by_data(file_path, query={prop: value, ...}, page_name?, limit?)`
  — AND-match search on the page. Strings are matched case-insensitively
  as substrings; numbers/booleans are exact. "Find every device whose
  Manufacturer contains 'Cisco'" is two property names away.

### Changed (Phase 9)
- `list_shapes` gains an optional `include_data: bool = False` parameter
  — when True each shape's `data` is included in the response (same
  shape as `get_shape_data`'s output). Default off to keep responses
  small on big pages.
- `drop_master` and `drop_masters` gain an optional `strict_data: bool =
  False` parameter. Default behavior (silent skip on unknown property
  names) is unchanged; `strict_data=True` routes through the same
  validation as `set_shape_data` so a typo in the inline data dict
  rolls the drop back instead of silently dropping the wrong-named
  property.
- Internal: `_format_prop_value` lifted out of `tools/drop.py` into
  `com/document.py` as `format_prop_value` (public), and a new
  `read_shape_data(shape)` helper centralizes ShapeSheet Prop-section
  parsing. Both `drop.py`, `shape_data.py`, and the enhanced
  `list_shapes` share one source of truth.
- Package version 3.2.0 → 3.3.0 (additive; no breaking changes).

### Verification
Smoke test against the Cisco stencil: dropped 3 masters, scanned for
one with custom properties (shape 42 had 9 fields), round-tripped
`DeviceName` (`"MNT"` → `"SMOKETEST-..."` → read back matches),
verified `find_shapes_by_data` located the modified shape, and verified
strict mode rejects a bogus property name with `INVALID_ARGUMENT`.

### Added (Phase 8 — drop master shapes, v3.2.0)
The payoff for Phase 7. Once the stencil index knows what masters exist,
this phase actually places them onto pages. `add_shape(s)` draws generic
primitives; `drop_master(s)` drops real stencil shapes (Crestron NVX
codecs, Cisco mics, Extron switchers, etc.) with their built-in
geometry, connection points, and shape data fields.

- `drop_master(file_path, stencil, master, x, y, page_name?, width?,
  height?, text?, data?)` — single drop. Optional inline property writes
  via `data: {prop_name: value, ...}` write to ShapeSheet `Prop.<name>`
  cells inherited from the master. Unknown property names skip silently
  (Phase 9 will add proper discovery and validation).
- `drop_masters(file_path, items=[...], page_name?)` — batch drop.
  Strongly preferred when placing more than one master. Pre-opens every
  distinct stencil before any drops so typos fail fast, disables Visio's
  screen redraw, wraps in one undo scope.

- `visio_mcp_server/com/stencils.py` gains `get_or_open_stencil()` and
  `close_all_stencils()`. A per-process cache of opened stencil
  Documents (hidden + read-only) so subsequent drops from the same
  stencil cost ~50ms each instead of ~500ms (each cold open). Cleanup is
  hooked into the existing `close_visio_app` atexit handler.

### Changed (Phase 8)
- `add_shape` and `add_shapes` docstrings now explicitly steer the model
  toward `drop_master`/`drop_masters` when the diagram involves
  manufacturer-specific equipment. Generic primitives remain valid for
  abstract sketches, labels, and zone backgrounds.
- Package version 3.1.0 → 3.2.0 (additive; no breaking changes).

### Performance (Phase 8 baseline)
3 Cisco masters dropped (`CS-CODEC-EQ-NRK9++`, `CS-CODEC-EQ-RCK`,
`CS-MIC-TABLE-J=`) in a single `drop_masters` call: **0.30s** COM-side.
A typical AV diagram (15-30 devices) should drop in well under 1s of
Visio work.

### Added (Phase 7 — stencil indexing and search, v3.1.0)
Foundation for "real" Visio diagrams. The existing `add_shape(s)` calls
draw generic geometric primitives; this phase lets the model discover and
reference manufacturer stencils (Crestron, Cisco, Extron, etc.) so the
next phase (drop_master) can produce diagrams that look like actual AV
system drawings instead of boxes-and-arrows abstractions.

- `visio_mcp_server/com/stencils.py` — `StencilIndex` foundation:
  - Path discovery: env var `CTI_VISIO_STENCIL_PATHS` (semicolon-separated),
    plus Visio's `Application.StencilPaths`, plus
    `%USERPROFILE%\Documents\My Shapes`. Dedupes; only existing dirs make
    the final list.
  - Walks `.vss`/`.vssx`/`.vssm` files and indexes each by opening it
    hidden + read-only + no-workspace, harvesting masters
    (`name`, `base_id`, `prompt`, `prop_names`).
  - **Macro safety**: sets `Application.AutomationSecurity = 3`
    (`msoAutomationSecurityForceDisable`) before opening any stencil so
    VBA macros in `.vssm` files never run during the index pass; restores
    the prior value when done.
  - On-disk cache at `%USERPROFILE%\.cti-visio-mcp\stencil_index.json`.
    Per-file mtime invalidation — unchanged stencils are reused on
    subsequent rebuilds, only modified files reopen in Visio.
  - Manufacturer extracted from filename stem (e.g. `Crestron NVX.vssx`
    -> `Crestron`).
- `visio_mcp_server/tools/stencils.py` — 5 new tools:
  - `list_stencils(manufacturer?, limit?, offset?)` — paginated listing.
  - `list_masters(stencil, query?, limit?)` — masters in one stencil.
  - `find_masters(query, manufacturer?, limit=10)` — rank-search across
    every indexed master. Ranking: exact match > substring (earlier =
    better) > token overlap on name+manufacturer+prompt > fuzzy
    similarity. Tuned for "Crestron DM-NVX-360"-style queries.
  - `stencil_index_status()` — paths scanned, counts, in-progress state,
    per-file errors. Use this when `find_masters` returns nothing.
  - `reindex_stencils(force?)` — admin/maintenance.

### Changed (Phase 7)
- Package version 3.0.0 → 3.1.0 (additive; no breaking changes).
- `scripts/bootstrap.ps1` now actively refuses to run if `claude.exe`
  processes are detected, with a clear error pointing at File > Exit and
  Task Manager. This prevents the half-broken-install state we hit on
  the first colleague update attempt — root cause was the MS Store
  app's package virtualization locking the launcher .exe during install.
- `scripts/smoke_test.py` extended with stencil-index coverage. Indexes
  the configured test directory (`CTI_VISIO_STENCIL_PATHS` defaults to
  `C:\Users\blaine.brown\Documents\AI Testing\VisioMCP\Visio_Stencils`
  in the smoke test), validates counts, picks the first master out of
  the first non-empty stencil, and verifies it ranks itself first in a
  `find_masters` call.

### Performance (Phase 7 baseline)
Indexing 15 stencils with 634 masters (Cisco, Crestron, Crown, Extron,
JBL, LG, Logitech, NEAT, Planar, QSC, SAMSUNG, Shure, SONY, Wattbox,
Favorites): **14.3 seconds** cold. Extrapolating to a ~200-stencil
library: roughly 3 minutes for first index, ~milliseconds for cache hits
on subsequent server starts.

### Changed — BREAKING (Phase 6 — rename and first release)
- **Package renamed**: `office-visio-mcp-server` → `cti-visio-mcp-server`.
  The upstream PyPI name is owned by the original author (who has
  archived their repo); the new name avoids that conflict and reflects
  the CTI-specific direction this fork is going (AV system design,
  company stencil/template libraries).
- **Version bumped to 3.0.0.** The Python import path
  (`visio_mcp_server`) and the launcher binary name (`visio_mcp_server.exe`)
  are unchanged on purpose — existing MCP client configs that reference
  the .exe path keep working.
- LICENSE updated with a CTI/Blaine Brown copyright line for the fork's
  substantial modifications. Stays MIT.
- `setup_mcp.py` (legacy interactive installer carried over from
  upstream) removed. `scripts/bootstrap.ps1` + `BOOTSTRAP.md` are now
  the colleague-facing install path.

### Added (Phase 6)
- `CONTRIBUTING.md` — repo conventions, smoke-test workflow, PR checklist.
- `.github/ISSUE_TEMPLATE/` — `bug_report.md` and `feature_request.md`
  templates that capture the env info we'll need to triage real reports.
- Pre-emptive `.gitignore` for `.cti-visio-mcp/` — the on-disk stencil-
  index cache directory that Phase 7 will write to.

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
