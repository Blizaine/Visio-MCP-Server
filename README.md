# Visio MCP Server

A MCP server that provides tools for creating and editing Microsoft Visio diagrams programmatically via a standardized API.

![](https://badge.mcpx.dev?type=server "MCP Server")

## Overview

Visio MCP Server allows you to automate Visio diagram creation and editing using Python. It leverages Microsoft's COM interface to control Visio, enabling you to programmatically create diagrams, add shapes, connect them, add text, and more.

## Example

![demo](./public/demo.gif)

## Requirements

- Windows operating system
- Microsoft Visio (Professional or Standard) installed
- Python 3.10+
- Python packages:
  - `mcp.server`
  - `win32com.client` (pywin32)

## Installation

The recommended path uses `uv` (Astral's Python tool manager) — no admin
rights and no pre-installed Python required. See
**[BOOTSTRAP.md](BOOTSTRAP.md)** for the colleague-facing full guide, plus
troubleshooting for locked-down corporate environments.

**Automated (from a clone of this repo):**

```powershell
scripts\bootstrap.ps1 -Source "<install-source>"
```

where `<install-source>` is the git URL, wheel path, PyPI name, or `.` for
a local editable install. The script handles uv, Python, and the package
itself; it prints the MCP client config snippet at the end.

**Local dev install (you have the source cloned):**

```powershell
# One-time: install uv
powershell -ExecutionPolicy Bypass -c "irm https://astral.sh/uv/install.ps1 | iex"

# Per-project
uv venv
uv pip install -e .
```

Run the server directly:

```powershell
.\.venv\Scripts\python.exe -m visio_mcp_server.visio_server
```

Verify the install with the included smoke test:

```powershell
.\.venv\Scripts\python.exe scripts\smoke_test.py
```

## Features

The server provides the following functionality:

### Creating and Opening Files
- Create new Visio diagrams
- Open existing Visio diagrams

### Shape Management
- Add various shapes (Rectangle, Circle, Line, etc.)
- Connect shapes with different connector types
- Add text to shapes
- List all shapes in a document

### Page Management
- List, add, delete, duplicate, and activate pages
- Every shape tool accepts an optional `page_name` to target a specific page

### File Operations
- Save documents to specified locations
- Close documents safely
- Export pages to images (PNG, JPG, SVG, BMP, TIF, EMF, WMF)
- Export documents to PDF (full doc, current page, single page, or range)

### Shape Styling
- Set fill color and pattern
- Set outline color, weight (in points), and dash pattern
- Set text font, size, color, bold/italic/underline, and horizontal alignment
- Color input accepts `#RRGGBB`, `RGB(r,g,b)`, or named colors

## MCP Configuration

### Option 1: Local Python Server

Add the server to your MCP settings configuration file:

```json
{
  "mcpServers": {
    "visio-server": {
      "command": "python",
      "args": ["-m", "visio_mcp_server.visio_server"],
      "env": {}
    }
  }
}
```

### Option 2: Using UVX (No Local Installation Required)

If you have `uvx` installed, you can run the server directly from PyPI without local installation:

```json
{
  "mcpServers": {
    "visio-server": {
      "command": "uvx",
      "args": [
        "--from", "office-visio-mcp-server", "visio_mcp_server"
      ]
    }
  }
}
```

## Response Format

Every tool returns a JSON string with a consistent envelope so clients can
parse results uniformly:

```json
{ "ok": true,  "data": { /* tool-specific */ }, "error": null }
{ "ok": false, "data": null, "error": { "code": "FILE_NOT_FOUND", "message": "...", "details": null } }
```

Error codes currently emitted:

| Code | When |
| --- | --- |
| `FILE_NOT_FOUND` | The Visio file at the given path doesn't exist (and the tool doesn't auto-create). |
| `SHAPE_NOT_FOUND` | A requested shape ID isn't on the target page. |
| `PAGE_NOT_FOUND` | A requested `page_name` isn't in the document. |
| `VISIO_UNAVAILABLE` | Visio couldn't be launched (not installed, COM blocked, etc.). |
| `COM_ERROR` | A Visio COM call failed during a save/open/create. |
| `INVALID_ARGUMENT` | Reserved for future validation; not yet emitted by current tools. |
| `INTERNAL` | Uncaught exception. The `message` includes the original exception type. |

> **Breaking change (v2.0.0)**: prior versions returned plain strings like
> `"Visio file created successfully at: ..."`. Clients that parsed those
> strings need to switch to reading `data.path` on the envelope.

## API Reference

### Create a Visio File
Creates a new Visio diagram.

```json
{
  "template_path": "[optional] Path to Visio template (.vstx, .vst)",
  "save_path": "[optional] Where to save the file"
}
```

Example:
```json
{
  "save_path": "C:\\Users\\YourUsername\\Documents\\MyDiagram.vsdx"
}
```

### Open a Visio File
Opens an existing Visio diagram.

```json
{
  "file_path": "Path to the Visio file to open"
}
```

### Add Shape
Adds a shape to a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "shape_type": "Type of shape (Rectangle, Circle, Line, etc.)",
  "x": 1.0,
  "y": 1.0,
  "width": 1.0,
  "height": 1.0
}
```

### Connect Shapes
Connects two shapes in a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "shape1_id": 1,
  "shape2_id": 2,
  "connector_type": "Dynamic, Straight, or Curved"
}
```

### Add Text
Adds text to a shape in a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "shape_id": 1,
  "text": "Text to add to the shape"
}
```

### List Shapes
Lists all shapes on a page of a Visio diagram.

```json
{
  "file_path": "Path to the Visio file",
  "page_name": "[optional] Name of the page; defaults to the active page"
}
```

### Page Tools

All shape tools above accept an optional `"page_name"` argument; omit it
to target the active page. In addition:

- **`list_pages`** — `{"file_path": "..."}` → returns `{"pages": [{"index", "name", "width", "height", "background"}]}`.
- **`add_page`** — `{"file_path", "name"?, "width"?, "height"?, "background"?}`. Width/height in inches; if omitted, uses the document defaults.
- **`delete_page`** — `{"file_path", "page_name"}`. Refuses to delete the last remaining page.
- **`set_active_page`** — `{"file_path", "page_name"}`. Activates the page in Visio's window.
- **`duplicate_page`** — `{"file_path", "source_page_name", "new_name"?}`. Copies the page including all shapes.

### Styling Tools

Set visual properties on existing shapes. Color values accept
`#RRGGBB`, `#RGB`, `RGB(r,g,b)`, or named colors (red, blue, green,
yellow, cyan, magenta, gray, orange, purple, pink, brown, black, white).

- **`set_shape_fill`** — `{"file_path", "shape_id", "color", "pattern"?, "page_name"?}`. Pattern 0=none, 1=solid (default), 2+=various hatches.
- **`set_shape_line`** — `{"file_path", "shape_id", "color"?, "weight"?, "pattern"?, "page_name"?}`. Weight in points; pattern 0=no line, 1=solid, 2-23=dashes. Partial updates supported — pass only what you want to change.
- **`set_shape_text_format`** — `{"file_path", "shape_id", "font"?, "size"?, "color"?, "bold"?, "italic"?, "underline"?, "align"?, "page_name"?}`. Size in points. `align` accepts "left", "center", "right", or "justify". Partial updates supported; bold/italic/underline toggle individually without disturbing the others.

### Export Tools

- **`export_page`** — `{"file_path", "output_path", "page_name"?}`. Single-page export; format inferred from the `output_path` extension (.png, .jpg, .jpeg, .gif, .bmp, .tif, .tiff, .svg, .emf, .wmf).
- **`export_pdf`** — `{"file_path", "output_path", "page_range"?}`. Multi-page PDF. `page_range` accepts `null`/`"all"` (default), `"current"`, `"N"` (single page, 1-based), or `"N-M"` (inclusive range).

### Batch Tools (preferred for >1 shape)

Strongly preferred over looping the single-shape tools when you have
multiple operations to perform. One model round-trip, screen updates
disabled during the batch, single undo scope, no per-op save — typically
5-20x faster end-to-end.

- **`add_shapes`** — `{"file_path", "shapes": [{"shape_type", "x", "y", "width"?, "height"?, "text"?}, ...], "page_name"?}`. Each entry mirrors `add_shape`'s parameters; `text` (optional) sets the shape's text inline so you don't need a separate `add_text` call.
- **`connect_shapes_bulk`** — `{"file_path", "connections": [{"shape1_id", "shape2_id", "connector_type"?}, ...], "page_name"?}`.
- **`style_shapes`** — `{"file_path", "updates": [{"shape_id", "fill"?, "line"?, "text"?, "text_format"?}, ...], "page_name"?}`. `fill` / `line` / `text_format` are sub-objects with the same shape as `set_shape_fill` / `set_shape_line` / `set_shape_text_format`. `text` sets `shape.Text` directly.
- **`delete_shapes`** — `{"file_path", "shape_ids": [int, ...], "page_name"?}`. All-or-nothing: if any ID is missing, the whole batch rolls back.
- **`transform_shapes`** — `{"file_path", "updates": [{"shape_id", "x"?, "y"?, "width"?, "height"?, "angle_degrees"?}, ...], "page_name"?}`. Move, resize, and/or rotate. Partial updates supported.

### Persistence

As of v2.3.0, the mutating single-shape and page tools no longer auto-save
the document. This was the biggest hidden latency cost (~100-300ms per
call for a re-zip + disk write). Call **`save_document(file_path)`**
explicitly when you want the on-disk file in sync, or rely on
`close_document` (saves by default on close).

## Usage Example

Here's a complete workflow example:

1. Create a new Visio file:
```json
{
  "save_path": "C:\\Diagrams\\FlowChart.vsdx"
}
```

2. Add a rectangle shape:
```json
{
  "file_path": "C:\\Diagrams\\FlowChart.vsdx",
  "shape_type": "Rectangle",
  "x": 2.0,
  "y": 2.0,
  "width": 1.5,
  "height": 1.0
}
```

3. Add another shape:
```json
{
  "file_path": "C:\\Diagrams\\FlowChart.vsdx",
  "shape_type": "Circle",
  "x": 5.0,
  "y": 2.0,
  "width": 1.0,
  "height": 1.0
}
```

4. Get shape IDs:
```json
{
  "file_path": "C:\\Diagrams\\FlowChart.vsdx"
}
```

5. Connect the shapes:
```json
{
  "file_path": "C:\\Diagrams\\FlowChart.vsdx",
  "shape1_id": 1,
  "shape2_id": 2,
  "connector_type": "Straight"
}
```

6. Add text to shapes:
```json
{
  "file_path": "C:\\Diagrams\\FlowChart.vsdx",
  "shape_id": 1,
  "text": "Start"
}
```

## Future Features

The following features are planned for future releases:

### Enhanced Shape Styling
- Shadow and 3D effects
- Gradient fills
- True curved connectors (current "curved" routes as Dynamic)

### Advanced Visio Objects
- Layers
- Group creation and manipulation
- Container management
- Text-only shapes and callouts

### Template Management
- Template library access
- Custom template creation
- Favorite templates list

### Batch Operations
- Bulk shape creation
- Mass formatting changes
- Import from CSV/JSON data sources

### Custom Stencil Support
- Loading custom stencils
- Creating and saving custom stencils
- Searching stencil shapes

### Diagram Analysis
- Shape relationship analysis
- Path finding between shapes
- Validation against diagram rules

### Export Options
- PNG/JPG with custom resolution (current exports use the page's natural resolution)
- Per-page PDF metadata (bookmarks, ISO 19005-1 compliance)

### Integration Capabilities
- REST API wrapper
- Webhook support for diagram changes
- Version control integration
- CI/CD pipeline support

### Headless Operation
- Server operation without visible Visio UI
- Background diagram processing
- Scheduled operations

## Troubleshooting

### Common Issues:

1. **Visio Not Launching**:
   - Ensure Visio is correctly installed and can be opened manually
   - Check that you have sufficient permissions to launch COM applications

2. **Template Not Found**:
   - The server will create a blank diagram if templates aren't found
   - Specify an absolute path to a template if needed

3. **Invalid Shape Type**:
   - If a shape type isn't recognized, the server will default to a rectangle
   - Check spelling and case of shape names

4. **COM Errors**:
   - Restarting Visio manually may help resolve COM interface issues
   - Ensure no existing Visio processes are hanging in Task Manager


## License

This project is licensed under the MIT License - see the LICENSE file for details.

