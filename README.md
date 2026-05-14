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

### File Operations
- Save documents to specified locations
- Close documents safely

> Image/PDF export is on the roadmap (see "Future Features" below) but not yet implemented.

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
| `SHAPE_NOT_FOUND` | A requested shape ID isn't on the active page. |
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
Lists all shapes in a Visio diagram.

```json
{
  "file_path": "Path to the Visio file"
}
```

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
- Color and fill pattern customization
- Line weight, style, and color options
- Text formatting (font, size, alignment)
- Shadow and 3D effects

### Advanced Visio Objects
- Support for layers and pages
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
- PDF export with options
- SVG export for web use
- PNG/JPG with custom resolution
- Export specific pages or sections

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

