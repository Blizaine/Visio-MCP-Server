# Visio MCP Server — Colleague Bootstrap Guide

Get this server installed on a new Windows machine and wired into your MCP
client in about 5 minutes. **No admin rights needed.**

> Before you start: confirm with your team lead which install source to use.
> Throughout this guide we'll write `<INSTALL_SOURCE>` for the URL or path
> you've been given (e.g. a private git repo, a wheel file, or a local
> clone). See [Install sources](#install-sources) below.

## Prerequisites

You need both of these already working on your machine:

- **Windows 10 or 11** (this server uses COM and is Windows-only).
- **Microsoft Visio** installed and openable from the Start menu.
  Confirm at a PowerShell prompt:
  ```powershell
  Get-Item "Registry::HKEY_CLASSES_ROOT\Visio.Application"
  ```
  If that returns a key, Visio's COM bridge is registered and you're good.

You do **not** need Python pre-installed. `uv` (next step) handles that.

## Quick start (automated)

From this repo's root in PowerShell:

```powershell
scripts\bootstrap.ps1 -Source "<INSTALL_SOURCE>"
```

The script will:

1. Install `uv` to `%USERPROFILE%\.local\bin` (user-local, no admin).
2. Use `uv` to install a managed CPython 3.12 (also user-local).
3. `uv tool install` the Visio MCP Server from `<INSTALL_SOURCE>`.
4. Print the JSON snippet you need to paste into your MCP client's
   config (it will not auto-edit the file — see [MCP client setup](#mcp-client-setup)).
5. Print the path to the smoke test so you can verify the install.

If anything goes wrong, the script prints which step failed; the
[Manual install](#manual-install-step-by-step) section below walks the same
steps individually.

## Manual install (step by step)

Pick this path if your environment blocks scripts, your IT approval process
needs you to install one tool at a time, or the automated script failed and
you want to see exactly what's happening.

### 1. Install uv

In PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

This downloads `uv.exe`, `uvx.exe`, and `uvw.exe` into
`%USERPROFILE%\.local\bin` and adds that directory to your **user** PATH.
You may need to close and reopen PowerShell for the PATH change to take
effect. Confirm:

```powershell
uv --version
```

### 2. Have uv install Python

```powershell
uv python install 3.12
```

This downloads a portable CPython 3.12 build to `%APPDATA%\uv\python`. It
is *not* registered as a system Python and won't conflict with any other
Python you may install later.

### 3. Install the Visio MCP Server

```powershell
uv tool install "<INSTALL_SOURCE>"
```

This isolates the server in its own venv under
`%USERPROFILE%\.local\share\uv\tools\office-visio-mcp-server\` and exposes
the launcher at `%USERPROFILE%\.local\bin\visio_mcp_server.exe`. That
directory is already on your PATH from step 1.

Confirm:

```powershell
Get-Command visio_mcp_server
```

### 4. Configure your MCP client

See [MCP client setup](#mcp-client-setup) below.

### 5. Verify

Run the smoke test from a clone of the repo (or download `scripts/smoke_test.py`):

```powershell
.\.venv\Scripts\python.exe scripts\smoke_test.py
```

> The smoke test needs the repo's `.venv` to exist. If you only installed
> via `uv tool install` and don't have the source, ask your team lead for a
> standalone copy of `smoke_test.py` and point it at `visio_mcp_server.exe`
> as the server command.

The test will pop Visio up, draw two shapes, connect them, and save a
`.vsdx` to your TEMP folder. Open that file and confirm the connector line
is visible.

## MCP client setup

### Claude Desktop

Add the server to `%APPDATA%\Claude\claude_desktop_config.json`. If the
file doesn't exist yet, create it with this content. If it does, merge
just the `visio-server` entry into the `mcpServers` object:

```json
{
  "mcpServers": {
    "visio-server": {
      "command": "C:\\Users\\<YOUR_USERNAME>\\.local\\bin\\visio_mcp_server.exe",
      "args": [],
      "env": {}
    }
  }
}
```

Replace `<YOUR_USERNAME>` with your actual Windows username (PowerShell:
`echo $env:USERNAME`). Restart Claude Desktop after editing.

### Claude Code

Add to `%USERPROFILE%\.claude\settings.json` (create if missing) under the
`mcpServers` key with the same shape as above. Restart Claude Code.

### Other MCP clients

The server speaks the standard MCP stdio transport. Any client that lets
you point at a command + args (Cline, Continue, etc.) will work — just
point it at the `visio_mcp_server.exe` path.

## Install sources

How you get the package depends on how your team is distributing it:

| Source style | What `<INSTALL_SOURCE>` looks like |
|---|---|
| Private git repo (recommended for early sharing) | `git+https://github.com/<org>/<repo>.git` |
| Public PyPI (when published) | `office-visio-mcp-server` |
| Internal PyPI mirror | `office-visio-mcp-server` (with `UV_INDEX_URL` set to the mirror) |
| Local clone, editable | `.` (run inside the repo, after `uv venv`) |
| Wheel file you've been emailed | `path\to\office_visio_mcp_server-2.0.0-py3-none-any.whl` |

If you have the source cloned and just want a dev install, the simpler
flow is:

```powershell
uv venv
uv pip install -e .
```

Then point your MCP client at `<repo>\.venv\Scripts\python.exe -m visio_mcp_server.visio_server`.

## Troubleshooting

**`uv` command not found after install.**
Open a new PowerShell window. The installer modifies user PATH; existing
shells don't see the change until they're restarted. As a one-session
workaround:
```powershell
$env:Path = "$env:USERPROFILE\.local\bin;$env:Path"
```

**`irm | iex` blocked by AV or by your execution policy.**
Download the installer manually from
<https://github.com/astral-sh/uv/releases> (look for
`uv-x86_64-pc-windows-msvc.zip`), unzip, copy `uv.exe` into
`%USERPROFILE%\.local\bin\`, then add that directory to your user PATH.

**`pip install` fails with TLS / cert errors.**
Your corporate network is proxying or rewriting PyPI requests. Point uv at
your internal mirror:
```powershell
$env:UV_INDEX_URL = "https://<internal-pypi-host>/simple/"
```

**The server starts but Visio never appears.**
Check `Get-Item "Registry::HKEY_CLASSES_ROOT\Visio.Application"` returns a
key. If not, Visio's COM bridge isn't registered — usually fixed by
running `VISIO.EXE /regserver` from an elevated prompt (requires admin
once).

**Connector lines are invisible in the saved diagram.**
You're running v1.x; upgrade to >= 2.0.0. The invisible-connector bug was
fixed in Phase 1 (`LinePattern="0"` removed from `connect_shapes`).

## Uninstall

```powershell
uv tool uninstall office-visio-mcp-server
```

To remove uv and the managed Python:

```powershell
Remove-Item -Recurse -Force "$env:USERPROFILE\.local\bin"
Remove-Item -Recurse -Force "$env:APPDATA\uv"
Remove-Item -Recurse -Force "$env:USERPROFILE\.local\share\uv"
```
