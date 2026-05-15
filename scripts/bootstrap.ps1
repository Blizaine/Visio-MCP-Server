<#
.SYNOPSIS
    Bootstrap the Visio MCP Server on a fresh Windows machine — no admin required.

.DESCRIPTION
    Installs uv (a single-binary Python package manager), uses it to install a
    managed CPython 3.12, then installs the Visio MCP Server from the source you
    specify. Prints the MCP client config snippet at the end so you can paste it
    into Claude Desktop, Claude Code, or another MCP client.

    See BOOTSTRAP.md in this repo for the full guide and troubleshooting.

.PARAMETER Source
    Where to install the Visio MCP Server from. Examples:
      - git+https://github.com/<org>/<repo>.git    (private/public git repo)
      - cti-visio-mcp-server                     (PyPI, if/when published)
      - C:\path\to\cti_visio_mcp_server-3.0.0-py3-none-any.whl   (wheel)
      - .                                           (local clone — run from repo root)

.PARAMETER SkipVisioCheck
    Don't verify that Visio's COM bridge is registered. Useful only on machines
    where Visio will be installed later.

.EXAMPLE
    .\scripts\bootstrap.ps1 -Source "git+https://github.com/blaine-cti/visio-mcp-server.git"

.EXAMPLE
    .\scripts\bootstrap.ps1 -Source "."

.NOTES
    Does not require admin. Does not modify system PATH or system Python.
    Everything lands under $env:USERPROFILE\.local\ or $env:APPDATA\uv\.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Install source: git URL, PyPI name, wheel path, or '.'")]
    [string]$Source,

    [switch]$SkipVisioCheck
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([int]$Number, [string]$Message)
    Write-Host ""
    Write-Host "[$Number] $Message" -ForegroundColor Cyan
}

function Write-Ok    { param([string]$Message) Write-Host "    OK: $Message" -ForegroundColor Green }
function Write-Warn  { param([string]$Message) Write-Host "    WARN: $Message" -ForegroundColor Yellow }
function Write-Fail  { param([string]$Message) Write-Host "    FAIL: $Message" -ForegroundColor Red }


# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------

Write-Step 1 "Pre-flight checks"

if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
    Write-Fail "This bootstrap is Windows-only (Visio is Windows-only)."
    exit 1
}
Write-Ok  "Windows host detected"

# Claude must be fully closed. Its Microsoft Store package keeps the
# launcher .exe locked while the app is running, and `uv tool install`
# (even with --force) produces a half-broken install in that state.
# See BOOTSTRAP.md > "Updating the server".
$claudeProcs = Get-Process claude -ErrorAction SilentlyContinue
if ($claudeProcs) {
    Write-Fail "Claude is currently running ($($claudeProcs.Count) process(es)). Close it fully before installing."
    Write-Host  "       Quit from File > Exit (not just close the window). Then check Task Manager"
    Write-Host  "       for any lingering claude.exe before retrying."
    exit 1
}
Write-Ok  "No Claude processes detected"

if (-not $SkipVisioCheck) {
    $visio = Get-Item "Registry::HKEY_CLASSES_ROOT\Visio.Application" -ErrorAction SilentlyContinue
    if (-not $visio) {
        Write-Fail "Visio.Application COM class not found in the registry."
        Write-Host  "       Install Microsoft Visio first, or rerun with -SkipVisioCheck"
        Write-Host  "       if you'll install it later."
        exit 1
    }
    Write-Ok "Visio.Application COM class registered"
} else {
    Write-Warn "Visio check skipped per -SkipVisioCheck"
}


# ---------------------------------------------------------------------------
# 2. Install uv (user-local, no admin)
# ---------------------------------------------------------------------------

Write-Step 2 "uv (Astral)"

$uvBin = Join-Path $env:USERPROFILE ".local\bin"
$uvExe = Join-Path $uvBin "uv.exe"
$env:Path = "$uvBin;$env:Path"  # make uv visible in this session even before install

$existingUv = Get-Command uv -ErrorAction SilentlyContinue
if ($existingUv) {
    Write-Ok "uv already installed at $($existingUv.Source)"
} else {
    Write-Host "    Installing uv from https://astral.sh/uv/install.ps1 ..."
    try {
        powershell -ExecutionPolicy Bypass -NoProfile -Command "irm https://astral.sh/uv/install.ps1 | iex"
    } catch {
        Write-Fail "uv install failed: $($_.Exception.Message)"
        Write-Host  "       See BOOTSTRAP.md > Troubleshooting > 'irm | iex blocked' for a manual alternative."
        exit 1
    }
    if (-not (Test-Path $uvExe)) {
        Write-Fail "uv installer ran but $uvExe is missing."
        exit 1
    }
    Write-Ok "uv installed at $uvExe"
}


# ---------------------------------------------------------------------------
# 3. Managed Python 3.12
# ---------------------------------------------------------------------------

Write-Step 3 "Managed Python 3.12"

# `uv python list` shows installed + downloadable; `uv python install` is
# idempotent (no-ops if the version is already present).
& uv python install 3.12
if ($LASTEXITCODE -ne 0) {
    Write-Fail "uv python install 3.12 failed"
    exit 1
}
Write-Ok "Python 3.12 available to uv"


# ---------------------------------------------------------------------------
# 4. Install the Visio MCP Server
# ---------------------------------------------------------------------------

Write-Step 4 "Visio MCP Server"

Write-Host "    Source: $Source"
# If a previous install is present, uninstall first. We deliberately avoid
# `--force` because, when the launcher .exe is held open by a running
# Claude Store app, `--force` produces a partial install with files
# hardlinked into the app's package cache. See BOOTSTRAP.md > "Updating
# the server" for the gory details.
& uv tool list 2>$null | Select-String -Pattern "^cti-visio-mcp-server " | Out-Null
if ($LASTEXITCODE -eq 0) {
    Write-Host "    Removing existing install first..."
    & uv tool uninstall cti-visio-mcp-server 2>$null | Out-Null
}
& uv tool install --python 3.12 $Source
if ($LASTEXITCODE -ne 0) {
    Write-Fail "uv tool install failed."
    Write-Host  "       Double-check the -Source value. See BOOTSTRAP.md > 'Install sources'."
    Write-Host  "       If you saw a 'reparse point' or 'access denied' error, fully exit"
    Write-Host  "       Claude (verify in Task Manager) and try again. See BOOTSTRAP.md >"
    Write-Host  "       'Recovering from a broken install'."
    exit 1
}

$serverExe = Join-Path $uvBin "visio_mcp_server.exe"
if (-not (Test-Path $serverExe)) {
    Write-Fail "Server binary not found at $serverExe after install."
    Write-Host  "       Run ``uv tool list`` to see what was installed."
    exit 1
}
Write-Ok "Server installed; launcher at $serverExe"


# ---------------------------------------------------------------------------
# 5. Print the MCP client config snippet
# ---------------------------------------------------------------------------

Write-Step 5 "MCP client config"

# JSON requires backslashes to be escaped. PowerShell's literal $serverExe
# already uses single backslashes, so we double them for the snippet.
$escapedExe = $serverExe -replace '\\', '\\'

$snippet = @"
{
  "mcpServers": {
    "visio-server": {
      "command": "$escapedExe",
      "args": [],
      "env": {}
    }
  }
}
"@

Write-Host ""
Write-Host "Paste this into your MCP client's config file:" -ForegroundColor Cyan
Write-Host "  Claude Desktop:  $env:APPDATA\Claude\claude_desktop_config.json"
Write-Host "  Claude Code:     $env:USERPROFILE\.claude\settings.json"
Write-Host ""
Write-Host "If the file already has an ``mcpServers`` object, just merge the"
Write-Host "``visio-server`` entry into it (don't replace the whole file)."
Write-Host ""
Write-Host $snippet
Write-Host ""
Write-Host "Restart your MCP client after saving the config."


# ---------------------------------------------------------------------------
# 6. Verify
# ---------------------------------------------------------------------------

Write-Step 6 "Verify"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$smoke = Join-Path $scriptDir "smoke_test.py"
if (Test-Path $smoke) {
    Write-Host "    To exercise every tool against a real Visio, run:"
    Write-Host ""
    Write-Host "      uv run --with mcp python $smoke" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    This will pop Visio up, draw two shapes, connect them, and save"
    Write-Host "    a .vsdx in your TEMP folder. Open it and confirm the connector"
    Write-Host "    line is visible."
} else {
    Write-Warn "smoke_test.py not found alongside this script; can't suggest a verify command."
}

Write-Host ""
Write-Host "Bootstrap complete." -ForegroundColor Green
