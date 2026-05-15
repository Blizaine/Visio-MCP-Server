---
name: Bug report
about: Something doesn't work the way you expect
title: ''
labels: bug
---

## What happened

<!-- A short description. -->

## What you expected

<!-- One or two sentences. -->

## Reproduction

<!--
The fewest tool calls needed to trigger the bug. JSON arguments + the
envelope you got back (especially the `error` block if ok=false) is the
fastest way to a fix.
-->

## Environment

- Visio version: <!-- Help > About in Visio, e.g., 2406 Build 17726.20126 -->
- OS: <!-- Windows 10 / 11 / Enterprise / Pro -->
- Package version: <!-- `uv tool list` or pyproject.toml -->
- Python version: <!-- python --version -->
- MCP client: <!-- Claude Desktop, Claude Code, Cline, ... -->

## Logs

<!--
If the failure includes a Python traceback or a Visio COM error, paste
the relevant lines here. The server logs to stderr (look in the MCP
client's log directory for visio-server logs).
-->
