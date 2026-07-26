# OfficeIMO Document Tools Plugin

This repo-local Codex plugin exposes compact OfficeIMO tools for working with local documents and mailboxes, plus contributor skills for conversion, PDF, website/WASM, builds, releases, and PSWritePDF retirement.

The bundled STDIO MCP server runs from the versioned `OfficeIMO.Tool` package:

```powershell
dotnet dnx OfficeIMO.Tool@3.0.2 mcp serve --stdio
```

It exposes five bounded tools:

- `officeimo_inspect`
- `officeimo_search`
- `officeimo_fetch`
- `officeimo_convert`
- `officeimo_capabilities`

The intended agent flow is inspect or search first, then fetch selected content. PST, OST, OLM, EMLX, Mbox, MBX, and message directories are query-first; whole-store conversion is intentionally unavailable.

Set `OFFICEIMO_MCP_ALLOWED_ROOTS` to a platform path-separator-delimited list when the server should be restricted to specific local roots. If it is unset, normal filesystem permissions apply.
