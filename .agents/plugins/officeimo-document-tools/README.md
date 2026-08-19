# OfficeIMO Document Tools Plugin

This repo-local Codex plugin exposes compact OfficeIMO tools for working with local documents and mailboxes, plus contributor skills for conversion, PDF, website/WASM, builds, releases, and PSWritePDF retirement.

The bundled STDIO MCP server runs from the versioned `OfficeIMO.Tool` package:

```powershell
dotnet dnx OfficeIMO.Tool@3.2.4 mcp serve --stdio
```

It exposes five bounded tools:

- `officeimo_inspect`
- `officeimo_search`
- `officeimo_fetch`
- `officeimo_convert`
- `officeimo_capabilities`

The intended agent flow is inspect or search first, then fetch selected content. PST, OST, OLM, EMLX, Mbox, MBX, and message directories are query-first; whole-store conversion is intentionally unavailable.

The server defaults filesystem access to the working directory inherited from Codex, which normally scopes it to the current workspace. Paths outside that directory are rejected. Set `OFFICEIMO_MCP_ALLOWED_ROOTS` to an exact-cased, platform path-separator-delimited replacement list only when the server should use different local roots; include the working directory explicitly when it should remain available.
