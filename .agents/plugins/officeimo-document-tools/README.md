# OfficeIMO Document Tools Plugin

This repo-local Codex plugin exposes OfficeIMO operating skills for conversion, PDF, website/WASM, and PSWriteOffice retirement work.

The plugin intentionally does not declare an MCP server yet. The OfficeIMO MCP server should be added only when the CLI/server entrypoint exists, so plugin installation does not advertise a dead command.

Planned MCP entrypoint:

```powershell
officeimo mcp serve
```

The MCP server should wrap reusable OfficeIMO APIs and expose tools such as `convert_document`, `inspect_document`, `compare_pdf_outputs`, `list_supported_formats`, and `run_conversion_fixture`.
