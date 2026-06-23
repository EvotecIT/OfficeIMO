---
name: officeimo-website-wasm
description: Use when adding or changing OfficeIMO.com browser conversion features, Blazor WebAssembly publishing, GitHub Pages-safe conversion demos, or static-site integration for drag/drop document conversion.
---

# OfficeIMO Website WASM

Use this skill for OfficeIMO.com browser conversion work.

## Decision Rules

- Public browser conversion belongs in a static Blazor WebAssembly app published under `Website/static/apps/officeimo-converter/`.
- The app should call OfficeIMO byte and stream APIs directly.
- The website should provide routing, content, discoverability, and static hosting only.
- Do not require server processes, native binaries, Office, LibreOffice, Redis, queues, or a database for the GitHub Pages path.
- Keep local automation, batch conversion, repository inspection, and agent workflows in an MCP server or CLI, not in the public browser app.

## Browser-Safe Feature Order

1. Upload or drag/drop a local DOCX, XLSX, or PPTX file.
2. Convert with OfficeIMO core APIs in the browser runtime.
3. Return a downloadable PDF or converted document.
4. Show structured diagnostics for unsupported features, missing fonts, large files, or conversion exceptions.
5. Add optional sample fixtures only after user-supplied files work.

## Validation

For every meaningful WASM change:

```powershell
dotnet publish <BlazorWasmProject>.csproj -c Release
```

Then verify with a real browser runtime:

- DOCX to PDF basic fixture returns bytes starting with `%PDF`.
- XLSX to PDF returns bytes starting with `%PDF`.
- PPTX to PDF returns bytes starting with `%PDF`.
- A known rich Word fixture reports the current Unicode/font gap instead of silently failing.

Use `Docs/officeimo.blazor-wasm-conversion-proof.md` as the baseline evidence and update it when the browser matrix changes.
