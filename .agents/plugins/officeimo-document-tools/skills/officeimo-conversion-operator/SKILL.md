---
name: officeimo-conversion-operator
description: Use when changing or validating OfficeIMO document conversion behavior, especially DOCX/XLSX/PPTX/HTML/Markdown/PDF byte and stream conversions, browser-safe conversion paths, or fidelity diagnostics. Prefer OfficeIMO core APIs first and keep hosts, websites, cmdlets, MCP tools, and demos thin.
---

# OfficeIMO Conversion Operator

Use this skill for OfficeIMO conversion work across Word, Excel, PowerPoint, Markdown, HTML, and PDF.

## Golden Path

1. Start by locating the owning reusable API.
   - Word behavior belongs under `OfficeIMO.Word*`.
   - Excel behavior belongs under `OfficeIMO.Excel*`.
   - PowerPoint behavior belongs under `OfficeIMO.PowerPoint*`.
   - Shared PDF behavior belongs in the PDF engine projects, not in a host wrapper.
2. Treat file paths as adapter concerns.
   - Prefer byte array and stream entrypoints for conversion APIs.
   - Keep browser, CLI, PowerShell, MCP, and website shells as orchestration only.
3. Validate real artifacts.
   - For PDF output, at minimum verify the returned bytes start with `%PDF`.
   - For browser work, validate inside a real Blazor WebAssembly runtime before calling it GitHub Pages-safe.
   - For fidelity work, keep a named fixture and record the expected known gap.
4. Separate platform limits from conversion defects.
   - Browser-safe means no Office automation, LibreOffice process, native graphics dependency, server filesystem dependency, or environment-only state.
   - Font coverage, Unicode embedding, large memory use, and startup size are conversion quality gaps even when the browser runtime works.
5. Keep public surfaces friendly.
   - Add simple options and presets at host layers only after the reusable OfficeIMO API already supports the behavior.
   - Avoid hidden environment-variable switches for behavior that should be an explicit option.

## Useful Checks

```powershell
dotnet build OfficeIMO.Word.Pdf\OfficeIMO.Word.Pdf.csproj -c Release -f net10.0
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "*SaveAsPdf*"
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "*Html*|*Markdown*"
```

When the work targets GitHub Pages or OfficeIMO.com, pair these with the browser conversion proof described in `Docs/officeimo.blazor-wasm-conversion-proof.md`.
