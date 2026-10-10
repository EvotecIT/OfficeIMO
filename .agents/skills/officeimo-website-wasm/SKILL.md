---
name: officeimo-website-wasm
description: Use when adding or changing OfficeIMO.com browser tools (/browser/<tool>/ pages and the /convert/ directory), the browser-local OfficeIMO engine (Website/Apps/OfficeIMO.Web.Converter), its WebAssembly publishing, or static-site integration for drag/drop document tools.
---

# OfficeIMO Website browser tools

Use this skill for OfficeIMO.com browser tool work.

## Architecture

- **Tool pages are static HTML.** `Website/themes/officeimo/layouts/browser-tool.html` renders `/browser/<tool>/` from
  `Website/data/browser_tools.json`. The page is usable before any .NET code loads.
- **One script drives every tool.** `Website/static/js/browser-tool.js` (ES5 style for the site minifier) reads the page's
  `data-*` attributes, stages files, runs the engine, and renders the uniform result (verdict, facts, items, preview,
  downloads, "continue with" hand-off through a randomized, short-lived BroadcastChannel). Files remain in memory;
  the sender retains them until the receiving page accepts them. The script deletes the retired IndexedDB store.
- **The engine runs in a Web Worker.** `Website/Apps/OfficeIMO.Web.Converter` is a `Microsoft.NET.Sdk.WebAssembly` app with
  no UI. `wwwroot/engine-worker.js` boots it and calls `Engine/EngineExports.cs` through `[JSExport]`. Conversions never
  block the page. Each tool lists the lazily loaded engine assemblies it needs; the worker prefetches them while the
  runtime starts.
- **Every tool returns the same result shape** (`Engine/ToolResults.cs`): a one-sentence verdict, facts, named items with
  states (found, removed, kept, warning…), artifacts by role (primary, report, overlay, support), and a preview kind.
- The `/convert/` directory (`partials/shortcodes/browser-tools.html`) lists tools from the same data file and forwards
  old `/convert/?route=…` and `?workspace=…` links using each tool's `legacy` keys.

## Decision rules

- Do not require server processes, native binaries, Office, LibreOffice, queues, or a database for the GitHub Pages path.
- Keep local automation, batch conversion, and agent workflows in the CLI or an MCP server, not in the public browser app.
- Values shown on pages come from catalogs (`office_conversion_routes.json`, the engine catalogs); write plain-language
  copy for visitors, not engine terms. Name findings, don't just count them.
- One primary action at a time. Read-only checks (inspect, origin, compare) run as soon as files are added.
- Keep reflection out of the engine: it is fully trimmed with trim analysis on. Use source-generated JSON contexts.

## Add a browser tool

1. Add an entry to `Website/data/browser_tools.json`: `id` (URL slug), `group`, `title`, `summary`, `seoTitle`, `input`
   (`kind`: file, files, pair, or text; `accept`; optional `sample`, `sampleSecond`, `sampleOptions`, `auto`, `live`),
   `options` (choice, flag, number, text, password, pages), `engine` (`kind`, `target`, `action`, optional `commit`,
   `find`, `confirm`, and the lazy `assemblies`), `run` labels, `expect`, `next`, `legacy`, `guide`, `package`.
2. If the engine can already run it (a new conversion route or PDF operation), no C# is needed. Otherwise add a handler
   beside `Engine/ConvertTool.cs` / `PdfTool.cs` / `OriginTool.cs` / `TextTool.cs` and dispatch it in `EngineExports.Run`.
   Return a `ToolResultDocument` with a plain verdict.
3. Mark any new engine assembly with `BlazorWebAssemblyLazyLoad` in the project file.
4. Run `pwsh Website/scripts/Sync-BrowserToolPages.ps1` to create the page front matter.
5. Run `dotnet test Website/Apps/OfficeIMO.Web.Converter.Tests`. `BrowserToolCatalogTests` checks the data against the
   engine catalogs, lazy-load list, generated pages, samples, and legacy links.

## Validation

```powershell
# Native relink needs the wasm-tools workload. On machines with a very long PATH, run publish with the machine PATH only.
dotnet publish Website/Apps/OfficeIMO.Web.Converter/OfficeIMO.Web.Converter.csproj -c Release
dotnet test Website/Apps/OfficeIMO.Web.Converter.Tests/OfficeIMO.Web.Converter.Tests.csproj -c Release
pwsh Website/build.ps1 -Dev -Only 'build-site,deploy-theme-css'   # then overlay the engine wwwroot into _site/apps/officeimo-converter
pwsh Website/scripts/Test-ConverterPublish.ps1 -SiteRoot <absolute path to Website/_site>
pwsh Website/scripts/Test-ConverterPerformance.ps1 -SiteRoot <absolute path to Website/_site>
```

Then check in a real browser that each changed tool runs end to end from its sample: verdict, preview, primary download,
and that the page shows nothing until asked except the tool itself (no runtime download on `/convert/`).
