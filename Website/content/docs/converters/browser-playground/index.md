---
title: Browser Conversion Playground
description: Static Blazor WebAssembly conversion path for OfficeIMO.com and GitHub Pages-style hosting.
order: 90
---

The OfficeIMO browser conversion playground is planned as a static Blazor WebAssembly app hosted by OfficeIMO.com. It should convert user-selected files locally in the browser and return downloadable output without sending document contents to a server.

## Boundary

The browser playground can use:

- OfficeIMO byte and stream APIs;
- Blazor WebAssembly;
- static files published with the website;
- local browser file upload and download APIs.

It should not require:

- Office automation;
- LibreOffice or a server-side conversion process;
- native graphics dependencies;
- Redis, queues, databases, or background jobs;
- private server storage.

## Initial Scope

The first useful version should support:

- DOCX to PDF;
- XLSX to PDF;
- PPTX to PDF;
- structured diagnostics for unsupported features, font gaps, and large files.

The current proof shows basic DOCX, XLSX, and PPTX conversion can produce `%PDF` bytes inside the browser runtime. A richer Word fixture still exposes a Unicode/font embedding gap, so diagnostics need to be part of the first public surface.

## Static Mount

Published app output should land under:

```text
Website/static/apps/officeimo-converter/
```

The public URL should be:

```text
/apps/officeimo-converter/
```

The website page at `/playground/` can introduce the feature and link to the mounted app when the Blazor build is ready.

## Validation

Before publishing a playground build:

```powershell
dotnet publish <BlazorWasmProject>.csproj -c Release
```

Then verify in a real browser:

- DOCX to PDF returns bytes beginning with `%PDF`.
- XLSX to PDF returns bytes beginning with `%PDF`.
- PPTX to PDF returns bytes beginning with `%PDF`.
- Known unsupported Word font cases produce actionable diagnostics.
