---
title: Browser Conversion Playground
description: Static Blazor WebAssembly conversion path for OfficeIMO.com and GitHub Pages-style hosting.
order: 90
---

The OfficeIMO browser conversion playground is a static Blazor WebAssembly app hosted by OfficeIMO.com. It converts user-selected files locally in the browser and returns downloadable PDF output without sending document contents to a server.

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

## Current Scope

The current version supports:

- DOCX to PDF;
- XLSX to PDF;
- PPTX to PDF;
- structured diagnostics for unsupported features, font gaps, and large files.

The playground includes small built-in DOCX, XLSX, and PPTX samples so browser smoke tests can prove conversion without uploading local files. User-selected files still use the same byte and stream conversion path.

Basic DOCX, XLSX, and PPTX conversion can produce `%PDF` bytes inside the browser runtime. A richer Word fixture still exposes a Unicode/font embedding gap, so diagnostics remain part of the public surface.

## Static Mount

Published app output should land under:

```text
Website/static/apps/officeimo-converter/
```

The public URL should be:

```text
/apps/officeimo-converter/
```

The website page at `/playground/` embeds the mounted app under the OfficeIMO site chrome.

## Validation

Before publishing a playground build:

```powershell
dotnet build Website\Apps\OfficeIMO.Web.Converter\OfficeIMO.Web.Converter.csproj -c Release
dotnet publish Website\Apps\OfficeIMO.Web.Converter\OfficeIMO.Web.Converter.csproj -c Release -o Website\_temp\officeimo-converter-publish
```

Then verify in a real browser:

- DOCX to PDF returns bytes beginning with `%PDF`.
- XLSX to PDF returns bytes beginning with `%PDF`.
- PPTX to PDF returns bytes beginning with `%PDF`.
- Known unsupported Word font cases produce actionable diagnostics.
