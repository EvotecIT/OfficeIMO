---
title: Browser Conversion Playground
description: Static Blazor WebAssembly conversion path for OfficeIMO.com and GitHub Pages-style hosting.
order: 90
---

The OfficeIMO browser conversion playground is a static Blazor WebAssembly app hosted by OfficeIMO.com. It converts selected files and text locally in the browser and returns downloadable output without sending document contents to a server.

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

The current live browser version supports:

- DOCX to PDF;
- XLSX to PDF;
- PPTX to PDF;
- Markdown to HTML;
- HTML to Markdown;
- Markdown to DOCX;
- structured diagnostics for unsupported features, font gaps, and large files.

The playground includes small built-in DOCX, XLSX, and PPTX samples plus Markdown and HTML samples so browser smoke tests can prove conversion without uploading local files. User-selected files still use the same byte and stream conversion path.

Basic DOCX, XLSX, and PPTX conversion can produce `%PDF` bytes inside the browser runtime. A richer Word fixture still exposes a Unicode/font embedding gap, so diagnostics remain part of the public surface.

The text workspace calls `OfficeIMO.MarkdownRenderer`, `OfficeIMO.Markdown.Html`, and `OfficeIMO.Word.Markdown` directly from WebAssembly. It can render Markdown HTML previews, download Markdown converted from HTML, and generate DOCX bytes from Markdown.

## Engine Map

The playground also shows OfficeIMO conversion families that should become richer playground, CLI, PowerShell, MCP, plugin, skill, or server routes:

- DOCX to Markdown;
- DOCX to HTML through the Markdown bridge;
- HTML to DOCX;
- Markdown to PDF through Markdown to Word and the PDF engine;
- Excel to CSV, JSON-style records, HTML tables, and sheet previews;
- CSV or JSON to Excel workbook generation;
- Reader extraction to Markdown, JSON, chunks, tables, and assets;
- PDF split, merge, stamp, inspect, extract, fill, and metadata workflows;
- OfficeIMO.Markup exporters for Word, Excel, PowerPoint, C#, and PowerShell;
- repeatable agent tools exposed through MCP servers, Codex skills, and release/build plugins.

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

Copy the published static app from:

```text
Website/_temp/officeimo-converter-publish/wwwroot/
```

into:

```text
Website/static/apps/officeimo-converter/
```

Do not publish directly into `Website/static/apps/officeimo-converter/`, because the Blazor publish layout places the GitHub Pages-ready app under `wwwroot/`.

Then verify in a real browser:

- DOCX to PDF returns bytes beginning with `%PDF`.
- XLSX to PDF returns bytes beginning with `%PDF`.
- PPTX to PDF returns bytes beginning with `%PDF`.
- Markdown to HTML renders an HTML preview and download.
- HTML to Markdown returns Markdown text and download.
- Markdown to DOCX returns downloadable DOCX bytes.
- Known unsupported Word font cases produce actionable diagnostics.
