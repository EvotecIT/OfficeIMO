---
title: Browser Converter
description: Run supported OfficeIMO conversions locally through the WebAssembly app on OfficeIMO.com.
order: 90
---

The [browser converter](/playground/) is a static Blazor WebAssembly application. Supported conversions execute inside the current tab; selected files are not uploaded to OfficeIMO.

## Supported browser routes

| Source | Output | Engine |
|---|---|---|
| DOCX | PDF | `OfficeIMO.Word.Pdf` |
| XLSX | PDF | `OfficeIMO.Excel.Pdf` |
| PPTX | PDF | `OfficeIMO.PowerPoint.Pdf` |
| Markdown | HTML preview and download | `OfficeIMO.MarkdownRenderer` |
| HTML | Markdown | `OfficeIMO.Markdown.Html` |
| Markdown | DOCX | `OfficeIMO.Word.Markdown` |

The app includes sample inputs for every route. Uploads are limited to 25 MB so the browser tab remains responsive. Conversion warnings stay visible with the result instead of being hidden behind a successful download.

## Privacy and hosting

Browser-local processing is the strongest privacy default for a public demo because document bytes do not cross a server boundary. It is not the right execution model for every production workload.

Host OfficeIMO in your own service when you need larger inputs, authentication, queues, storage, audit logs, or formats that are not suitable for WebAssembly. In that model, your organization owns the transport, access, logging, and retention policy.

## Publishing contract

The website pipeline builds the converter from its project source and mounts the published `wwwroot` output under `/apps/officeimo-converter/`. This keeps the deployed WebAssembly assets and integrity metadata aligned with the source in the same build.

For a local production-shaped publish:

```powershell
pwsh -NoProfile -File .\Website\build.ps1 -PowerForgeRoot C:\path\to\PSPublishModule
```

Use `-Dev` for a faster content/API build. The converter project itself can be checked directly with:

```powershell
dotnet build .\Website\Apps\OfficeIMO.Web.Converter\OfficeIMO.Web.Converter.csproj -c Release
```
