---
title: Browser Document Workspace
description: Run supported OfficeIMO conversions and focused PDF workflows locally in your browser through the WebAssembly app on OfficeIMO.com.
order: 90
---

The [browser document workspace](/convert/) is a static Blazor WebAssembly application. Supported conversions and PDF operations execute inside the current tab; selected file bytes are not uploaded to OfficeIMO.

## ChatGPT Website Tool

Open the [full-screen browser workspace](/apps/officeimo-converter/) in a browser that supports Website Tools to make `convert_selected_document` available on that page. Choose or drop a document and select the route and settings in the visible workspace first. ChatGPT can then invoke the same **Convert** action, and the generated download, warnings, and diagnostics remain visible on the page.

The tool has no file, path, URL, or format arguments. It can act only on the document and settings already selected by the user, and processing remains browser-local. It returns bounded metadata rather than document contents. Closing or navigating away from the workspace unregisters the tool. Cancellation is honored before conversion starts; once the synchronous browser-local conversion has started, it completes atomically.

## Supported browser routes

| Source | Output | Engine |
|---|---|---|
| DOCX | PDF | `OfficeIMO.Word.Pdf` |
| XLSX | PDF | `OfficeIMO.Excel.Pdf` |
| PPTX | PDF | `OfficeIMO.PowerPoint.Pdf` |
| HTML | PDF | `OfficeIMO.Html.Pdf` |
| Markdown | HTML preview and download | `OfficeIMO.MarkdownRenderer` |
| HTML | Markdown | `OfficeIMO.Markdown.Html` |
| Markdown | DOCX | `OfficeIMO.Word.Markdown` |
| PDF | DOCX | `OfficeIMO.Word.Pdf` |
| PDF | XLSX | `OfficeIMO.Excel.Pdf` |
| PDF | PPTX | `OfficeIMO.PowerPoint.Pdf` |
| PDF | HTML | `OfficeIMO.Html.Pdf` |

PDF-to-Office routes reconstruct logical content and return conversion diagnostics. They do not claim that an arbitrary fixed-layout PDF can reproduce its original editable source.

## PDF tools

Switch to **PDF tools** for twelve task-oriented workflows backed by `OfficeIMO.Pdf`:

| Group | Tools |
|---|---|
| Understand | Inspect; visual compare with a self-contained difference gallery |
| Organize | Merge, split, extract, delete, reorder, and rotate pages |
| Publish | Deterministic lossless optimization, including Fast Web View |
| Secure | AES-256 password protection, unlock, and verified literal-text redaction |

Successful PDF operations return the result and a JSON report containing browser-local execution evidence, input and output SHA-256 fingerprints, and operation details. The built-in sample is the same [showcase PDF](/downloads/showcase/pdf/showcase-dashboard.pdf) published with the product examples.

PDF results include a page-image preview rendered by OfficeIMO, with previous/next controls for multipage output. **Protect PDF** returns a download without a preview because its output requires the password you supplied. Other previews report rendering limits and failures while keeping the generated file available to download. Download the PDF for selectable text, interactive forms, and review in another viewer.

## Limits and input policy

The app includes sample inputs for every route. Files are limited to 25 MiB. Multi-file PDF tools accept up to ten PDFs and 75 MiB combined. PDF parsing is capped at 500 pages, split at 100 outputs and 64 MiB of serialized PDFs, visual comparison at 25 pages, and any generated artifact at 96 MiB. Before a DOCX, XLSX, or PPTX file is parsed, the app also rejects packages with more than 5,000 parts, an individual expanded part over 32 MiB, more than 128 MiB expanded in total, or a part compression ratio over 200:1.

Excel workbooks that pass those package checks are converted in full while every sheet's used range stays within 50,000 cells. If a sheet exceeds that budget, the app automatically generates a preview of up to 250 rows per sheet. Conversion warnings stay visible with the result instead of being hidden behind a successful download.

## Privacy and hosting

Browser-local processing is the strongest privacy default for a public demo because document bytes do not cross a server boundary. It is not the right execution model for every production workload.

Passwords are used for the selected operation and cleared from component state when it finishes. Browser-local execution still means the user controls the device, browser extensions, downloads, and local storage policy.

The workspace does not expose OCR, searchable-PDF generation, lossy scan compression, or cryptographic signing. Those capabilities need provider, quality, identity, or trust decisions that do not belong behind a generic one-click browser action.

Host OfficeIMO in your own service when you need larger inputs, authentication, queues, storage, audit logs, or formats that are not suitable for WebAssembly. In that model, your organization owns the transport, access, logging, and retention policy.

## Publishing contract

The website pipeline builds the converter from its project source and mounts the published `wwwroot` output under `/apps/officeimo-converter/`. This keeps the deployed WebAssembly assets and integrity metadata aligned with the source in the same build.

The production-shaped publish relinks the converter's native WebAssembly assets, including HarfBuzz. Install the matching SDK's `wasm-tools` workload once before running it:

```powershell
dotnet workload install wasm-tools
pwsh -NoProfile -File .\Website\build.ps1 -PowerForgeRoot C:\path\to\PSPublishModule
```

Use `dotnet workload list` to verify the workload on an existing build machine. Use `-Dev` for a faster content/API build. The converter project itself can be checked directly with:

```powershell
dotnet build .\Website\Apps\OfficeIMO.Web.Converter\OfficeIMO.Web.Converter.csproj -c Release
dotnet test .\Website\Apps\OfficeIMO.Web.Converter.Tests\OfficeIMO.Web.Converter.Tests.csproj -c Release
```


## Continue with a result

The workspace keeps original and working files in memory for the current tab. After an operation, select **Use result as working file**, then choose a compatible tool. **Restore originals** returns to the uploaded source files. **Clear session** removes the session selection; downloaded files remain on your device. Archive bundles and reports are downloadable outputs and cannot become working documents. Use **Focus result** to give the preview more space.

## Inspect and remove provenance

Choose **Provenance** from the tools menu. JPEG, PNG, WebP, PDF, DOCX, XLSX, and PPTX files can be inspected locally. The sample image includes a deliberately added AI source declaration for trying the workflow.

1. Choose a file and inspect its supported provenance carriers.
2. Select embedded Content Credentials, external credential references, or AI-specific IPTC source declarations.
3. Create a separate cleaned copy. Review remaining findings and diagnostics from re-inspection.
4. Download the copy and report, or reuse the copy as the working file.

This is structural inspection, not cryptographic authenticity verification or a verdict about AI authorship. External references are not fetched. Cleanup does not remove visible marks or reconstruct image pixels, and it is not a general personal-metadata scrubber. Ambiguous carriers remain preserved; mutations that would invalidate document signatures are blocked. Input and output are capped at 25 MiB, with additional limits on expanded package data and embedded assets.
