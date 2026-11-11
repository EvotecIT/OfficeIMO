---
title: Browser Document Tools
description: Run supported OfficeIMO conversions and focused PDF workflows locally in your browser through the WebAssembly app on OfficeIMO.com.
order: 90
---

The [browser tools](/convert/) are static pages, one per task, such as [Word to PDF](/browser/word-to-pdf/) or [Merge PDFs](/browser/merge-pdfs/). Each page is usable straight away. The OfficeIMO engine, compiled to WebAssembly, downloads in the background and runs in a Web Worker, so the page stays responsive while a file is processed. Selected file bytes are not uploaded to OfficeIMO.

Every tool answers in the same way: a one-sentence result, the facts that matter (pages, size, what was removed), a named list of anything worth checking, a preview, and one download button. Reports and extra files sit under **More files and next steps**. The **Continue with** links carry the result into the next tool without downloading it first.

## ChatGPT Website Tool

Open any tool page in a browser that supports Website Tools to make `convert_selected_document` available on that page. Choose or drop a file and set the options on the page first. ChatGPT can then run the same action, and the result, warnings, and download stay visible on the page.

The tool has no file, path, URL, or format arguments. It can act only on the file and settings already chosen by the user, and processing remains browser-local. It returns bounded metadata rather than document contents. Leaving the page unregisters the tool. Cancellation is honored before the tool starts; once the browser-local run has started, it completes atomically.

## Supported browser routes

| Source | Output | Engine |
|---|---|---|
| DOCX | PDF | `OfficeIMO.Word.Pdf` |
| XLSX | PDF | `OfficeIMO.Excel.Pdf` |
| PPTX | PDF | `OfficeIMO.PowerPoint.Pdf` |
| [PUB](/browser/publisher-to-pdf/) | PDF | `OfficeIMO.Publisher.Pdf` |
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

Twelve task-oriented PDF tools are backed by `OfficeIMO.Pdf`. The page tools (extract, delete, reorder, rotate) show page thumbnails: click pages to select them or drag them into a new order. You can still type ranges such as `1-3,5,last`. Redaction searches first and lists every match by page before anything is removed.

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

Text files are limited to 1 MiB. HTML and Markdown conversion accepts up to 500,000 characters; the hidden-character tool accepts up to 1,048,576 UTF-16 code units. Oversized pasted text stays in the editor with an error, so you can shorten it without losing its ending. The directory offers only tools whose input limits the selected file meets.

Excel workbooks that pass those package checks are converted in full while every sheet's used range stays within 50,000 cells. If a sheet exceeds that budget, the app automatically generates a preview of up to 250 rows per sheet. Conversion warnings stay visible with the result instead of being hidden behind a successful download.

## Privacy and hosting

Browser-local processing is the strongest privacy default for a public demo because document bytes do not cross a server boundary. It is not the right execution model for every production workload.

Passwords stay in the page and the engine worker of the current tab, and are only used for the selected operation. Browser-local execution still means the user controls the device, browser extensions, downloads, and local storage policy.

The browser tools do not expose OCR, searchable-PDF generation, lossy scan compression, or cryptographic signing. Those capabilities need provider, quality, identity, or trust decisions that do not belong behind a generic one-click browser action.

Host OfficeIMO in your own service when you need larger inputs, authentication, queues, storage, audit logs, or formats that are not suitable for WebAssembly. In that model, your organization owns the transport, access, logging, and retention policy.

## Publishing contract

Tools are defined in `Website/data/browser_tools.json`, which drives the tool pages, the `/convert/` directory, and the engine wiring. The website pipeline builds the engine (`Website/Apps/OfficeIMO.Web.Converter`, a WebAssembly app with no UI) from its project source and mounts the published `wwwroot` output under `/apps/officeimo-converter/`. This keeps the deployed WebAssembly assets and integrity metadata aligned with the source in the same build.

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

After an operation, choose a compatible tool under **Continue with**. It opens in a new tab and receives the result through an in-memory transfer. Read-only inspections carry the selected input instead. The current tab keeps its input and result; choose a new input on that page to start again. Closing the tab releases its session, while downloaded files remain on your device. Archive bundles and reports are downloads rather than inputs for another tool.

## Clean a file before sharing

Open [Clean a file before sharing](/browser/clean-before-sharing/) from the browser directory. It reports everything that would leave with a DOCX, XLSX, PPTX, PDF, JPEG, PNG or WebP file, grouped by category, and saves a copy without what you tick. The sample is a short proposal with invented names, a comment, a tracked change, hidden text and a picture carrying a location and an AI label.

1. Choose a file. The check starts straight away and nothing is uploaded.
2. Review the groups: personal and hidden data, Content Credentials and AI labels, and for Word and PowerPoint, accessibility.
3. Untick anything you want to keep. Tracked changes, hidden content and origin records start unticked.
4. Download the clean copy. The tool checks the copy again and lists what was removed, what was kept and why.

[Prepare a document for sharing](/docs/before-sharing/) explains the defaults. Input and output are capped at 25 MiB; signed documents are not changed.
