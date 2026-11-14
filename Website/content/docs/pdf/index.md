---
title: "PDF"
description: "Create, inspect, transform, render, sign, and validate PDF documents with OfficeIMO.Pdf. Includes examples and API links."
order: 35
meta.seo_title: "PDF authoring and conversion | OfficeIMO"
verified: 2026-11-14
---

Create, read and change PDF files in .NET with OfficeIMO's own PDF engine — no third-party PDF library, browser or Office installation.

- **Create** reports, invoices and forms from headings, paragraphs, tables, lists, images, drawings, headers and footers, with embedded fonts.
- **Read** text, tables, images, links, attachments and form values, or export the content as Markdown or JSON.
- **Merge, split and organize** pages: extract, reorder, rotate, crop, duplicate, delete, stamp and watermark.
- **Fill forms** and keep them editable, or flatten the values into the page.
- **Protect** files with AES-256 passwords and permissions, or remove protection with the owner password.
- **Sign** with certificate-based signatures (PAdES B-B and B-T) and validate existing signatures.
- **Redact** text and areas permanently, with a verification report for every removed match.
- **Produce standard output**: PDF/A, PDF/UA, PDF/X, Factur-X and ZUGFeRD, checked against the saved bytes, and convert existing PDFs to PDF/A with a report of every change.
- **Render** pages to PNG, JPEG, TIFF, WebP or SVG, and compare two PDFs page by page.
- **Print** to installed printers on Windows, macOS and Linux, with page ranges, 1, 2, 4, 6 or 9 pages per sheet, scaling, duplex and paper source, through `OfficeIMO.Workflows`.
- **Convert** Word, Excel, PowerPoint, HTML and Markdown to PDF, and PDF back to Word, Excel, PowerPoint, HTML or Markdown.

## Task guides

- [Quick start](/docs/pdf/quick-start/): create a PDF and read it back.
- [Merge, split, and select pages](/docs/pdf/merge-split/)
- [Extract text and tables](/docs/pdf/extract-text-tables/)
- [Fill and flatten forms](/docs/pdf/forms/)
- [Redact content](/docs/pdf/redaction/)
- [Encrypt, decrypt, and sign](/docs/pdf/encrypt-sign/)
- [Convert to PDF/A](/docs/pdf/pdfa/)
- [Watermarks, stamps, and metadata](/docs/pdf/watermarks-stamps/)

## Install

```shell
dotnet add package OfficeIMO.Pdf
```

`OfficeIMO.Pdf` depends only on `OfficeIMO.Core`. Add a format adapter such as `OfficeIMO.Word.Pdf` when the source is another document model, and `OfficeIMO.Security` for certificate signatures.

## Choose a workflow

| Goal | Guide |
|---|---|
| Build a PDF report, form, or positioned layout | [Authoring and layout](/docs/pdf/authoring/) |
| Inspect, extract, merge, split, stamp, optimize, convert to PDF/A, repair, redact, or print | [Inspection, extraction, and operations](/docs/pdf/operations/) |
| Encrypt, sign, verify, or preserve existing signed revisions | [Security and digital signatures](/docs/pdf/security/) |
| Convert Word, Excel, PowerPoint, HTML, Markdown, or open formats | [Conversion and delivery](/docs/pdf/conversion/) |

The [PDF API reference](/api/pdf/) is generated from the built assembly and is the complete type-level index.

## Create a PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .AddParagraph("All scheduled checks passed. Two follow-ups remain open.")
        .AddTable(new[] {
            new[] { "Task", "Owner", "Due" },
            new[] { "Rotate certificates", "Platform", "4 Nov" } })))
    .Meta(title: "Service review", author: "Operations")
    .Save("Service-review.pdf");
```

## Read text and tables

```csharp
using OfficeIMO.Pdf;

PdfDocumentReadResult result = PdfDocument.Load("Service-review.pdf").Read();
foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
```

## Pick the correct package

| Starting point | Package |
|---|---|
| Author or edit PDF | `OfficeIMO.Pdf` |
| Word document | `OfficeIMO.Word.Pdf` |
| Excel workbook | `OfficeIMO.Excel.Pdf` |
| PowerPoint presentation | `OfficeIMO.PowerPoint.Pdf` |
| HTML | `OfficeIMO.Html.Pdf` |
| Markdown | `OfficeIMO.Markdown.Pdf` |
| OpenDocument | `OfficeIMO.OpenDocument.Odt.Pdf`, `.Ods.Pdf`, or `.Odp.Pdf` |

Use the [PDF API reference](/api/pdf/) for type-level detail. For conversion, inspect returned diagnostics before accepting approximated or omitted features.

## Production checks

- Preflight incoming PDFs before extraction or mutation.
- Separate “can read” from “can safely rewrite.”
- Apply explicit limits to untrusted page trees, streams, attachments, and decompression.
- Validate font availability and international-text shaping on the deployment host.
- Verify redaction by re-extracting content; an opaque drawing alone is not redaction.
- Treat signature structure, cryptographic validity, certificate trust, later revisions, and certification permissions as separate decisions.
- Reopen generated output and assert the semantic features your workflow depends on.
