# OfficeIMO.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

Create, read and change PDF files in .NET with OfficeIMO's own PDF engine: no third-party PDF library, browser or Office installation.

## Install

```sh
dotnet add package OfficeIMO.Pdf
```

Add an optional package only when you need its workflow:

| To | Add |
| --- | --- |
| Sign with certificates or validate CMS signatures | `OfficeIMO.Security` |
| Save Word, Excel or PowerPoint as PDF | `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf` |
| Convert HTML to PDF, or PDF to HTML | `OfficeIMO.Html.Pdf` |
| Save Markdown as PDF | `OfficeIMO.Markdown.Pdf` |
| Convert PDF to Word, Excel or PowerPoint | `OfficeIMO.Pdf.Word`, `OfficeIMO.Pdf.Excel`, `OfficeIMO.Pdf.PowerPoint` |
| Read scanned pages with an OCR provider | `OfficeIMO.Pdf.Ocr` |
| Render an invoice with its Factur-X XML | `OfficeIMO.Invoicing.Pdf` |
| Print to an installed printer | [`OfficeIMO.Workflows`](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Workflows/README.md#print-a-document) |
| Full OpenType shaping for complex scripts | `OfficeIMO.Drawing.HarfBuzz` |

## Quick start

Page numbers and ranges are 1-based. Page, merge and form operations return a new document; save it to a different path to keep the source file.

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Create" -->
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
<!-- /powerforge:code -->

Read it back as paragraphs, headings, lists and tables per page, or as Markdown:

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Read" -->
```csharp
PdfDocumentReadResult result = PdfDocument.Load("Service-review.pdf").Read();
foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
```
<!-- /powerforge:code -->

Use `PdfLoadOptions.UntrustedDefaults` for files from outside your organization, and `PdfDocument.LoadSeekable` to read a few pages of a large PDF without loading all of it.

## What it does

- **Create** reports, invoices and forms from headings, paragraphs, tables, lists, images, drawings, headers and footers, with embedded fonts.
- **Read** text, tables, images, links, attachments and form values, or export the content as Markdown or JSON.
- **Merge, split and organize** pages: extract, reorder, rotate, crop, duplicate, delete, stamp and watermark.
- **Fill forms** and keep them editable, or flatten the values into the page.
- **Protect** files with AES-256 passwords and permissions, or remove protection with the owner password.
- **Sign** with certificate-based signatures (PAdES B-B and B-T) and validate existing signatures.
- **Redact** text and areas permanently, with a verification report for every removed match.
- **Remove hidden data** before sharing: metadata, comments, scripts and other active content, attachments, bookmarks and layers, with the same `InspectHiddenData` and `RemoveHiddenData` report as the Office formats.
- **Produce standard output**: PDF/A, PDF/UA, PDF/X, Factur-X and ZUGFeRD, checked against the saved bytes.
- **Check accessibility** of any PDF against PDF/UA-1 with the Matterhorn Protocol conditions: `PdfDocument.Load(path).Proof.CheckAccessibility()` locates each failure by page, object and tag, and lists the checks a person still has to review.
- **Render** pages to PNG, JPEG, TIFF, WebP or SVG, and compare two PDFs page by page.
- **Redline** two versions of a document: compare their text word by word and mark insertions, deletions, replacements and moves on the newer PDF, with a JSON change list and a side-by-side HTML report.
- **Print** to installed printers on Windows, macOS and Linux, with page ranges, 1, 2, 4, 6 or 9 pages per sheet, scaling, duplex and paper source, through `OfficeIMO.Workflows`.
- **Convert** Word, Excel, PowerPoint, HTML and Markdown to PDF, and PDF back to Word, Excel, PowerPoint, HTML or Markdown.

The [PDF guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.pdf-guide.md) shows each of these with examples, from page layout and fonts to merging, forms, signatures, redaction, PDF/A, rendering and asset extraction.

## Limits

Each workflow reports what it couldn't do instead of silently dropping content: unsupported fonts, color profiles, transparency and form features show up as diagnostics. [PDF current state](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.pdf.current-state.md) describes what each workflow covers, and the [PDF conversion support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.pdf-conversion-support-matrix.md), [accessibility matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.pdf-accessibility-support-matrix.md) and [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) list the details.

## Learn more

- [OfficeIMO.Pdf on officeimo.com](https://officeimo.com/products/pdf/), the [PDF docs](https://officeimo.com/docs/pdf/) and the [browser PDF tools](https://officeimo.com/browser/).
- [PDF guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.pdf-guide.md).
- [Examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Pdf) and the [migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) (`New-OfficePdf`, `Open-OfficePdf`). Command line: `officeimo convert report.pdf report.md` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only, plus `System.Memory` on .NET Standard 2.0 and .NET Framework.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 3 | 0 | 0 | 0 | 0 | 0 |
| Read | 3 | 0 | 0 | 0 | 0 | 0 |
| Edit | 2 | 0 | 0 | 0 | 1 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 0 | 0 | 0 | 0 | 2 |
| Remove | 2 | 0 | 0 | 1 | 0 | 0 |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
