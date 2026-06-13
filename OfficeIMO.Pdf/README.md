# OfficeIMO.Pdf - Dependency-free PDF engine

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

`OfficeIMO.Pdf` is the first-party PDF package for OfficeIMO. It creates, reads, inspects, edits, merges, splits, stamps, and exports PDFs without runtime package dependencies.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Pdf
```

## Quick start

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(new PdfOptions {
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 11
    })
    .Meta(title: "Hello PDF", author: "OfficeIMO")
    .H1("OfficeIMO.Pdf")
    .Paragraph(p => p
        .Text("A dependency-free PDF builder with ")
        .Bold("rich text")
        .Text(", links, tables, images, and document operations."))
    .Table(new[] {
        new[] { "Area", "Status" },
        new[] { "Runtime dependencies", "None in OfficeIMO.Pdf" },
        new[] { "License", "MIT" }
    })
    .Save("hello.pdf");
```

## What it does

- Creates PDFs with page setup, headings, paragraphs, rich text, links, lists, panels, rows/columns, tables, images, vector drawing, headers, footers, watermarks, metadata, and form primitives.
- Reads and inspects PDFs through text extraction, logical document objects, page metadata, links, images, attachments, outlines, forms, active-content diagnostics, and security/revision markers.
- Manipulates existing PDFs with page extraction, split, merge, delete, duplicate, move, rotate, metadata editing, stamps, and watermarks.
- Provides conversion reports and diagnostics so adapters can expose unsupported or simplified source content honestly.
- Serves as the shared engine for Word, Excel, Markdown, HTML, and PowerPoint PDF adapters.

## Existing PDF workflows

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("input.pdf")
    .Pages.Extract("1-2,4")
    .MergeWith("appendix.pdf")
    .UpdateMetadata(title: "Merged report")
    .Stamp.Text("Reviewed")
    .Save("output.pdf");

string text = PdfDocument.Open("output.pdf").Read.Text();
```

## Examples

### Write a generated PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(new PdfOptions {
        PageSize = PageSizes.A4,
        Margins = PageMargins.UniformCentimeters(1.6),
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 10
    })
    .Meta(
        title: "Service report",
        author: "OfficeIMO",
        subject: "Generated PDF")
    .Header(h => h.AlignCenter().Text("Service report"))
    .Footer(f => f.AlignRight().Text("Page {page} of {pages}"))
    .H1("Service report")
    .Paragraph(p => p
        .Text("Generated ")
        .Bold(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm 'UTC'"))
        .Text(" with first-party PDF primitives."))
    .Table(new[] {
        new[] { "System", "Status", "Owner" },
        new[] { "Identity", "Green", "Operations" },
        new[] { "Messaging", "Yellow", "Exchange" }
    })
    .Save("service-report.pdf");
```

### Rich report layout

```csharp
PdfDocument.Create()
    .H1("Operational summary")
    .Paragraph(p => p
        .Text("Generated ")
        .Bold(DateTime.Today.ToString("yyyy-MM-dd"))
        .Text(" with links, lists, panels, and tables."))
    .Bullets(list => list
        .Item("No runtime package dependencies")
        .Item("Word-like document flow")
        .Item("Reusable PDF primitives for adapters"))
    .Panel(panel => panel
        .H2("Review note")
        .Paragraph(p => p.Text("Keep polished report designs in samples; keep reusable primitives in the engine.")))
    .Table(new[] {
        new[] { "Area", "Status" },
        new[] { "Layout", "Ready" },
        new[] { "Reading", "Evolving" }
    })
    .Save("summary.pdf");
```

### Read text, Markdown, tables, images, and attachments

```csharp
using OfficeIMO.Pdf;

using var pdf = PdfDocument.Open("statement.pdf");

string text = pdf.Read.Text();
string firstPages = pdf.Read.Text("1-2");
string markdown = pdf.Read.Markdown();
IReadOnlyList<string> pages = pdf.Read.TextByPage();
PdfLogicalDocument logical = pdf.Read.Logical();

foreach (var table in logical.Tables) {
    Console.WriteLine($"Table on page {table.PageNumber}: {table.Rows.Count} rows");
}

string markdownTables = PdfLogicalTableTextExport.ExtractMarkdownTables("statement.pdf");
IReadOnlyList<PdfExtractedImage> images = pdf.Read.Images();
IReadOnlyList<PdfExtractedAttachment> attachments = pdf.Read.Attachments();
```

### Split and extract pages

```csharp
using OfficeIMO.Pdf;

using var source = PdfDocument.Open("packet.pdf");

source.Pages.Extract("1-3")
    .Save("cover-and-summary.pdf");

IReadOnlyList<PdfDocument> singlePageDocuments = source.Pages.Split();
for (int index = 0; index < singlePageDocuments.Count; index++) {
    singlePageDocuments[index].Save($"packet-page-{index + 1:000}.pdf");
}
```

### Merge, reorder, delete, duplicate, move, and rotate

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("packet.pdf")
    .MergeWith("appendix.pdf")
    .Pages.Delete("2,5-6")
    .Pages.Duplicate("1")
    .Pages.Move(insertBeforePageNumber: 3, pageRanges: "7-8")
    .Pages.Rotate(90, "4")
    .UpdateMetadata(title: "Cleaned packet")
    .Save("packet-clean.pdf");
```

### Stamp and watermark an existing PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("contract.pdf")
    .Stamp.Text("Reviewed", new PdfTextStampOptions {
        X = 72,
        Y = 720,
        FontSize = 18,
        Color = PdfColor.FromRgb(180, 30, 30)
    })
    .Stamp.TextWatermark("CONFIDENTIAL", new PdfTextStampOptions {
        FontSize = 54,
        Color = PdfColor.Gray,
        RotationDegrees = -35
    })
    .Save("contract-reviewed.pdf");
```

### Fill and flatten a PDF form

```csharp
using OfficeIMO.Pdf;

PdfDocument.Open("application-form.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Applicant.Email"] = "adele@example.com",
        ["Approval.Status"] = "Approved"
    })
    .Save("application-form-filled.pdf");
```

### Page setup, watermarks, and metadata

```csharp
PdfDocument.Create(new PdfOptions {
        PageSize = PageSize.FromCentimeters(21, 29.7).Portrait(),
        Margins = PageMargins.UniformCentimeters(1.5),
        TextWatermark = new PdfTextWatermark("DRAFT") {
            Opacity = 0.12,
            RotationAngle = -35
        }
    })
    .Meta(title: "Draft report", author: "OfficeIMO")
    .H1("Draft report")
    .Paragraph("This document uses page-level options instead of post-processing.")
    .Save("draft.pdf");
```

### Inspect and preflight before rewriting

```csharp
using OfficeIMO.Pdf;

byte[] bytes = File.ReadAllBytes("incoming.pdf");
PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);

if (!preflight.Can(PdfPreflightCapability.ManipulatePages)) {
    foreach (string diagnostic in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages)) {
        Console.WriteLine(diagnostic);
    }
}

var result = PdfDocument.Open(bytes).Pages.TryExtract(PdfPageSelection.Parse("1-2"));
if (result.Succeeded) {
    result.RequireValue().Save("incoming-first-pages.pdf");
}
```

### Inspect before automating

```csharp
using var pdf = PdfDocument.Open("incoming.pdf");

var inspection = pdf.Inspect();
Console.WriteLine($"Pages: {inspection.PageCount}");
Console.WriteLine($"Links: {inspection.LinkAnnotationCount}");
Console.WriteLine($"Forms: {inspection.FormFields.Count}");
Console.WriteLine($"Active content: {inspection.HasActiveContent}");

foreach (var page in inspection.Pages) {
    Console.WriteLine($"{page.PageNumber}: {page.Width} x {page.Height}");
}
```

### Convert PDFs through adapter packages

```csharp
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var word = WordDocument.Load("proposal.docx");
word.SaveAsPdf("proposal.pdf");

PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
    "bank-statement.pdf",
    "bank-statement-tables.xlsx");

PdfHtmlConverter.SaveAsHtml(
    "proposal.pdf",
    "proposal-review.html",
    new PdfHtmlSaveOptions {
        Profile = PdfHtmlProfile.PositionedReview,
        IncludeLinkAnnotations = true,
        IncludeFormWidgets = true
    });
```

## Conversion adapters

| Package | Role |
| --- | --- |
| [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md) | Maps Word documents into PDF primitives. |
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Maps Excel workbooks into PDF primitives. |
| [OfficeIMO.Markdown.Pdf](../OfficeIMO.Markdown.Pdf/README.md) | Maps Markdown documents into PDF primitives. |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Maps PowerPoint slides into PDF primitives. |
| [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) | Bridges HTML to PDF and PDF to HTML. |

## Boundaries

- `OfficeIMO.Pdf` should stay dependency-free at runtime. Rasterizers, visual comparison tools, and external renderers belong in tests or development tooling.
- Polished invoice, report, and statement examples belong in samples and visual fixtures, not as special engine concepts.
- Adapter-specific mapping belongs in the source adapter packages. Shared PDF layout, reading, and manipulation behavior belongs here.
- Current-state inventories belong in [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md), not in this NuGet README.

## Current state

The PDF engine is useful and broad, but it is still evolving. It has strong first-party coverage for common generated business documents and conservative read/manipulation workflows, while advanced typography, complex PDF preservation, encryption/decryption, and signature validation remain deeper roadmap areas.

For the full capability inventory and roadmap, read [Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md).

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
