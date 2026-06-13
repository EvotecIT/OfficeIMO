# OfficeIMO.Word.Pdf - Word to PDF export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Pdf)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)

`OfficeIMO.Word.Pdf` exports `OfficeIMO.Word` documents to PDF through the first-party `OfficeIMO.Pdf` engine. It is the adapter layer: Word stays responsible for the `.docx` model, while PDF layout and writing stay in `OfficeIMO.Pdf`.

## Install

```powershell
dotnet add package OfficeIMO.Word.Pdf
```

## Quick start

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Create("report.docx");
document.AddParagraph("PDF export").SetStyle("Heading1");
document.AddParagraph("This document is exported through OfficeIMO.Pdf.");

document.SaveAsPdf("report.pdf");
```

## Examples

### Export with page and metadata options

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new PdfSaveOptions {
    Orientation = PdfPageOrientation.Portrait,
    Margins = PageMargins.UniformCentimeters(1.5),
    Title = "Customer proposal",
    Author = "Evotec",
    IncludePageNumbers = true,
    PageNumberFormat = "Page {current} of {total}"
};

document.SaveAsPdf("proposal.pdf", options);
```

### Export to bytes or streams

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("invoice.docx");

byte[] pdfBytes = document.SaveAsPdf();

using var stream = File.Create("invoice.pdf");
document.SaveAsPdf(stream);
```

### Capture conversion warnings without throwing away the report

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("complex-document.docx");
var options = new PdfSaveOptions {
    DefaultTableBorders = true
};

var result = document.TrySaveAsPdf("complex-document.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in options.ConversionReport.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}
```

### Import logical PDF tables into Word

```csharp
using OfficeIMO.Word.Pdf;

var imported = PdfWordTableConverterExtensions.SavePdfTablesAsWord(
    "statement.pdf",
    "statement-tables.docx");

foreach (var table in imported) {
    Console.WriteLine($"Page {table.PageNumber}, table {table.TableIndex + 1}");
}
```

## What it exports

- Paragraphs, headings, rich runs, links, bookmarks, page breaks, lists, and common spacing/indentation settings.
- Word sections, page size, orientation, margins, columns, headers, footers, page numbers, and document background color.
- Tables with common Word table styling, repeated headers, cell fills, borders, alignment, merged cells, and rich text in cells.
- Paragraph-aligned images, selected shapes, text boxes, content controls, simple form controls, footnote/endnote markers, and table-of-contents links where supported by the first-party PDF path.
- Conversion warnings through `PdfSaveOptions.Warnings` and `PdfSaveOptions.ConversionReport`.

## Options and diagnostics

Use `PdfSaveOptions` when callers need to override page geometry, metadata, page-number behavior, font family, or table-border fallback. Keep `PdfSaveOptions.Warnings` and `PdfSaveOptions.ConversionReport` visible in wrappers and user interfaces; unsupported Word features should become actionable diagnostics instead of silent README promises.

## Boundaries

- This package does not try to be a full Word renderer with perfect Microsoft Word parity.
- Unsupported or simplified Word features should surface warnings rather than being hidden in the README as broad claims.
- Reusable PDF layout work belongs in `OfficeIMO.Pdf`; Word-specific mapping belongs here.
- PowerShell PDF workflows should be exposed through [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice).

## Related packages

- [OfficeIMO.Word](../OfficeIMO.Word/README.md) - Word document model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF creation, reading, and manipulation engine.
- [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) - HTML/PDF bridge built on OfficeIMO converters.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
