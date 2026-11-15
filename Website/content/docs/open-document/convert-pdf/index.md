---
title: Convert OpenDocument to PDF and Images
description: Convert ODT, ODS and ODP files to PDF or page images, and rebuild editable OpenDocument files from a PDF, with a report of both conversion stages.
order: 41
verified: 2026-11-14
commands:
  - officeimo convert
packages:
  - OfficeIMO.OpenDocument.Odt.Pdf
  - OfficeIMO.OpenDocument.Ods.Pdf
  - OfficeIMO.OpenDocument.Odp.Pdf
cmdlets:
  - Convert-OfficeDocument
related:
  - /docs/open-document/
  - /docs/open-document/convert-office/
  - /docs/pdf/conversion/
  - /docs/cli/convert/
limits: PDF to ODT, ODS or ODP is semantic reconstruction through the Word, Excel or PowerPoint models, not a visual copy; check both stage reports when fidelity matters.
search_terms:
  - odt to pdf
  - ods to pdf
  - odp to pdf
  - pdf to odt
  - opendocument png
---

PDF export renders through the same Word, Excel and PowerPoint engines the Office formats use, so pick the package for the format you have: `OfficeIMO.OpenDocument.Odt.Pdf`, `.Ods.Pdf` or `.Odp.Pdf`. Each brings only its own engine, with no LibreOffice process. The returned report holds the diagnostics of both stages: the OpenDocument import and the PDF rendering.

## Convert ODT, ODS and ODP to PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;

using (OdtDocument text = OdtDocument.Load("proposal.odt")) {
    OfficeConversionReport report = text.SaveAsPdf("proposal.pdf", new OdtPdfExportOptions { FailOnLoss = true });
    Console.WriteLine($"proposal.pdf: {report.Diagnostics.Count} diagnostics");
}

using (OdsDocument workbook = OdsDocument.Load("budget.ods")) {
    workbook.SaveAsPdf("budget.pdf");
}

using (OdpPresentation presentation = OdpPresentation.Load("slides.odp")) {
    presentation.SaveAsPdf("slides.pdf");
}
```

`OdtPdfExportOptions`, `OdsPdfExportOptions` and `OdpPdfExportOptions` take `ConversionOptions` for the import stage, `PdfOptions` for the PDF engine and `FailOnLoss`. `ToPdf` returns a `PdfDocument` instead of writing a file, for example to merge it with other PDFs.

### PowerShell

```powershell
Convert-OfficeDocument -Path ./proposal.odt -OutputPath ./proposal.pdf
Get-ChildItem ./reports -Filter *.ods | Convert-OfficeDocument -To Pdf -OutputDirectory ./pdf
```

### CLI

```sh
officeimo convert proposal.odt proposal.pdf
officeimo convert budget.ods budget.pdf
officeimo convert slides.odp slides.pdf --require-no-loss
```

## Render pages as images

The Word, Excel and PowerPoint OpenDocument adapters render pages to PNG, JPEG, WebP, TIFF or SVG through the same visual engines.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

using OdtDocument source = OdtDocument.Load("proposal.odt");
var options = new WordImageExportOptions { Format = OfficeImageExportFormat.Png, TargetDpi = 144 };

IReadOnlyList<OfficeImageExportResult> pages = source.ToImages(options);
Console.WriteLine($"{pages.Count} pages rendered");

OfficeConversionReport saved = await source.SaveAsImagesAsync("pages", options);
```

For spreadsheets and presentations, use `OdsDocument.ToImages` with `ExcelImageExportOptions` from `OfficeIMO.Excel.OpenDocument`, or `OdpPresentation.ToImages` with `PowerPointImageExportOptions` from `OfficeIMO.PowerPoint.OpenDocument`.

### CLI

```sh
officeimo convert proposal.odt proposal.png
```

## Rebuild an OpenDocument file from a PDF

The reverse direction reconstructs editable content: PDF to ODT goes through the PDF-to-Word and Word-to-ODT adapters, and the report keeps both stages.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;

PdfDocument contract = PdfDocument.Load("contract.pdf");
OfficeConversionReport report = contract.SaveAsOdt("contract.odt");
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) Console.WriteLine(diagnostic.Message);

PdfDocument.Load("statement.pdf").SaveAsOds("statement.ods");
PdfDocument.Load("slides.pdf").SaveAsOdp("slides.odp");
```

### CLI

```sh
officeimo convert contract.pdf contract.odt
officeimo convert statement.pdf statement.ods
```
