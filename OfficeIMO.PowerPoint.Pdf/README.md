# OfficeIMO.PowerPoint.Pdf - PowerPoint to PDF export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint.Pdf)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.Pdf)

`OfficeIMO.PowerPoint.Pdf` exports `OfficeIMO.PowerPoint` presentations to PDF through the first-party `OfficeIMO.Pdf` engine. It also imports logical PDF tables into editable PowerPoint table slides.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint.Pdf
```

## Quick start

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Load("deck.pptx");
presentation.SaveAsPdf("deck.pdf");
```

## Examples

### Export with slide-content controls

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Load("board-review.pptx");

var options = new PowerPointPdfSaveOptions {
    IncludeHiddenSlides = false,
    IncludeSlideBackgrounds = true,
    IncludePictures = true,
    IncludeTextBoxes = true,
    IncludeTables = true,
    IncludeCharts = true,
    WarnOnPictureAspectRatioDistortion = true
};

presentation.SaveAsPdf("board-review.pdf", options);
```

### Export to bytes or a stream

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Load("training.pptx");

byte[] pdfBytes = presentation.ToPdf();

using var stream = File.Create("training.pdf");
presentation.SaveAsPdf(stream);
```

### Export speaker notes and handouts

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Load("training.pptx");

presentation.SaveAsPdf("training-notes.pdf", new PowerPointPdfSaveOptions {
    PageLayout = PowerPointPdfPageLayout.NotesPages,
    IncludeSpeakerNotes = true
});

presentation.SaveAsPdf("training-handout.pdf", new PowerPointPdfSaveOptions {
    PageLayout = PowerPointPdfPageLayout.Handouts,
    HandoutSlidesPerPage = 3,
    IncludeSpeakerNotes = true
});
```

Handouts support 1, 2, 3, 4, 6, or 9 slides per landscape page. Three-up output pairs each thumbnail with notes or writing lines. Notes are read without creating missing notes parts.

### Review conversion warnings

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;

using var presentation = PowerPointPresentation.Load("complex-deck.pptx");
var options = new PowerPointPdfSaveOptions {
    IncludeCharts = true,
    IncludeAutoShapes = true
}.UseProfile(PdfExportProfile.Faithful);

options.TextFallbacks = PdfTextFallbackFeatures.Default;
options.ResourcePolicy = PdfResourcePolicy.CreateTrustedHost();

var result = presentation.TrySaveAsPdf("complex-deck.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in result.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}

result.Report.RequireNoErrorWarnings();
```

## What it maps

- Full-slide pages use the authored slide size; notes pages use portrait letter and handouts use landscape letter.
- Slide backgrounds, text boxes, supported pictures, supported tables, supported charts, and basic auto-shapes.
- Text box fill, outline, margins, font defaults, alignment, vertical anchoring, rich runs, and hyperlinks.
- Supported JPEG/PNG pictures through the shared PDF image pipeline.
- Full-slide PDF output always uses the native per-shape PDF renderer, including hyperlinks and rich text. Conversion no longer chooses a different renderer from document content or an option toggle.
- PNG, SVG, visual-review HTML, and notes/handout thumbnails use the shared visual snapshot; those surfaces have a different scene/raster contract and do not select the PDF engine at runtime.
- Profile presets through `PowerPointPdfSaveOptions.UseProfile(...)`, plus shared `TextFallbacks` and `ResourcePolicy` controls. The balanced default uses installed fonts while denying arbitrary local and remote reads; portable deterministic mode is explicit.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

## PDF table import

`SaveTablesAsPowerPoint(...)` extracts logical tables from a PDF and writes editable PowerPoint table slides. It does not claim to convert unrelated page text, images, links, forms, annotations, or actions.

```csharp
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;

PdfLogicalDocument source = PdfLogicalDocument.LoadPageRanges(
    "financial-statement.pdf",
    PdfPageRange.From(2, 5));

PdfPowerPointTableImportReport report = source.SaveTablesAsPowerPoint(
    "financial-statement-tables.pptx",
    new PdfPowerPointTableImportOptions {
        MaxRows = 400,
        MaxRowsPerSlide = 18,
        MaxColumnsPerSlide = 6
    });

foreach (var table in report.Entries) {
    Console.WriteLine($"Page {table.PageNumber}, slide {table.SlideIndex + 1}");
}

Console.WriteLine($"Non-table page content detected: {report.HasOmittedPageContent}");
```

## Boundaries

- Presentation modeling stays in `OfficeIMO.PowerPoint`.
- PDF layout and writing stay in `OfficeIMO.Pdf`.
- This package should remain a thin adapter over shared PDF primitives.
- PDF import is intentionally table-only. `SourceScope` and `HasOmittedPageContent` make the omitted page surface explicit.
- Complex slide fidelity gaps should be reported through warnings and deeper docs rather than broad README claims.

## Related packages

- [OfficeIMO.PowerPoint](../OfficeIMO.PowerPoint/README.md) - PowerPoint presentation model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF engine.
- [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) - Markup to PowerPoint rendering.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages; no native or commercial PDF renderer.
- **OfficeIMO:** `OfficeIMO.PowerPoint`, `OfficeIMO.Pdf`, and `OfficeIMO.Drawing` own slide snapshots, PDF rendering, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
