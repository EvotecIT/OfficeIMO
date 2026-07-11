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

using var presentation = PowerPointPresentation.Open("deck.pptx");
presentation.SaveAsPdf("deck.pdf");
```

## Examples

### Export with slide-content controls

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Open("board-review.pptx");

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

using var presentation = PowerPointPresentation.Open("training.pptx");

byte[] pdfBytes = presentation.ToPdf();

using var stream = File.Create("training.pdf");
presentation.SaveAsPdf(stream);
```

### Export speaker notes and handouts

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointPresentation.Open("training.pptx");

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

using var presentation = PowerPointPresentation.Open("complex-deck.pptx");
var options = new PowerPointPdfSaveOptions {
    IncludeCharts = true,
    IncludeAutoShapes = true
}.UseProfile(PdfExportProfile.Faithful);

options.TextFallbacks = PdfTextFallbackFeatures.Default;
options.AllowSystemFontEmbedding = true;

var result = presentation.TrySaveAsPdf("complex-deck.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in options.ConversionReport.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}

options.ConversionReport.RequireNoErrorWarnings();
```

## What it maps

- Full-slide pages use the authored slide size; notes pages use portrait letter and handouts use landscape letter.
- Slide backgrounds, text boxes, supported pictures, supported tables, supported charts, and basic auto-shapes.
- Text box fill, outline, margins, font defaults, alignment, vertical anchoring, rich runs, and hyperlinks.
- Supported JPEG/PNG pictures through the shared PDF image pipeline.
- Faithful and print-ready profiles consume the same shared visual snapshot as PNG/SVG and visual-review HTML. Selective profiles retain the per-shape path.
- Profile presets through `PowerPointPdfSaveOptions.UseProfile(...)`, plus shared `TextFallbacks` and `AllowSystemFontEmbedding` controls for Unicode, symbols, and emoji.
- Conversion warnings through `PowerPointPdfSaveOptions.Warnings` and `PowerPointPdfSaveOptions.ConversionReport`.

## PDF table import

`SavePdfTablesAsPowerPoint(...)` extracts logical tables from a PDF and writes editable PowerPoint table slides. This is useful for review decks and migration workflows where the source PDF has table-like content that should become editable again.

```csharp
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;

var imported = PowerPointPdfConverterExtensions.SavePdfTablesAsPowerPoint(
    "financial-statement.pdf",
    "financial-statement-tables.pptx",
    new PdfPowerPointTableImportOptions {
        PageRanges = new[] { PdfPageRange.From(2, 5) },
        MaxRows = 400,
        MaxRowsPerSlide = 18,
        MaxColumnsPerSlide = 6
    });

foreach (var table in imported) {
    Console.WriteLine($"Page {table.PageNumber}, slide {table.SlideIndex + 1}");
}
```

## Boundaries

- Presentation modeling stays in `OfficeIMO.PowerPoint`.
- PDF layout and writing stay in `OfficeIMO.Pdf`.
- This package should remain a thin adapter over shared PDF primitives.
- Complex slide fidelity gaps should be reported through warnings and deeper docs rather than broad README claims.

## Related packages

- [OfficeIMO.PowerPoint](../OfficeIMO.PowerPoint/README.md) - PowerPoint presentation model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF engine.
- [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) - Markup to PowerPoint rendering.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
