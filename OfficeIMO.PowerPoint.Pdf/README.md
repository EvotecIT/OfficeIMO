# OfficeIMO.PowerPoint.Pdf - PowerPoint to PDF export

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint.Pdf)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint.Pdf)

`OfficeIMO.PowerPoint.Pdf` exports `OfficeIMO.PowerPoint` presentations to PDF through the first-party `OfficeIMO.Pdf` engine. In the reverse direction it reconstructs supported text, tables, safe shapes, and images as editable slide objects by default, with explicit visual, hybrid, and tables-only profiles.

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

## Import PDF pages

The general PDF route reconstructs supported page content as native slide objects. This is a new semantic projection, not recovery of the original slide deck: unsupported or ambiguous content remains explicit in the conversion report.

```csharp
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("handout.pdf");
PdfPowerPointConversionReport report = pdf.SaveAsPowerPoint("handout-editable.pptx");

foreach (var page in report.EditablePages) {
    Console.WriteLine(
        $"Page {page.PageNumber}: {page.TextBoxCount} text boxes, " +
        $"{page.TableCount} tables, {page.ShapeCount} shapes, {page.ImageCount} images");
}
```

Use the explicit visual profile when a page image is the intended result. Each image is movable and resizable, but text, vectors, charts, and tables inside it are not editable:

```csharp
var visual = PdfPowerPointImportOptions.CreateVisualPages();
PdfPowerPointConversionReport visualReport = pdf.SaveAsPowerPoint(
    "handout-visual.pptx",
    visual);

foreach (var page in visualReport.VisualPages) {
    Console.WriteLine($"PDF page {page.PageNumber}, slide {page.SlideIndex + 1}");
}
```

This is a new semantic projection, not recovery of the original slide deck. Original charts, groups, themes, animations, notes, and authoring intent cannot be recovered reliably from arbitrary PDFs; omissions and simplifications remain explicit warnings.

Use hybrid mode when the original page must remain visible while detected tables stay editable. Row and column caps split a large overlay across duplicate visual-page slides, and each overlay keeps the same centered, aspect-preserving page geometry as its background:

```csharp
var hybrid = PdfPowerPointImportOptions.CreateHybrid();
hybrid.MaxRowsPerSlide = 18;
hybrid.MaxColumnsPerSlide = 6;

PdfPowerPointConversionReport hybridReport = pdf.SaveAsPowerPoint(
    "handout-hybrid.pptx",
    hybrid);

Console.WriteLine($"Editable table segments: {hybridReport.TableEntries.Count}");
Console.WriteLine($"Visual-only page content: {hybridReport.HasNonEditablePageContent}");
```

Use editable-table mode when detected data is more important than page appearance:

```csharp
var options = PdfPowerPointImportOptions.CreateEditableTables();
options.MaxRows = 400;
options.MaxRowsPerSlide = 18;
options.MaxColumnsPerSlide = 6;

PdfPowerPointConversionReport report = pdf.SaveAsPowerPoint(
    "financial-statement-tables.pptx",
    options);

foreach (var table in report.TableEntries) {
    Console.WriteLine($"Page {table.PageNumber}, slide {table.SlideIndex + 1}");
}

Console.WriteLine($"Non-table page content detected: {report.HasOmittedPageContent}");
```

## Current limits

- Presentation content comes from `OfficeIMO.PowerPoint`; layout and PDF writing use `OfficeIMO.Pdf`.
- `PdfPowerPointImportMode.Auto` is the options default. It resolves an opened PDF to `EditableContent` and an already reduced `PdfLogicalDocument` to `EditableTables`; use `CreateVisualPages()` only when one rendered page image per slide is the intended output.
- `PdfPowerPointImportMode.EditableContent` reconstructs text blocks, detected tables, safe vector primitives, and supported images as native slide objects and reports anything it cannot represent safely.
- `PdfPowerPointImportMode.EditableTables` reconstructs detected tables and uses `SourceScope` / `HasOmittedPageContent` to expose unrelated page content.
- `PdfPowerPointImportMode.HybridVisualAndEditableTables` retains each selected page as a visual layer and overlays bounded editable table segments at source-relative geometry.
- The visual and hybrid modes accept caller-supplied fallback fonts for Base-14 and other unembedded font programs. Renderer capability diagnostics remain visible because a fallback is still a substitution, not the source font program.
- Navigation, groups, forms/controls, annotations, interactive media/animations, and complex vector or image placements are not claimed as editable slide objects; stable report warnings identify their visual-only, simplified, or omitted disposition.

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
- **OfficeIMO:** `OfficeIMO.PowerPoint`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own slide snapshots, PDF rendering, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
