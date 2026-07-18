# OfficeIMO.OneNote.Pdf

`OfficeIMO.OneNote.Pdf` converts the typed offline OneNote model to PDF without Microsoft Graph, a OneNote installation, or a commercial dependency. It provides a semantic document path and a position-preserving page path.

## Semantic PDF

The semantic path projects OneNote through `OfficeIMO.OneNote.Markdown`, then uses the first-party `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf` engines. It favors selectable text and normal document reflow:

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
byte[] pdf = section.ToPdf();
section.SaveAsPdf("Section.pdf");
```

OneNote pages are free-form canvases. The current `SemanticDocument` mode intentionally flattens them into reading order. `ToPdfDocumentResult()` reports canvas flattening, formatting simplification, unresolved asset placeholders, link-only binary assets, opaque omissions, and source diagnostics. It does not claim pixel parity with the OneNote desktop canvas.

Use `OneNotePdfSaveOptions.ProjectionOptions` for conflict/version inclusion and asset destinations, and `OneNotePdfSaveOptions.PdfOptions` for PDF layout, fonts, image policy, and diagnostics.

OneNote PDF export adds multilingual fallback candidates in addition to the normal document, monospace, and symbol candidates. The balanced default uses installed fonts while denying arbitrary local and remote reads; portable deterministic mode is explicit. Conversion clones both projection and PDF options so reusable caller configuration is not mutated:

```csharp
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;

var options = new OneNotePdfSaveOptions {
    PdfOptions = new MarkdownPdfSaveOptions {
        ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
    }
};

PdfDocumentConversionResult result = section.ToPdfDocumentResult(options);
foreach (PdfConversionWarning warning in result.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
result.Save("Section.pdf");
```

Invalid Unicode noncharacters and native control codes are normalized by `OfficeIMO.OneNote.Markdown` before layout. Valid characters that remain uncovered by the configured fonts still produce the normal PDF conversion diagnostic rather than being silently discarded.

## Visual page PDF

The visual path uses the same positioned `OfficeDrawing` canvas as PNG/JPEG/TIFF/SVG/WebP and visual HTML export. Each page is rasterized once and placed edge-to-edge at the canvas's page geometry, retaining freeform placement, images and printouts, ink, structured math, and attachment placeholders.

```csharp
var visualOptions = new OneNoteVisualPdfOptions {
    RasterScale = 2, // 144 DPI
    MaximumRasterPixels = 50_000_000,
    PageRendering = new OneNotePageRenderingOptions {
        // Optional: decoder for additional embedded source image formats.
        ImageCodec = applicationImageCodec
    },
    Title = "Project notebook"
};

byte[] visualPdf = section.ToVisualPdf(visualOptions);
var result = section.SaveAsVisualPdf("Section-visual.pdf", visualOptions);
foreach (var warning in result.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

`RasterScale` defaults to 2. If a page would exceed `MaximumRasterPixels`, the Drawing-owned limiter reduces its scale and reports `ONENOTE_PDF_RASTER_SCALE_LIMITED`. Embedded pictures that neither Drawing nor `PageRendering.ImageCodec` can decode are replaced visibly and reported as `DRAWING_RASTER_IMAGE_UNSUPPORTED`. Visual PDF is image-backed and does not provide selectable page text; use semantic PDF when search, selection, or reflow is more important than the original canvas layout.
