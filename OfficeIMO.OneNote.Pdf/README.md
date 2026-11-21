# OfficeIMO.OneNote.Pdf

`OfficeIMO.OneNote.Pdf` converts the typed offline OneNote model to PDF without Microsoft Graph, a OneNote installation, or a commercial dependency. It provides a semantic document path and a position-preserving page path.

## Semantic PDF

The semantic path projects OneNote through `OfficeIMO.OneNote.Markdown`, then uses the first-party `OfficeIMO.Markdown.Pdf` and `OfficeIMO.Pdf` engines. It favors selectable text and normal document reflow:

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Pdf;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
OfficeConversionReport report = section.SaveAsPdf("Section.pdf");
using PdfDocument pdf = section.ToPdf();
```

OneNote pages are free-form canvases. The current `SemanticDocument` mode intentionally flattens them into reading order. The conversion report covers canvas flattening, formatting simplification, unresolved asset placeholders, link-only binary assets, opaque omissions, and source diagnostics. It does not claim pixel parity with the OneNote desktop canvas.

Use `OneNotePdfExportOptions.ProjectionOptions` for conflict/version inclusion and asset destinations, and `OneNotePdfExportOptions.MarkdownOptions` for the intermediate Markdown-to-PDF layout, fonts, image policy, and diagnostics.

OneNote PDF export adds multilingual fallback candidates in addition to the normal document, monospace, and symbol candidates. The balanced default uses installed fonts while denying arbitrary local and remote reads; portable deterministic mode is explicit. Conversion clones both projection and PDF options so reusable caller configuration is not mutated:

```csharp
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.Pdf;

var options = new OneNotePdfExportOptions {
    MarkdownOptions = new MarkdownPdfExportOptions {
        ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
    }
};

using PdfDocument pdf = section.ToPdf(options, out OfficeConversionReport report);
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
result.Save("Section.pdf");
```

Invalid Unicode noncharacters and native control codes are normalized by `OfficeIMO.OneNote.Markdown` before layout. Valid characters that remain uncovered by the configured fonts still produce the normal PDF conversion diagnostic rather than being silently discarded.

## Visual page PDF

The visual path uses the same positioned `OfficeDrawing` canvas as PNG/JPEG/TIFF/SVG/WebP and visual HTML export. Each page is written edge-to-edge at the canvas's page geometry as vector content: text stays real PDF text that can be searched, selected, and extracted, with the installed fonts it names embedded as subsets and Unicode fallbacks added when the text needs them. Freeform placement, tables, ink, structured math, and attachment placeholders are drawn as vectors, and embedded pictures keep their own bytes and alternative text.

```csharp
var visualOptions = new OneNoteVisualPdfExportOptions {
    PageRendering = new OneNotePageRenderingOptions {
        // Optional: decoder for additional embedded source image formats.
        ImageCodec = applicationImageCodec
    },
    Title = "Project notebook"
};

OfficeConversionReport visual = section.SaveAsVisualPdf("Section-visual.pdf", visualOptions);
foreach (var warning in visual.Diagnostics) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

JPEG and PNG pictures are embedded as they are, and other formats the PDF writer decodes itself are normalized by it. A picture in any other format is decoded by `PageRendering.ImageCodec` and embedded as PNG for that picture only, reported as `ONENOTE_PDF_IMAGE_RASTERIZED`; when nothing can decode it, a visible placeholder is drawn and `image-source-decode-fallback` is reported. The rest of the page stays vector.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OneNote.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
