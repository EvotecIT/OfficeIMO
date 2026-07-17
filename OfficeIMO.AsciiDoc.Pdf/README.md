# OfficeIMO.AsciiDoc.Pdf

`OfficeIMO.AsciiDoc.Pdf` is the direct, loss-aware AsciiDoc PDF adapter. It does not introduce another renderer: native AsciiDoc is projected by `OfficeIMO.AsciiDoc.Markdown`, then rendered by `OfficeIMO.Markdown.Pdf` and the shared first-party PDF engine.

```csharp
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Pdf;

AsciiDocDocument document = AsciiDocDocument.Load("guide.adoc").Document;
PdfDocumentConversionResult result = document.SaveAsPdf("guide.pdf");

result.RequireNoLoss(); // optional strict gate
```

`PdfDocumentConversionResult` combines native parser diagnostics, semantic projection diagnostics, and PDF layout/resource/font diagnostics. Unsupported constructs remain visible when the projection policy allows source fallbacks; simplification, fallback, and omission are never silently reported as exact conversion.

The zero-options resource policy is inherited from `MarkdownPdfSaveOptions`: system fonts and bounded in-source resources are allowed, while arbitrary local and remote reads require explicit trust configuration.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.AsciiDoc.Markdown` and `OfficeIMO.Markdown.Pdf`; those packages retain ownership of semantic projection and PDF rendering.
