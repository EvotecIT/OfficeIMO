# OfficeIMO.AsciiDoc.Pdf

`OfficeIMO.AsciiDoc.Pdf` is the direct, loss-aware AsciiDoc PDF adapter. It does not introduce another renderer: native AsciiDoc is projected by `OfficeIMO.AsciiDoc.Markdown`, then rendered by `OfficeIMO.Markdown.Pdf` and the shared first-party PDF engine.

```csharp
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Pdf;

AsciiDocDocument document = AsciiDocDocument.Load("guide.adoc").Document;
PdfSaveResult result = document.SaveAsPdf("guide.pdf");

result.Report.RequireNoLoss(); // optional strict conversion gate
result.Pipeline.RequireSuccess(); // exact output pipeline gate
```

`PdfSaveResult` combines native parser, semantic projection, PDF layout/resource/font diagnostics, and exact output-pipeline evidence. Use `ToPdfDocumentResult(...)` when conversion and post-processing should happen before save. Unsupported constructs remain visible when the projection policy allows source fallbacks; simplification, fallback, and omission are never silently reported as exact conversion.

The zero-options resource policy is inherited from `MarkdownToPdfOptions`: system fonts and bounded in-source resources are allowed, while arbitrary local and remote reads require explicit trust configuration.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.AsciiDoc.Markdown` and `OfficeIMO.Markdown.Pdf`; those packages retain ownership of semantic projection and PDF rendering.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.AsciiDoc.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
