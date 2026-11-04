# OfficeIMO.Latex.Pdf

`OfficeIMO.Latex.Pdf` is the direct, loss-aware adapter for OfficeIMO's bounded LaTeX profile. It does not execute TeX or introduce another layout engine: `OfficeIMO.Latex.Markdown` owns semantic projection, and `OfficeIMO.Markdown.Pdf` plus the shared PDF engine own rendering.

```csharp
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Pdf;

LatexDocument document = LatexDocument.Load("article.tex");
OfficeConversionReport report = document.SaveAsPdf("article.pdf");

report.ThrowIfLoss(); // optional strict conversion gate
```

The report combines native parser, bounded-profile projection, PDF layout/resource/font diagnostics. `ToPdf(options, out report)` returns the PDF model with its `OfficeConversionReport`, `ToPdfAsync` returns both as `OfficeConversionResult<PdfDocument>`, and `SaveAsPdf`/`SaveAsPdfAsync` write a path or caller-owned stream and return the report. Set `FailOnLoss` on the options to make reported loss an `OfficeConversionException` before anything is written. TeX macros and package behavior are not executed; unsupported or simplified constructs are preserved visibly when configured and remain explicit warnings.

The zero-options resource policy is inherited from `MarkdownPdfExportOptions`: system fonts and bounded in-source resources are allowed, while arbitrary local and remote reads require explicit trust configuration.

Display metadata, figure and table captions, and theorem titles retain their decoded visible text. Table captions remain visible even when the table has no cell content. TeX numbering, placement, and nested formatting that a scalar caption cannot express remain diagnosed projection limits.

References to heading, theorem, table, figure and paragraph labels use PDF destinations registered by the shared Markdown renderer. A missing destination retains the visible reference text and produces an `UnresolvedInternalLink` warning.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Latex.Markdown` and `OfficeIMO.Markdown.Pdf`; those packages retain ownership of semantic projection and PDF rendering.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Latex.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
