# OfficeIMO.Latex.Pdf

`OfficeIMO.Latex.Pdf` is the direct, loss-aware adapter for OfficeIMO's bounded LaTeX profile. It does not execute TeX or introduce another layout engine: `OfficeIMO.Latex.Markdown` owns semantic projection, and `OfficeIMO.Markdown.Pdf` plus the shared PDF engine own rendering.

```csharp
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Pdf;

LatexDocument document = LatexDocument.Load("article.tex").Document;
PdfSaveResult result = document.SaveAsPdf("article.pdf");

result.Report.RequireNoLoss(); // optional strict conversion gate
result.Pipeline.RequireSuccess(); // exact output pipeline gate
```

`PdfSaveResult` combines native parser, bounded-profile projection, PDF layout/resource/font diagnostics, and exact output-pipeline evidence. Use `ToPdfDocumentResult(...)` when conversion and post-processing should happen before save. TeX macros and package behavior are not executed; unsupported or simplified constructs are preserved visibly when configured and remain explicit warnings.

The zero-options resource policy is inherited from `MarkdownToPdfOptions`: system fonts and bounded in-source resources are allowed, while arbitrary local and remote reads require explicit trust configuration.

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
