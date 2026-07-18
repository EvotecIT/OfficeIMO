# OfficeIMO.Latex.Pdf

`OfficeIMO.Latex.Pdf` is the direct, loss-aware adapter for OfficeIMO's bounded LaTeX profile. It does not execute TeX or introduce another layout engine: `OfficeIMO.Latex.Markdown` owns semantic projection, and `OfficeIMO.Markdown.Pdf` plus the shared PDF engine own rendering.

```csharp
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Pdf;

LatexDocument document = LatexDocument.Load("article.tex").Document;
PdfDocumentConversionResult result = document.SaveAsPdf("article.pdf");

result.RequireNoLoss(); // optional strict gate
```

`PdfDocumentConversionResult` combines native parser diagnostics, bounded-profile projection diagnostics, and PDF layout/resource/font diagnostics. TeX macros and package behavior are not executed; unsupported or simplified constructs are preserved visibly when configured and remain explicit warnings.

The zero-options resource policy is inherited from `MarkdownPdfSaveOptions`: system fonts and bounded in-source resources are allowed, while arbitrary local and remote reads require explicit trust configuration.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Latex.Markdown` and `OfficeIMO.Markdown.Pdf`; those packages retain ownership of semantic projection and PDF rendering.
