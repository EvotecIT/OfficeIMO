# OfficeIMO.Notebook.Pdf

Export a Jupyter notebook to a PDF using the managed OfficeIMO layout engine. Code, prose, tables, equations and supported saved images stay in reading order. Conversion uses the notebook's stored data and works offline.

## Install

```sh
dotnet add package OfficeIMO.Notebook.Pdf
```

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Pdf;

using var notebook = NotebookDocument.Load("analysis.ipynb");
var report = notebook.SaveAsPdf("analysis.pdf", new NotebookPdfExportOptions {
    Title = "Lab results",
    IncludeCode = false
});
```

`ToPdf` and `ToPdfAsync` return the PDF model. `SaveAsPdf` and `SaveAsPdfAsync` write a path or caller-owned stream and report projection and layout findings. Set `FailOnLoss` to reject an artifact before it is written when a completed stage reports fidelity loss. Supply page settings and embedded fonts through `MarkdownOptions`. For portable equation layout, register a compatible OpenType math face under the family name `math` in an `OfficeRenderingProfile`, then apply the profile to `MarkdownOptions.PdfOptions` using `UseRenderingProfile`.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
