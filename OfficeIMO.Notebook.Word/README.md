# OfficeIMO.Notebook.Word

Create editable DOCX reports from saved Jupyter notebook content. Prose, code and saved text are editable, and supported tables become native Word tables.

## Install

```sh
dotnet add package OfficeIMO.Notebook.Word
```

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Word;

using var notebook = NotebookDocument.Load("analysis.ipynb");
var report = notebook.SaveAsWord("analysis.docx", new NotebookWordExportOptions {
    IncludeCode = false,
    Title = "Analysis"
});
```

The conversion runs locally and reads the notebook's existing outputs. It does not run the notebook or fetch external images. The report carries diagnostics from the notebook projection and the Word renderer. Set `FailOnLoss` to reject a conversion that would omit or approximate source content.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
