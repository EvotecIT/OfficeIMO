# OfficeIMO.Notebook.Markdown

Project a Jupyter notebook into a typed Markdown document. Prose, code and saved results appear in cell order. Choose whether to include code and saved outputs independently. When code is included, you can also show its execution count.

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Markdown;

using var notebook = NotebookDocument.Load("analysis.ipynb");
using var markdown = notebook.ToMarkdown(new NotebookMarkdownExportOptions { IncludeCode = false }, out var report);
string content = markdown.ToMarkdown();
```

The projection handles Markdown prose and equations, plain output and tracebacks, supported HTML tables, and bounded PNG, JPEG and SVG images from saved outputs and cell-local attachments. It uses stored representations and keeps interactive output inert. Cell and output locations identify unavailable or simplified content in the conversion report. Use `FailOnLoss` to require a projection without reported fidelity loss.

Use `ToMarkdownText` or `SaveAsMarkdownAsync` to write Markdown directly. Set
`PreserveMarkdownSource` to keep the spelling, whitespace, line endings, HTML and
links within Markdown cells. Code and saved results are still rendered in cell
order. This mode reports omitted cell-local attachment data. Use the default
normalized export when you need attachments embedded in the result.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Notebook.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
