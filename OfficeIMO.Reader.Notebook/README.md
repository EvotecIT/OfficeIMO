# OfficeIMO.Reader.Notebook

`OfficeIMO.Reader.Notebook` adds bounded Jupyter `.ipynb` ingestion to an isolated `OfficeDocumentReader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Notebook;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddNotebookHandler()
    .Build();

string markdown = reader.ConvertToMarkdown("analysis.ipynb");
```

Markdown, raw, and code cells are projected in source order. Code is emitted as fenced Markdown using the notebook language when available. Stream, error, `text/markdown`, and `text/plain` outputs can be included; binary display outputs are deliberately not decoded or executed.

The adapter uses the runtime JSON API already present in the Reader graph. It does not run kernels, execute cells, start processes, load models, or contact notebook services. Input, cell, output-count, and output-character limits are explicit through `ReaderOptions` and `ReaderNotebookOptions`.
