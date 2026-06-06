# OfficeIMO.Reader.Pdf

Thin PDF adapter for `OfficeIMO.Reader`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;

DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

IReadOnlyList<ReaderChunk> chunks = DocumentReader
    .Read("invoice.pdf")
    .ToList();
```

The adapter uses `OfficeIMO.Pdf`'s logical read model and emits page-aware chunks with `ReaderLocation.Page`, Markdown text, detected tables, image placeholders, link annotations, and form widget summaries when available.
