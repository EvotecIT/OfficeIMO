# OfficeIMO.Reader.Markdown

Markdown support for `OfficeIMO.Reader.Core`, backed by the typed `OfficeIMO.Markdown` parser.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddMarkdownHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("README.md");
```

The adapter preserves source spans and heading paths and projects Markdown tables and supported visual fences into Reader's structured models.
