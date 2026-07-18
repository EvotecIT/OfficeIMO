# OfficeIMO.Epub.Html

`OfficeIMO.Epub.Html` is the thin visual adapter between the bounded EPUB package model and the existing OfficeIMO HTML/Drawing renderer. It adds no image encoder or layout engine.

Load raw chapter HTML and resource payloads when visual fidelity matters:

```csharp
EpubDocument book = EpubDocument.Load(path, new EpubReadOptions {
    IncludeRawHtml = true,
    IncludeResourceData = true
});

IReadOnlyList<OfficeImageExportResult> pages = await book
    .ToImages()
    .Paged()
    .AsPng()
    .ExportAsync();
```

When raw HTML or resource bytes were not retained, export remains available through a diagnosed plain-text or missing-resource fallback.
