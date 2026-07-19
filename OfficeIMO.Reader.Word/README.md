# OfficeIMO.Reader.Word

Word ingestion for `OfficeIMO.Reader.Core` with DOCX, DOCM, and legacy DOC support.

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .Build();
```
