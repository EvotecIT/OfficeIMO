# OfficeIMO.Reader.OpenDocument

Native ODT, ODS, and ODP ingestion for `OfficeIMO.Reader`. The adapter uses `OfficeIMO.OpenDocument` and does not invoke LibreOffice or Microsoft Office at runtime.

Register the handler once, then use the normal reader API:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.OpenDocument;

DocumentReaderOpenDocumentRegistrationExtensions.RegisterOpenDocumentHandler();
IReadOnlyList<ReaderChunk> chunks = DocumentReader.Read("report.odt").ToList();
```

ODT extraction emits paragraph-, heading-, and table-aligned chunks. ODS and ODP extraction use the same handler and are expanded by their native format milestones.
