# OfficeIMO.Reader.OpenDocument

Native ODT, ODS, and ODP ingestion for `OfficeIMO.Reader`. The adapter uses `OfficeIMO.OpenDocument` and does not invoke LibreOffice or Microsoft Office at runtime.

Register the handler once, then use the normal reader API:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.OpenDocument;

DocumentReaderOpenDocumentRegistrationExtensions.RegisterOpenDocumentHandler();
IReadOnlyList<ReaderChunk> chunks = DocumentReader.Read("report.odt").ToList();
```

The handler emits:

- paragraph-, heading-, and table-aligned ODT chunks;
- bounded sheet/table chunks for ODS, including sheet and A1-range locations;
- slide-aligned ODP chunks with tables and optional speaker notes.

`ReaderOptions.MaxTableRows`, `MaxChars`, `ExcelHeadersInFirstRow`, `ExcelSheetName`, and `IncludePowerPointNotes` apply to the corresponding OpenDocument extraction paths. ODS extraction caps one chunk at 256 columns so repeated or adversarial ranges remain bounded.
