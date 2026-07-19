# OfficeIMO.Reader.OpenDocument

Native ODT, ODS, and ODP ingestion for `OfficeIMO.Reader.Core`. The adapter uses `OfficeIMO.OpenDocument` and does not invoke LibreOffice or Microsoft Office at runtime.

Configure a reader once, then reuse it:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.OpenDocument;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddOpenDocumentHandler()
    .Build();
IReadOnlyList<ReaderChunk> chunks = reader.Read("report.odt").ToList();
```

The handler emits:

- paragraph-, heading-, and table-aligned ODT chunks;
- bounded sheet/table chunks for ODS, including sheet and A1-range locations;
- slide-aligned ODP chunks with tables and optional speaker notes.

`ReaderOptions.MaxTableRows`, `MaxChars`, `ExcelHeadersInFirstRow`, `ExcelSheetName`, and `IncludePowerPointNotes` apply to the corresponding OpenDocument extraction paths. ODS extraction caps one chunk at 256 columns so repeated or adversarial ranges remain bounded.

## Dependency footprint

- **External:** None; no LibreOffice runtime.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.OpenDocument`; ODT/ODS/ODP parsing stays in the native package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
