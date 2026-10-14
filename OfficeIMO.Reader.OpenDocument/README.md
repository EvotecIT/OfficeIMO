# OfficeIMO.Reader.OpenDocument

Native ODT, ODS, ODP, ODG, and FODG ingestion for `OfficeIMO.Reader.Core`. The adapter uses `OfficeIMO.OpenDocument` and does not invoke LibreOffice or Microsoft Office at runtime.

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
- slide-aligned ODP chunks with tables and optional speaker notes;
- page-aligned ODG/FODG chunks with shape text, including nested groups in paint order. Drawing warnings identify hidden-layer text inclusion and the absence of master text, OCR, and embedded-object extraction.

`ReaderOptions.MaxTableRows` bounds table extraction, and `MaxChars` splits oversized paragraph, table, sheet, and slide text using the shared Reader chunker. The Reader normalizes `MaxChars` to at least 256. Split segments preserve source locations and all extracted text; structured table metadata appears on the first segment and remains bounded separately by the row and column limits. Markdown tables may span segments.

Pass `ReaderOpenDocumentOptions` to `AddOpenDocumentHandler(...)` to select `SheetName`, `A1Range`, `HeadersInFirstRow`, and `IncludeSpeakerNotes`. ODS extraction caps a sheet at 256 columns. Set `ReaderOptions.OpenPassword` to read supported encrypted ODF packages.

`ReaderOpenDocumentOptions.MaxExtractedCharacters` defaults to 16,000,000 characters across one document, including every logical occurrence of repeated cells. Extraction rejects inputs exceeding that budget before building joined text and Markdown. Raise it explicitly for trusted larger workloads; `MaxXmlCharacters` separately bounds the parsed XML.

## Dependency footprint

- **External:** None; no LibreOffice runtime.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.OpenDocument`; ODT/ODS/ODP/ODG parsing stays in the native package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
