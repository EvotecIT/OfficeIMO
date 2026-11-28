# OfficeIMO.Reader.DocumentExtraction

`OfficeIMO.Reader.DocumentExtraction` plugs the OfficeIMO reader into Microsoft's preview
[`Microsoft.Extensions.DocumentExtraction`](https://devblogs.microsoft.com/dotnet/document-extraction-dotnet-preview/)
abstraction. Word, Excel, PowerPoint, PDF, HTML, EPUB, email, OpenDocument and the other formats OfficeIMO reads are
extracted in process, with no Office installation, service or model call, into the shared `ExtractionDocument` model.

The Microsoft packages are previews and their APIs are marked experimental (`EXTEXP0019`, `MEDE0001`), so code that
uses `ExtractionDocument` needs to suppress those diagnostics. This package targets `net8.0` and `net10.0`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.DocumentExtraction
dotnet add package OfficeIMO.Reader.All
```

`OfficeIMO.Reader.All` registers every OfficeIMO format. To keep the dependency graph small, reference only the
`OfficeIMO.Reader.{Format}` packages you need and register those handlers instead.

## Extract a document

```csharp
using Microsoft.Extensions.Documents;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.DocumentExtraction;

using var client = new OfficeIMODocumentExtractionClient(builder => builder.AddAllOfficeIMOHandlers());

await using FileStream stream = File.OpenRead("policy.docx");
ExtractionDocument document = await client.ExtractAsync(
    stream,
    "policy.docx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

foreach (ExtractionDocumentElement element in document.EnumerateContent()) {
    // Text, table and image elements in document order.
}
```

The media type selects the format when the identifier has no matching extension, so uploads named `upload-42` work.
An empty media type leaves the choice to the identifier's extension and content detection. The client never disposes
the stream.

If you already have an `OfficeDocumentReadResult`, `result.ToExtractionDocument("id")` does the same mapping without
reading again.

## Register with dependency injection

```csharp
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DocumentExtraction;
using OfficeIMO.Reader.All;

builder.Services
    .AddOfficeIMODocumentExtraction(reader => reader.AddAllOfficeIMOHandlers())
    .UseLogging();

// Next to other providers, under a key:
builder.Services.AddKeyedOfficeIMODocumentExtraction("office", reader => reader.AddAllOfficeIMOHandlers());
```

Both return the abstraction's `DocumentExtractionClientBuilder`, so `DelegatingDocumentExtractionClient` middleware and
`UseLogging()` compose as with any other provider. Without a reader callback, an `OfficeDocumentReader` registered in
the container is used. `GetService(typeof(OfficeDocumentReader))` returns the reader behind the client.

## What the document contains

| OfficeIMO data | `ExtractionDocument` |
| --- | --- |
| Headings | Nested `ExtractionDocumentSection`s; the heading is the first element, with role `heading` and `HeadingLevel` |
| Content before the first heading | A root section without a heading |
| Paragraphs | `ExtractionDocumentTextElement` without a role |
| List items | Role `listItem`; level, marker and index in additional properties |
| Page headers and footers | Roles `pageHeader` and `pageFooter` |
| Footnotes, comments, speaker notes, captions, code and other block kinds | The OfficeIMO kind as the role, for example `footnote` or `speaker-notes` |
| Tables | `ExtractionDocumentTableElement`; a source header row becomes cells with role `columnHeader` |
| Images | `ExtractionDocumentImageElement` with the bytes as `DataContent`, alternative text, and OCR text when an OCR processor ran |
| Pages, slides, sheets | `SourcePages` (`page-N`, `slide-N`, `sheet-Name`), referenced from each element's `Sources` |
| PDF geometry | Page dimensions in points with a bottom-left origin, and polygons for regions with an area |
| Title, author, subject, keywords and format metadata | `ExtractionDocument.Metadata` |
| The full OfficeIMO result | `ExtractionDocument.RawRepresentation`; each payload's `RawRepresentation` holds its block, table or asset |

Data the shared model has no slot for goes into `AIContent.AdditionalProperties` on the element payload, or into
`ExtractionDocument.Metadata` for document-wide lists. The keys are constants on `OfficeIMOExtractionPropertyNames`:

| Key | Where | Value |
| --- | --- | --- |
| `officeimo.blockKind`, `officeimo.blockId`, `officeimo.blockAnchor` | Text payloads | OfficeIMO block kind, identifier and source anchor |
| `officeimo.headingPath` | Any payload | Heading path, such as `Intro > Scope` |
| `officeimo.listLevel`, `officeimo.listMarker`, `officeimo.listIndex` | List items | Nesting depth, displayed marker, ordered index |
| `officeimo.sheet`, `officeimo.a1Range` | Spreadsheet content | Worksheet name and cell range |
| `officeimo.slide` | Presentation content | Slide number |
| `officeimo.tableTitle`, `officeimo.tableKind`, `officeimo.tableTruncated`, `officeimo.tableTotalRowCount` | Table payloads | Table details and truncation |
| `officeimo.assetId`, `officeimo.assetTitle`, `officeimo.assetFileName`, `officeimo.imageWidth`, `officeimo.imageHeight` | Image payloads | Asset details |
| `officeimo.ocrProvider`, `officeimo.ocrConfidence` | Image payloads | OCR provenance |
| `officeimo.links`, `officeimo.formFields`, `officeimo.diagnostics` | Document metadata | Lists of string-keyed dictionaries |
| `officeimo.inputKind`, `officeimo.sourceHash` | Document metadata | Detected format and SHA-256 of the source |

## Options

`OfficeIMODocumentExtractionOptions` sets client defaults: `ReaderOptions` (limits, passwords, hashing),
`IncludeImagePayloads` and `IncludeRegions`. Per call, pass reader options through
`DocumentExtractionOptions.AdditionalProperties["officeimo.readerOptions"]`. `ExtractorName` and any other additional
property are rejected with `NotSupportedException`, as the abstraction requires for unsupported settings.

## Limits

- Table cells are always 1×1. The OfficeIMO table model does not record merged cells yet, so row and column spans are
  not reported.
- Readers name columns `Column 1`, `Column 2`... when the source has no header row; those tables get no header cells.
- PDF text lines are reported as baselines, so text elements reference their page without a polygon. Tables and images
  carry polygons.
- Word images have no position in the read result, so they are placed in a trailing section.
