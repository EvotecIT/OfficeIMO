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
- page-aligned ODG/FODG chunks with shape text, including nested groups in paint order. Drawing warnings identify hidden-layer text inclusion and the bounds of the chunk profile.

`ReaderOptions.MaxTableRows` bounds table extraction, and `MaxChars` splits oversized paragraph, table, sheet, and slide text using the shared Reader chunker. The Reader normalizes `MaxChars` to at least 256. Split segments preserve source locations and all extracted text; structured table metadata appears on the first segment and remains bounded separately by the row and column limits. Markdown tables may span segments.

Pass `ReaderOpenDocumentOptions` to `AddOpenDocumentHandler(...)` to select `SheetName`, `A1Range`, `HeadersInFirstRow`, and `IncludeSpeakerNotes`. ODS extraction caps a sheet at 256 columns. Set `ReaderOptions.OpenPassword` to read supported encrypted ODF packages.

`ReaderOpenDocumentOptions.MaxExtractedCharacters` defaults to 16,000,000 characters across one document, including every logical occurrence of repeated cells. Extraction rejects inputs exceeding that budget before building joined text and Markdown. Raise it explicitly for trusted larger workloads; `MaxXmlCharacters` separately bounds the parsed XML. Integer native admission and extraction failures throw `ReaderResourceLimitException` with the effective maximum. Native compression-ratio and encrypted-package resource failures throw `ReaderAdmissionLimitException`, retaining the structured native cause. Ratio measurements remain exact doubles; encrypted failures leave numeric fields null when the native owner does not expose them. Both exception types remain terminal during lazy enumeration, folder, batch and nested ZIP reads. Malformed files keep their ordinary per-file failure behavior.

The platform XML-character ceiling remains a malformed-XML boundary. For a separate native load shared with Reader projection, an ingestion profile can keep XML entry and decrypted-entry byte limits at or below `MaxXmlCharacters`; flat XML additionally requires the input-byte limit at or below that character limit. DTDs are prohibited. Reader's ordinary character policy is not raised automatically.

## Text and metadata projection

```csharp
OfficeDocumentReadResult document = reader.ReadDocument("report.odt", new ReaderOptions {
    Projection = ReaderProjection.TextAndMetadata
});
```

This projection emits located native text and metadata while omitting secondary Markdown, image/object assets, visual and form projections. It uses the native deferred binary profile and reports binary entry paths, declared lengths and payload state with `not-requested` coverage. `SupportsTextAndMetadata` in the capability manifest describes this qualified profile. Native input, ZIP expansion and XML limits still apply; Reader output limits are separate.

Encrypted ODF containing non-XML binary entries is refused with the native `UnsupportedProfile` classification before key derivation. Encrypted XML/RDF-only input retains ordinary decryption checks. FODG with inline `office:binary-data` is refused before Base64 decoding. The adapter does not fall back to `Full` when this projection is requested. Use the native document's bounded embedded-payload session for explicit selected reads. The default `Full` projection retains rich extraction.

For one native load shared by Reader projection and selected payload reads, call the loaded model extension:

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.Reader.OpenDocument;

OdtDocument native = OdtDocument.Load("report.odt", new OdfLoadOptions { DeferBinaryPayloads = true });
OfficeDocumentReadResult projected = native.ReadDocument("report.odt", new ReaderOptions {
    Projection = ReaderProjection.TextAndMetadata
});
using var payloads = native.InspectEmbeddedPayloads();
```

The caller retains the native model. This entrypoint applies Reader projection and output budgets without reopening the source; its logical source name is not read or hashed as a file. Input and native loading limits apply at the preceding load, and already expanded payloads cannot be deferred retroactively.

## Rich document inspection

`ReadDocument` also retains native links and bounded image payloads, ODT header/footer/note stories, ODP slide notes and table links, and authored ODP/ODG masters, including unused masters. Source locations distinguish master stories from slide/page content; ODT flow content does not acquire inferred physical pages. Reading external links and image references never fetches their targets. Image placements share already-loaded package payloads by canonical entry path.

Enable `ReaderOptions.Inspection` to collect native revision declarations, sparse ODS formula/cache and annotation prototypes, authored visibility, package feature observations, and bounded safety evidence. Repeated ODS ranges keep original repeat counts and source ranges while the selected A1 range identifies the inspected subset. The result records unsupported form-widget interpretation explicitly because the current owner preserves form XML without a typed widget surface. Source transforms, image OCR, and embedded object execution are not performed.

`ReaderOpenDocumentOptions` bounds rich blocks, link counts/characters, individual image bytes, and total unique image bytes. An exceeded rich resource or native safety ceiling is terminal. Table row/column omissions remain explicit in table metadata and diagnostics. The native `HeaderFooters`, master `Shapes`, and their paragraph/table/inline/list editing wrappers persist changes to their owning package part.

## Dependency footprint

- **External:** None; no LibreOffice runtime.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.OpenDocument`; ODT/ODS/ODP/ODG parsing stays in the native package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
