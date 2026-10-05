# OfficeIMO JavaScript

`@evotecit/officeimo` provides strict TypeScript document libraries for evergreen browsers, Web Workers and Node 18 or newer. It writes streaming XLSX workbooks and UTF-8 CSV with no runtime dependencies. One package owns six public layers: `core`, `zip`, `xml`, `opc`, `xlsx` and `csv`.

## Install

Install the npm archive in your JavaScript application:

```sh
npm install /path/to/evotecit-officeimo-0.1.0.tgz
```

To produce the archive from this package's source directory:

```sh
npm ci
npm run build
npm pack --pack-destination /path/to/local-packages
```

`tsc` produces ES modules and declarations in `dist/`. The archive contains those files, portable bundles, this README and the MIT license. TypeScript is a development dependency; installing the archive as an application dependency installs no compiler or runtime packages.

## XLSX: workbook, worksheets and cells

```ts
import { Workbook, Cell, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";

const workbook = new Workbook({ creator: "Report", dateMode: "utc" });
const highlighted = workbook.styles.add({
  font: { bold: true, color: "123456" },
  fill: { color: "D9E1F2" },
  border: { bottom: { style: "thin", color: "123456" } },
  numberFormat: NumberFormats.Decimal
});
const worksheet = workbook.addWorksheet("Controllers", {
  columns: [
    { header: "Name", key: "name", width: 28 },
    { header: "Last seen", key: "seen", type: "date", format: "yyyy-mm-dd hh:mm" },
    { header: "Latency", key: "latency", type: "number", style: highlighted }
  ],
  freezeHeader: true,
  autoFilter: true
});
await worksheet.addRows([{ name: "DC01", seen: new Date(), latency: 12.5 }]);
await worksheet.addRows([["DC02", new Date(), new Cell(8, highlighted)]]);
saveBlob(await workbook.toBlob(), "controllers.xlsx");
```

`createWorkbook(options)` constructs the same `Workbook`; `addSheet` is the retained tabular entry point for `addWorksheet`. `workbook.worksheets` is a readonly snapshot. Each worksheet exposes its final sanitized `name` and data `rowCount`.

Rows can be synchronous or asynchronous iterables. Arrays are positional. Object rows use each column's literal `key`, or its `header` when no key is supplied; dots never traverse an object. Missing values produce empty cells and extra array values throw. Await each `addRows` call before appending to that worksheet or calling `toBlob`. A failed append prevents finalization. Repeated `toBlob` calls return the same output.

The public `StyleRegistry` owns fonts, fills, borders, number formats and cell-style indexes. `addFont`, `addFill`, `addBorder` and `addNumberFormat` register normalized definitions; `add` combines definitions or existing indexes. Equivalent definitions reuse their indexes. A column's `style` supplies a base, and `format`, `wrapText` and `alignment` override the corresponding fields. A `Cell` can select another registered style for a single value. Indexes belong to that workbook.

Cells accept strings, numbers, booleans, `Date`, `null` and `undefined`. Declared built-in types validate without coercion. Non-finite numbers and invalid dates become empty cells. Strings remain literal text, including `=...` and OOXML `_xHHHH_` sequences. XLSX never interprets a string as a formula.

Dates use Excel's 1900 system, including the fictitious leap day. `dateMode` defaults to local wall-clock fields; `"utc"` writes UTC clock fields. Cell dates carry no timezone. Property dates are UTC instants and are captured at construction. Valid cell date years are 1900 through 9999.

Excel limits are enforced: 1,048,576 rows including the header, 16,384 columns, 32,767 UTF-16 code units per cell, 64,000 styles and widths from 0 through 255 characters. Names are trimmed, illegal characters become underscores, and names are shortened to 31 UTF-16 code units without splitting a surrogate pair. Blank names become `Sheet`, reserved `History` becomes `History_`, and case-insensitive duplicates receive a suffix.

`invalidCharacterPolicy` defaults to `"strip"`. XML 1.0-invalid controls, lone surrogates, U+FFFE and U+FFFF are removed. `"reject"` throws an `OfficeIMOError` with `code: "INVALID_XML"` for names, cell/header text, property text, font names and number formats. CR, LF, tabs, emoji, Polish and right-to-left text are preserved.

## CSV

```ts
import { writeCsv, writeCsvTo, saveBlob } from "@evotecit/officeimo/csv";

const columns = [{ header: "Name" }, { header: "Healthy" }];
const rows = [["Łódź", true], ["=literal", false]] as const;
saveBlob(await writeCsv(rows, { columns, bom: true }), "status.csv");

// A caller-owned destination can avoid retaining the final CSV Blob.
await writeCsvTo(rows, { write: bytes => uploadChunk(bytes) }, { columns });
```

`uploadChunk` above stands for an application's asynchronous byte destination; its promise supplies backpressure. The library performs no network access itself. Both CSV entry points use RFC 4180 quoting, comma/semicolon/tab delimiters, CRLF by default, optional UTF-8 BOM and a final line ending. Alternative endings are LF and CR. Booleans are `True`/`False`; valid dates are UTC ISO 8601 with milliseconds. Numbers use JavaScript's invariant string representation.

Formula protection defaults to true and follows `OfficeIMO.CSV`: after leading ASCII spaces, a string beginning with `=`, `+`, `-`, `@`, tab, CR or LF gains an apostrophe. Typed negative numbers remain numeric. Set `formulaInjectionProtection: false` for trusted non-spreadsheet consumers. Widths and XLSX style/type options do not change CSV formatting.

The [shared vectors](../OfficeIMO.TestAssets/CSV/browser-exports.json) qualify byte-identical output with C#. The C# date lane uses `UseUtc = true` and `DateTimeFormat = "yyyy-MM-dd'T'HH:mm:ss.fff'Z'"`. JavaScript and .NET use their own floating-point formatting conventions; equality is qualified for the shared numeric vectors, not every possible double.

## Core: sinks, cancellation and download

```ts
import { BlobByteSink, ChunkedTextSink, detectFeatures } from "@evotecit/officeimo/core";

const bytes = new BlobByteSink();
const text = new ChunkedTextSink(bytes);
await text.write("Zażółć gęślą jaźń 🧪");
await text.close();
const blob = bytes.toBlob("text/plain;charset=utf-8");
console.log(detectFeatures().deflateRaw, blob.size);
```

A `ByteSink` implements `write(Uint8Array): void | Promise<void>`. It must accept the bytes before resolving. `BlobByteSink` copies input buffers, so callers can reuse them. `ChunkedTextSink` emits bounded UTF-8 batches and preserves surrogate pairs at append/chunk boundaries. It yields after roughly 8 ms of encoding work; a caller can still supply one large source string. `writeBytes` connects a byte iterable to a sink.

Pass an `AbortSignal` and synchronous `onProgress` to the writers. CSV/XLSX progress counts exclude headers; XLSX row events identify the sheet and completion reports the total output bytes. Throwing from a callback fails the export. Cancellation returns producer iterators and rejects pending input or sink operations. Give I/O producers and destinations the same signal so they can release their own resources.

`OfficeIMOError.code` identifies package/state/platform failures; `NotSupportedError` adds a `feature`. Input type and range errors use native `TypeError` and `RangeError`; cancellation preserves the signal's original reason. `saveBlob` requires a browser document. Node and workers should deliver their Blob through their own host mechanisms.

## ZIP

```ts
import { ZipWriter } from "@evotecit/officeimo/zip";

const zip = new ZipWriter();
await zip.add("data/status.txt", new TextEncoder().encode("Łódź 🧪"));
const blob = await zip.toBlob();
```

Pass a `ByteSink` as the constructor's first argument for direct streaming, then call `finish`. Each entry accepts bytes, a byte iterable or a callback that writes chunks to the supplied entry sink. Writes are sequential and use ZIP data descriptors. Only central-directory metadata is retained by the streaming writer; the caller owns its destination and partial-byte disposal on failure. `Crc32` exposes incremental CRC for future readers and independent integrations.

`compression: "auto"` uses platform `CompressionStream("deflate-raw")`; missing raw-deflate support falls back to stored entries. `"store"` forces that fallback. No external compressor is loaded. Paths must be relative and reject traversal, backslashes, empty segments and duplicates. ZIP64 is explicitly unsupported: entry/archive sizes must stay below 4 GiB and an archive can contain at most 65,534 entries. Reaching a ZIP64 boundary throws `code: "ZIP64_REQUIRED"`.

XLSX retains compressed worksheet chunks while appending, then packages them. Blob-returning APIs retain output proportional to file size; stored output can be substantially larger. Streaming input does not imply constant memory for a final Blob. Readers are outside the current API.

## XML

```ts
import { XmlWriter } from "@evotecit/officeimo/xml";
import { BlobByteSink } from "@evotecit/officeimo/core";

const sink = new BlobByteSink();
const xml = new XmlWriter(sink, "reject");
await xml.startElement("report", { title: "A < B" });
await xml.text("Łódź & Warsaw");
await xml.endElement();
await xml.close();
```

The writer validates schema-owned ASCII QNames, escapes every attribute/text value, preserves XML whitespace through character references and checks document completeness. Always pass user content as values or `text`, never as names. There is no raw-XML method on `XmlWriter`. `escapeXml` and `cleanXml` expose the same strip/reject policy for format implementations.

## OPC

```ts
import { OpcPackage, relationshipTypes } from "@evotecit/officeimo/opc";

const packageFile = new OpcPackage();
packageFile.addPart({ uri: "/customXml/item1.xml", contentType: "application/xml", data: '<data xmlns="urn:report">value</data>' });
packageFile.addRelationship("/", { id: "data", type: relationshipTypes.customXml, target: "/customXml/item1.xml" });
packageFile.setProperties({ creator: "Report" }, { application: "OfficeIMO", company: "Evotec" });
const blob = await packageFile.toBlob();
```

OPC owns `[Content_Types].xml`, relationship parts, core/app properties and canonical part URIs. Internal relationship targets are absolute part URIs; the package serializes their relative spelling and checks that source/target parts exist. Case-insensitive part collisions, reserved metadata paths, invalid content types and duplicate relationship IDs throw. Unicode part names use UTF-8 percent encoding; `partUri`, `relationshipPartUri` and `relativePartTarget` expose URI operations. `ContentTypes` can also be used independently.

Part sources are strings, Blobs, bytes, byte iterables or entry-sink callbacks. Some local-file WebKit worker contexts block every native Blob-reading API. A Blob part that cannot be read throws `OfficeIMOError` with `code: "PLATFORM_UNAVAILABLE"` and the native error as `cause`; no package is returned. For those workers, read input Blobs in the host and transfer byte arrays, or supply text/byte producers. Returning an output Blob from the worker remains supported because the library does not read it there. Raw XML part sources are trusted schema documents supplied by the format module or extension; use `XmlWriter` for user data. OPC does not parse or repair supplied XML.

## Extensions and current boundaries

```ts
import { Workbook } from "@evotecit/officeimo/xlsx";
import { relationshipTypes } from "@evotecit/officeimo/opc";

const workbook = new Workbook({ cellValueWriters: { milliseconds: value => Number(value) / 1000 } });
await workbook.addWorksheet("Timings", { columns: [{ header: "Seconds", type: "milliseconds", format: "0.000" }] }).addRows([["1250"]]);
workbook.addPart({
  uri: "/customXml/item1.xml", contentType: "application/xml", data: '<data xmlns="urn:report">metadata</data>',
  relationship: { id: "custom", type: relationshipTypes.customXml }
});
const blob = await workbook.toBlob();
```

Column writers synchronously convert a domain column type to a supported scalar or styled `Cell`. They receive column, row, column index and sheet-name context. They cannot inject raw cell XML. Extra parts can carry custom XML, a valid theme or other schema-owned content. An optional internal relationship defaults to `/xl/workbook.xml`; supply `source` for another existing part. Part/relationship definitions are copied when registered; producer iterables remain caller-owned until consumed.

`mergedCells`, `hyperlinks`, `conditionalFormats` and `dataValidation` are reserved worksheet options. Supplying any of them, including an empty array, throws `NotSupportedError` instead of silently ignoring it. There is no DOCX, PDF, PPTX, reader, formula engine, image placement or report-grid adapter in this package's current writer contract. The [roadmap](../Docs/ROADMAP.md#officeimo-javascript) records the consumer-driven work.

## Portable classic scripts and the .NET asset package

The archive exports `@evotecit/officeimo/classic.js`, `/xlsx.js` and `/csv.js` for copying into a portable report. Load the chosen asset with a plain script tag:

```html
<script src="officeimo.js"></script>
<script>
  const workbook = new OfficeIMO.xlsx.Workbook();
  const sink = new OfficeIMO.core.BlobByteSink();
</script>
```

The combined bundle exposes all six namespaces on global `OfficeIMO` plus the established root helpers. The XLSX bundle exposes `core`, `zip`, `xml`, `opc`, `xlsx`; CSV exposes `core`, `csv`. The standalone bundles compose in either order. Classic scripts work from `file://` and in host-owned workers without a module loader. ES modules use normal import paths on HTTP(S) and in Node; browser local-file module policy can block `file://` imports.

[OfficeIMO.Browser](../OfficeIMO.Browser/README.md) embeds the same committed TypeScript-built bundles in .NET. Its `BrowserAsset` API returns exact content and content-hashed filenames; consuming .NET applications require no Node tooling. Standalone `.mjs` assemblies preserve the asset package's module API. The npm subpaths remain genuine separate `tsc` modules with generated declarations.

## Build and validation

```sh
npm run build
npm test
npm run test:consumer
npm run test:pack -- /path/to/evidence/npm
npm run test:browser -- /path/to/evidence/browser
npm run sizes -- /path/to/evidence/sizes.json
```

`npm test` compiles strictly, checks committed bundles and runs Node's built-in runner. `npm run check` fails when a bundle differs from the compiled graph. No bundler or minifier is used. On Windows, `npm --script-shell pwsh` can select PowerShell 7 when the machine's default npm shell cannot resolve its tools.

The browser command uses the repository's existing isolated HtmlTinkerX/Playwright runner and pinned .NET SDK. It checks Chromium, Firefox and WebKit, classic `file://` scripts, ESM imports, workers, downloads and the shared corpus. [C# conformance tests](../OfficeIMO.Browser.Tests/JavaScriptConformanceTests.cs) independently open produced workbooks with OfficeIMO.Excel/OfficeIMO.Reader.Excel and validate them with the Open XML SDK. The [architecture guide](../Docs/officeimo.javascript-architecture.md) describes API stability, format ownership and adding a module.
