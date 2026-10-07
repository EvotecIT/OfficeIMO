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

Excel limits are enforced: 1,048,576 rows including headings and the footer, 16,384 columns, 32,767 UTF-16 code units per cell, 64,000 styles and widths from 0 through 255 characters. Names are trimmed, illegal characters become underscores, and names are shortened to 31 UTF-16 code units without splitting a surrogate pair. Blank names become `Sheet`, reserved `History` becomes `History_`, and case-insensitive duplicates receive a suffix.

`invalidCharacterPolicy` defaults to `"strip"`. XML 1.0-invalid controls, lone surrogates, U+FFFE and U+FFFF are removed. `"reject"` throws an `OfficeIMOError` with `code: "INVALID_XML"` for names, cell/header text, property text, font names and number formats. CR, LF, tabs, emoji, Polish and right-to-left text are preserved.

## Streamed output and resource limits

Select a caller-owned sink when constructing the workbook to deliver ZIP bytes during row production. Complete streamed worksheets in order: `close()` completes a worksheet explicitly, and starting the next worksheet closes its predecessor. Await each append before switching worksheets. A closed worksheet cannot receive more rows. `finish()` completes the ZIP directory and returns data-row, sheet and byte counts; repeated calls reuse the same result. The row count excludes headings, footers and preservation-sheet records. The sheet count includes the preservation sheet when present.

```ts
import { createWorkbook } from "@evotecit/officeimo/xlsx";
import type { Rows } from "@evotecit/officeimo/core";

async function exportRows(rows: Rows, destination: WritableStream<Uint8Array>) {
  const writer = destination.getWriter();
  try {
    const book = createWorkbook({
      sink: { write: bytes => writer.write(bytes) }, dateMode: "utc",
      limits: { maxRows: 1_000_000, maxCells: 8_000_000, maxOutputBytes: 512_000_000 }
    });
    const sheet = book.addSheet("Data", {
      columns: [{ header: "Name", key: "name" }, { header: "Amount", key: "amount", type: "number", format: "0.00" }]
    });
    await sheet.addRows(rows);
    const result = await book.finish();
    await writer.close();
    return result;
  } catch (error) {
    await writer.abort(error).catch(() => {});
    throw error;
  } finally { writer.releaseLock(); }
}
```

The library awaits byte acceptance and never closes a supplied sink. The destination owns cancellation and disposal of partial output. Use the same signal for the writer, paged source and destination I/O. `toBlob()` remains the default output path and supports interleaved appends to different worksheets; it retains compressed output proportional to file size. A workbook with a caller-owned sink uses `finish()` instead of `toBlob()`.

Both CSV and XLSX accept optional `limits`: `maxRows` counts data rows per worksheet/export, including generated preservation rows; `maxCells` counts emitted cells, including titles, headings, footers and preservation records. `maxTextCharacters` counts UTF-16 units of string values, including CSV formatter/display/null-text results and XLSX titles, headings, footers, previews and full preservation records. Numeric/date encodings, CSV quoting and XML markup do not count toward that text limit. `maxOutputBytes` limits actual UTF-8 CSV or ZIP bytes accepted by the sink and also bounds compressed worksheet retention in Blob mode. XLSX also accepts `maxStyles`, `maxSheets`, `maxHyperlinks`, `maxImageBytes` and `maxMergedRanges`. The merge limit defaults to 10,000 across all worksheets, including generated title/group merges. Cell/text budgets are checked at row ingress, reserving generated cells/text before retaining input. Exceeding a ceiling throws `OfficeIMOError` with `code: "RESOURCE_LIMIT"`. Limits are nonnegative safe integers; `maxStyles` is from 1 through 64,000. No host-dependent timing or heap threshold is enforced by the library.

XLSX rejects oversized text by default. Set `oversizedText: "preserve"` to write a bounded preview with an internal link to the full value on a `Text overflow` worksheet. Each overflow record identifies the source sheet/cell, one-based part number, text and total part count. Concatenate its `Text` values in part order to reconstruct the complete XML-valid value. Chunks never split a supplementary Unicode character. The same strip/reject policy applies to invalid XML characters. Preserve mode does not replace an existing external link on that cell; that conflict throws.

ZIP entries cannot interleave, so preservation retains a bounded text spool until the report worksheets finish. `maxOverflowCharacters` defaults to 4,000,000 UTF-16 units. Width sampling has separate defaults of 100,000 retained cells and 1,000,000 UTF-16 units of encoded row XML, including markup, configurable through `maxBufferedCells` and `maxBufferedCharacters`. Exceeding either budget fails visibly; the writer does not silently truncate full values. CSV retains long text directly and is an alternative when Excel's cell/storage constraints do not suit the data.

## Resolved values, grouped headings, totals and print layout

`ExportCell` captures a typed value, optional display text and portable presentation before an export begins. A report producer can resolve highlighting once and reuse the same cells for XLSX and CSV. XLSX uses the typed value; CSV uses it by default and uses supplied display text with `valueMode: "display"`. Formula protection and quoting still run after display-text selection and CSV formatting.

```ts
import { createWorkbook, ExportCell } from "@evotecit/officeimo/xlsx";
import { writeCsv } from "@evotecit/officeimo/csv";

const columns = [
  { header: "Name", key: "name", groups: ["Identity"], width: 28 },
  { header: "Latency", key: "latency", groups: ["Metrics"], type: "number", format: "0.000" }
] as const;
const rows = [{ name: "Łódź", latency: new ExportCell(125.75, {
  text: "125.750 ms", presentation: { background: "FCE4D6", bold: true }
}) }];
const book = createWorkbook();
const sheet = book.addSheet("Report", {
  title: { text: "Controller health", style: { font: { size: 20 }, fill: { color: "D9E1F2" } }, height: 32 },
  columns, table: { name: "ReportData" }, freezeHeader: true,
  autoSize: { sampleRows: 100, minWidth: 8, maxWidth: 40 },
  footer: { values: ["Totals"], totals: { latency: "average" }, style: { font: { bold: true } } },
  print: { paper: "A4", orientation: "landscape", repeatHeaders: true, header: "Health report", footer: "Measured latency" }
});
await sheet.addRows(rows);
const excel = await book.toBlob();
const csv = await writeCsv(rows, { columns, valueMode: "display" });
```

Portable presentation supports background/text color, bold, italic, wrapping, alignment and number format. It overlays the resolved row/cell presentation while preserving unspecified fields; an explicit workbook-local `Cell.style` retains precedence. CSV carries scalar values/display text; it has no cell-style format.

`title` adds one merged row above the headings. Its default font is bold, 18 points, with a 28-point row height; `style` composes a normal `CellStyle` patch and `height` overrides the row height. Titles use literal XML-valid text of at most 32,767 UTF-16 units. Contiguous column `groups` with matching ancestor labels form merged heading spans above the leaf headers, with at most 16 levels. Native tables start at the leaf-header row, and `freezeHeader` freezes the title and all heading rows. Style callback row numbers, explicit hyperlink/image anchors and merge ranges use the final worksheet coordinates, including the title.

`mergedCells: ["A1:B2", "A3:B3"]` merges explicit uppercase A1 ranges. Each range must contain at least two cells, stay within the declared columns and rows actually exported, and avoid other merges, including generated title/group merges. A native table cannot contain merged cells. The top-left cell retains its value and presentation; covered cells must contain null/undefined or an empty string, and cannot carry hyperlinks or computed totals. A nonempty covered value throws instead of being silently lost. Use `includeHeader: false` for a free-form region without leaf headings:

```ts
const regions = book.addSheet("Summary", {
  columns: [{ header: "A" }, { header: "B" }], includeHeader: false,
  mergedCells: ["A1:B2", "A3:B3"]
});
await regions.addRows([["Report summary", null], [null, null], ["End", null]]);
```

`footer.values` supplies explicit typed footer cells in column order. `footer.totals` maps unambiguous column keys (or headers without keys) to `sum`, `count`, `average`, `min` or `max`. Aggregation uses the finite numbers written to Excel, including converted date serials and column-writer results, and retains only the state needed by each column's operation. Count means numeric count and uses General format rather than a column's date format. Formulas use `SUBTOTAL`, with cached values for readers; averages/minima/maxima remain blank when no visible numeric value exists, including when filtering hides every numeric row. Their guarded expressions are registered as custom table totals so native Excel sees consistent formula metadata. Totals respond to Excel filtering; strings supplied as ordinary values never become formulas. A sum outside JavaScript's finite numeric range fails visibly.

`autoSize` inspects at most `sampleRows` leading rows, default 100 and maximum 10,000, then starts the worksheet. Sampling can span append calls. Each sampled row is validated and serialized once within the explicit buffer limits; custom value/style callbacks do not run again when the sample is written. Widths use an approximate character count, preferring `ExportCell.text`, clamped to `minWidth`/`maxWidth`; column `width` wins. This is bounded sizing rather than font measurement.

`print` sets A4/Letter paper, portrait/landscape orientation, fit-to-page dimensions and inch-based margins. Defaults are A4, landscape, one page wide and unlimited pages high. Print area includes the title, headings, data and footer; `repeatHeaders` repeats grouped/leaf headings and leaves the report title on its first page. Center header/footer strings are literal text: ampersands are protected from Excel control-code interpretation.

## Report tables, highlighting and images

Use worksheet presentation options to carry a report's highlights into Excel while retaining typed values and column number formats:

```ts
import { Workbook, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";

const book = new Workbook({ creator: "Health report", dateMode: "utc" });
const headerStyle = book.styles.add({
  font: { bold: true, color: "FFFFFF" }, fill: { color: "203864" },
  verticalAlignment: "center", wrapText: true
});
const sheet = book.addWorksheet("Controllers", {
  columns: [
    { header: "Name", key: "name", width: 28 },
    { header: "Latency", key: "latency", type: "number", format: "0.000" },
    { header: "Seen", key: "seen", type: "date", format: NumberFormats.DateTime },
    { header: "Healthy", key: "healthy", type: "boolean" }
  ],
  table: { name: "Controllers", style: "TableStyleMedium9" },
  headerStyle, headerHeight: 30, rowHeight: 24,
  freezeHeader: true, freezeColumns: 1,
  alternatingRowStyle: { fill: { color: "EAF1F8" } },
  rowStyle: ({ values }) => values[3] === false
    ? { fill: { color: "FCE4D6" }, font: { bold: true } } : undefined,
  cellStyle: ({ value, columnIndex }) => columnIndex === 2 && typeof value === "number" && value > 100
    ? { font: { color: "C00000" } } : undefined
});
await sheet.addRows([
  { name: "Łódź", latency: 125.75, seen: new Date("2026-10-06T12:00:00Z"), healthy: false }
]);
sheet.addHyperlink({ cell: "A2", target: "https://example.com/report", tooltip: "Open report" });
saveBlob(await book.toBlob(), "health-report.xlsx");
```

`table` writes a native Excel table with filtering and a built-in `TableStyleLight1..21`, `TableStyleMedium1..28` or `TableStyleDark1..11` style. Row banding defaults to true; `bandedColumns`, `firstColumn` and `lastColumn` control other table accents. Names are workbook-unique ASCII identifiers, at most 255 characters, and cannot be cell references. Omitted names are generated. Table headers must be nonblank, unique ignoring case and at most 255 characters. An empty export writes the headers without creating a table or an artificial data row.

Styles apply in this order: column style/format, alternating-row patch, row patch, then cell patch. A `Cell` with an explicit style replaces the column/alternating/row presentation; the cell callback can still overlay it. `StyleRegistry.compose(base, patch)` exposes the same composition. Font fields and border edges merge; numeric font/fill/border indexes replace their component. Unspecified number formats, wrapping and alignment remain intact. Fonts support bold, italic, underline and strike; `verticalAlignment` complements horizontal `alignment`.

Style callbacks are synchronous. Their `row` is the one-based worksheet row, including its header; `columnIndex` is one-based. Row `values` follow the declared export order and unwrap `Cell` values. A cell callback sees the value after a custom column writer. Headers use `headerStyle` separately. Alternating patches start on the second data row and remain consistent across appends. Heights are points, positive and at most 409; `freezeColumns` freezes leading columns alongside an optional header.

Native tables default unspecified column widths to 20 characters so date/time and numeric columns have useful space. Set `defaultColumnWidth` for another fallback or a column's `width` for an individual override. Worksheets without a table retain Excel's default width unless a width is supplied. These are fixed widths, not font-measured autofit.

Hyperlinks can also be supplied through `SheetOptions.hyperlinks`. Each link addresses one exported cell and uses an absolute HTTP, HTTPS or mailto URL. HTTP credentials, unsafe schemes and duplicate link cells are rejected. Links preserve the existing cell value and style; use a style patch for blue/underlined link text when desired.

`sheet.addImage({ data: pngBytes, row: 6, column: 1, width: 640, height: 320, description: "Latency chart" })` places a PNG, such as a rendered chart, without changing table data. Supply a `Uint8Array` containing the complete PNG; the library checks its signature/IHDR and copies the bytes. Anchors are one-based; display dimensions are CSS pixels at 96 DPI. The writer embeds the image and a one-cell drawing anchor. It does not render charts or convert SVG. Multiple worksheets can each contain a table, links and images.

## Live Excel conditional formatting

`conditionalFormats` writes native Excel rules that recalculate when users edit workbook values. Use these alongside resolved row/cell highlights when the report needs both an exported appearance and live spreadsheet behavior:

```ts
import { Workbook } from "@evotecit/officeimo/xlsx";

const book = new Workbook({ limits: { maxConditionalFormats: 100, maxDifferentialStyles: 50 } });
const sheet = book.addSheet("Latency", {
  columns: [
    { header: "Name", key: "name", width: 24 },
    { header: "Latency", key: "latency", type: "number", format: "0.000", width: 18 }
  ],
  table: { name: "LatencyReport" },
  conditionalFormats: [
    { type: "cellIs", range: { column: "latency" }, operator: "greaterThan", value: 100,
      style: { fill: { color: "FFC7CE" }, font: { color: "9C0006" } }, stopIfTrue: true },
    { type: "expression", range: { column: "name", through: "latency" }, formula: "$B2>100",
      style: { font: { bold: true } } },
    { type: "dataBar", range: { column: "latency" }, color: "638EC6" }
  ]
});
await sheet.addRows([{ name: "Łódź", latency: 125.75 }, { name: "Warsaw", latency: 25 }]);
const blob = await book.toBlob();
```

A column target uses an unambiguous declared key or a one-based column number. Optional `through` extends it across adjacent columns. The writer resolves the final data range after all appends, excluding report titles, grouped/leaf headings and footer totals. An empty data export emits no column-targeted rule. An explicit uppercase A1 cell or rectangle, such as `"B2:B20"`, can include headings or totals but must stay within the declared columns and exported rows; row bounds are checked when the worksheet closes. Overlapping rules are supported.

| Rule type | Contract |
| --- | --- |
| `cellIs` | `equal`, `notEqual`, `greaterThan`, `greaterThanOrEqual`, `lessThan` or `lessThanOrEqual` with one finite numeric `value`; `between` or `notBetween` with an ordered pair of numeric `values`. Dates can be compared using Excel serial numbers or an expression. |
| `expression` | An Excel formula and a differential `style`. Relative references are anchored to the top-left cell of the target range; use the first data row appropriate to your title/heading layout. |
| `colorScale` | Two or three `stops`, each with a hex `color` and `threshold`. |
| `dataBar` | A hex `color`, optional `minimum`/`maximum` thresholds and `showValue`, default true. Writes standard gradient data bars. |

Thresholds use `{ type: "min" }` at the start, `{ type: "max" }` at the end, `{ type: "number", value: 100 }`, percent/percentile values from 0 through 100, or `{ type: "formula", value: "MAX($B$2:$B$20)" }`. Color scales require two or three stops; a three-color scale can use a 50th-percentile middle stop. Numeric thresholds of the same type must be ordered. Data bars default to the range minimum and maximum. Icon sets and Office extension features such as solid bars, negative-bar colors and axes are outside this writer contract.

Array order sets worksheet-wide rule priority, starting at one. `stopIfTrue` applies to comparison and expression rules. A differential style changes only supplied font flags/color, solid fill, individual border edges or a number-format string. Explicit `false` clears a font flag. Omitting the number format preserves the cell's existing numeric/date formatting; font family/size, wrapping, alignment and numeric style-component indexes are unsupported in differential styles and fail explicitly.

Rules are captured at worksheet creation and retain metadata rather than source rows or per-cell style assignments. Identical differential styles share one workbook definition. `maxConditionalFormats` and `maxDifferentialStyles` each default to 1,000 across the workbook, independently of `maxStyles`; zero disables the corresponding feature. Even a column rule on an empty worksheet counts toward the rule limit. Formula text uses invariant Excel syntax, accepts an optional leading `=`, and is limited to 8,192 characters. The writer stores formulas without evaluating them; XML-invalid formula characters are rejected under either text policy so cleanup cannot change their meaning. CSV does not carry conditional rules.

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

CSV carries values rather than colors or fonts. Use a column's synchronous `valueFormatter` to export status labels, chosen date formats or other scalar representations. Formatting runs before formula protection and quoting, so return plain values rather than escaped CSV:

```ts
const blob = await writeCsv([{ name: "Łódź", healthy: false }], {
  columns: [
    { header: "Name", key: "name" },
    { header: "Status", key: "healthy", valueFormatter: value => value ? "Healthy" : "Needs attention" }
  ],
  bom: true, quote: "strings", nullValue: "missing"
});
```

`quote` defaults to `"minimal"`; `"all"` quotes every field and `"strings"` always quotes string values. Required delimiter, quote and newline escaping applies in every mode. `nullValue` replaces null/undefined values and receives the same protection and quoting as other strings. Formatters run on data only; their context contains the one-based data `row`, one-based `columnIndex`, column definition and original projected `values`. Both Blob and caller-owned sink APIs share these options. CSV retains long text that exceeds Excel's per-cell text limit.

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

Pass an `AbortSignal` and synchronous `onProgress` to the writers. CSV/XLSX progress counts exclude headers; XLSX row events identify the sheet and completion reports the total output bytes. Throwing from a callback fails the export. Cancellation returns producer iterators and rejects pending input or sink operations. A cancelled `finish()` or `toBlob()` releases owned output and rejects its returned promise with the original signal reason, including cancellation between calls. Give I/O producers and destinations the same signal so they can release their own resources.

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

`openEntry(name)` returns an appendable entry sink with `write` and `close`; close it before starting the next entry or finishing the archive. `maxOutputBytes` can bound actual archive bytes. Blob-returning APIs retain output proportional to file size; stored output can be substantially larger. Streaming input does not imply constant memory for a final Blob. Readers are outside the current API.

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

`OpcPackage({ sink })` supports `openPart(uri, contentType)` for direct incremental output. Close each part before opening the next; register relationships and deferred parts as usual, then call `finish()`. It returns the archive byte count without closing the sink. A package with a caller-owned sink uses `finish()` rather than `toBlob()`.

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

`dataValidation` is a reserved worksheet option. Supplying it, including an empty array, throws `NotSupportedError` instead of silently ignoring it. There is no DOCX, PDF, PPTX, reader, formula engine or report-grid adapter in this package's current writer contract. Row/cell callbacks resolve highlighting during export; `conditionalFormats` separately writes the supported native rules for Excel to evaluate. The [roadmap](../Docs/ROADMAP.md#officeimo-javascript) records the remaining work.

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
