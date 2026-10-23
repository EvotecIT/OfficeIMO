# OfficeIMO JavaScript

`@evotecit/officeimo` provides strict TypeScript document libraries for evergreen browsers, Web Workers and Node 18 or newer. It writes streaming XLSX workbooks, UTF-8 CSV and paginated PDF tables with no runtime dependencies. One package owns seven public layers: `core`, `zip`, `xml`, `opc`, `xlsx`, `csv` and `pdf`. The optional `execution` entry connects host-owned workers to those writers with bounded row and byte messages.

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

## Export a table

Use the same columns for XLSX, CSV and PDF. Object columns select a typed key or compute a synchronous value; array rows use positional columns. A getter can return `ExportCell` to keep a typed value together with display text and presentation.

```ts
import { writeXlsx, writeCsv, ExportCell, saveBlob } from "@evotecit/officeimo";
import type { Column } from "@evotecit/officeimo";

interface Sale { customer: { name: string }; amount: number; seen: Date; }
const columns = [
  { header: "Customer", key: "customerName", value: (row: Sale) => row.customer.name },
  { header: "Amount", key: "amount", type: "number", format: "0.00",
    value: (row: Sale) => new ExportCell(row.amount, {
      text: row.amount.toFixed(2) + " USD",
      presentation: row.amount < 0 ? { background: "FFC7CE" } : {}
    }) },
  { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }
] satisfies readonly Column<Sale>[];
const rows: Sale[] = [{ customer: { name: "Łódź 🧪" }, amount: 12.5, seen: new Date() }];

saveBlob(await writeXlsx(rows, {
  columns, dateMode: "utc",
  sheet: { name: "Sales", table: {}, freezeHeader: true }
}), "sales.xlsx");
saveBlob(await writeCsv(rows, { columns }), "sales.csv");
```

`Column<Sale>` checks literal keys against `Sale`. Getters support nested fields and calculated values; their optional `key` is a stable report identifier for totals and conditional ranges. Unselected domain fields can be arbitrary objects. Arrays are positional unless every column has a getter; that explicit projection can select from a wider array. Dots in literal keys do not traverse objects. TypeScript 5.4 or newer is required for the typed writer signatures.

| Task | Entry point | Completion |
| --- | --- | --- |
| One Excel table | `writeXlsx(rows, options)` | `Blob` |
| One CSV table | `writeCsv(rows, options)` | `Blob` |
| One paginated PDF table | `writePdf(rows, options)` | `Blob` |
| Write a table to a destination | `writeXlsxTo(rows, destination, options)` / `writeCsvTo(...)` / `writePdfTo(...)` | `{ rows, columns, bytes }` |
| Multiple worksheets, registered styles, images or extra parts | `new Workbook(options)`, `addWorksheet`, `addRows` | `toBlob()` or streamed `finish()` |
| Export an installed DataTables grid | Optional integration below | Same writers and destination ownership |
| Export a CanopyX record capture | `exportCanopy` / `writeCanopyTo` in the optional CanopyX entry | `Blob` or streamed `{ rows, columns, bytes }` |

The table writers accept synchronous iterables and async iterables. XLSX defaults to the sheet name `Data`, a bold header, filtering when a header is present, and width sampling of up to 100 rows clamped to 6–54 characters. `sheet` supplies the existing worksheet layout options. Workbook-local style indexes belong to the advanced `Workbook` API; portable `ExportCell` presentation and row/cell style patches work with the table helper.

An `ExportCell` can carry an external link alongside its typed value, display text and presentation:

```typescript
new ExportCell(12.5, {
  text: "12.50 USD",
  presentation: { color: "0563C1" },
  link: { target: "https://example.com/invoices/42", tooltip: "Open invoice 42" }
});
```

XLSX keeps the numeric value and writes a cell hyperlink. PDF renders the display text and writes a native link annotation over each cell fragment, including fragments continued onto another page. CSV keeps its selected raw/display value mode. Links accept absolute HTTP, HTTPS and mailto targets; credentials, whitespace and other schemes are rejected. Resolve relative report links against an explicit host base before creating the cell. The target and tooltip are copied when the cell is created, and links do not add automatic font styling. An XLSX cell using the oversized-text preservation link cannot also carry an external link; that conflict fails explicitly.

An async generator is consumed once. Use a source factory for separate exports rather than passing the same generator twice:

```ts
async function* loadSales(): AsyncGenerator<Sale> {
  // Fetch pages here and yield typed Sale records, including Date values.
  for (const row of rows) yield row;
}
const excel = await writeXlsx(loadSales(), { columns });
const csv = await writeCsv(loadSales(), { columns, valueMode: "display" });
```

Data callbacks use zero-based `rowIndex` and `columnIndex`, excluding titles, headings and totals. XLSX additionally supplies one-based `worksheetRow` and `sheetName`. A getter runs once per selected cell per export, before formatting, preservation and resource checks on the resolved value. Getters are synchronous; perform asynchronous loading in the row source. Missing object values produce empty cells. Invalid nested objects, asynchronous getter results and extra unprojected array values fail visibly.

## DataTables Excel, CSV and PDF exports

The optional `@evotecit/officeimo/integrations/datatables` entry connects an installed DataTables/Buttons instance to the XLSX, CSV and PDF writers. It imports neither DataTables nor jQuery and uses no third-party document writer. Importing the main OfficeIMO package does not load the integration.

```js
import DataTable from "datatables.net";
import "datatables.net-buttons";
import { registerDataTablesButtons } from "@evotecit/officeimo/integrations/datatables";

registerDataTablesButtons(DataTable, {
  exportOptions: { columns: ":visible" },
  columnOptions: { 2: { type: "number", format: "0.00" } },
  xlsx: { sheet: { table: { style: "TableStyleMedium2" }, freezeHeader: true } },
  onError: error => console.error("Export failed", error)
});
const table = new DataTable("#report", {
  layout: { topStart: { buttons: [
    { extend: "officeimoExcel", filename: "Report" },
    { extend: "officeimoCsv", filename: "Report" },
    { extend: "officeimoPdf", filename: "Report", orientation: "landscape" }
  ] } }
});
```

Register once, after installing Buttons. Processing clears when saving completes or before calling `onError`. The default destination downloads a Blob; `save(blob, filename)` replaces delivery. Each button can override registration options through its `officeimo` property. Familiar `filename`, `exportOptions` and `footer: false` properties are supported. Filename patterns use Buttons' `exportInfo()` title substitution and sanitization: `"*"` uses the document title, and a synchronous `filename(configuration, table)` callback can return a pattern. Excel and CSV buttons reject native `customize`, `header: false`, `title`, `messageTop` and `messageBottom`; express Excel report layout through `officeimo.xlsx.sheet`. The PDF button accepts the title, messages and heading settings described below, with typed layout options in `officeimo.pdf`.

For application-owned buttons, `exportDataTable(DataTable, table, "xlsx" | "csv" | "pdf", options)` returns a Blob. `writeDataTableTo(DataTable, table, format, destination, options)` accepts a caller-owned `ByteSink` or `WritableStream<Uint8Array>` and returns `{ rows, columns, bytes }`. Both accept `signal`, `onProgress`, shared resource `limits`, and format-specific `xlsx`, `csv` or `pdf` options. XLSX uses the same settings as `writeXlsx`: set workbook policy through `xlsx.dateMode` and report layout or the worksheet name through `xlsx.sheet`, for example `xlsx: { dateMode: "utc", sheet: { name: "Inventory", freezeHeader: true } }`. A defined numeric shared limit overrides its matching format limit; `undefined` leaves that format's ceiling intact. XLSX defaults to a bold header, filtering and bounded width sampling of up to 100 rows, clamped to 6–54 characters. Explicit column widths avoid sampling that column.

`project(value, { rowIndex, columnIndex, sourceRowIndex, sourceColumnIndex })` resolves each rendered/normalized value into a scalar or an `ExportCell`. Data indexes are zero-based positions in the selected export; source indexes identify DataTables rows/columns. `columnOptions` is keyed by DataTables column index. Use an orthogonal renderer for raw numbers/dates when display text contains currency, markup or localized dates. The adapter does not infer types from formatted strings. For example:

```js
import { exportDataTable, ExportCell } from "@evotecit/officeimo/integrations/datatables";
const controller = new AbortController();
const blob = await exportDataTable(DataTable, table, "xlsx", {
  signal: controller.signal,
  exportOptions: { columns: ":visible", orthogonal: "export", modifier: { selected: null } },
  columnOptions: { 2: { type: "number", format: "0.00" } },
  project: (value, cell) => cell.sourceColumnIndex === 2
    ? new ExportCell(value, { presentation: { background: "E2F0D9" } }) : value,
  limits: { maxRows: 100000, maxCells: 2000000, maxOutputBytes: 100000000 }
});
```

CSS classes and computed DOM styles are not copied automatically. CSV retains values or explicit display text and formula protection; it cannot carry colors or Excel layout.

| Contract | Batched mode, the default | Compatibility mode |
| --- | --- | --- |
| Row/column selectors, search/order modifiers | Captured through public DataTables APIs | Passed to `buttons.exportData()` |
| Select extension | Selected rows when a selection exists; `selected: null` exports all matching rows | Same installed Buttons behavior |
| Orthogonal rendering and synchronous `format.header/body/footer` | Supported; cleanup uses installed `Buttons.stripData` | Supported by installed Buttons |
| Whole-matrix `customizeData` | Rejected | Supported; cannot combine with source-index `project` |
| Body memory | Selected row indexes and bounded cell batches | Complete Buttons matrix, then batched writing |
| Horizontal grouped headers | Preserved through shared column group paths | Same |
| Vertical spans, blank spanning groups, adjacent separate equal group paths | PDF preserves structured headings by default. XLSX/CSV preserve them with `headings: "structured"`; CSV also needs an explicit span rule. `headings: "leaf"` selects one heading row | Same |
| Footers | Structured exports preserve up to 16 physical rows; XLSX native table/filter constraints and CSV flattening apply. `includeFooter: false` omits footers | Same |
| Server-side processing | Rejected unless `serverSide: "loaded"` acknowledges loaded rows only | Same |

Qualified pairs are DataTables 2.3.7/Buttons 3.2.6 and DataTables 3.1.3/Buttons 4.1.2, including Select and ColReorder. TypeScript 5.9.3 consumer checks use bundler resolution. The newer pair has duplicate upstream declaration index signatures and requires `skipLibCheck`; OfficeIMO's declarations remain strictly checked. This declaration limit is separate from browser interoperability.

`createDataTablesExport(DataTable, table, options)` exposes readonly `columns`, heading rows, optional footer, `rowCount` and a single-use async `rows` source. `headings: "structured"` additionally exposes `headerStructure`/`footerStructure` for the writers' span matrices. Selection/headings are captured immediately; batched body values are read during iteration. Each batch replaces bounded cell indexes through public API result-set operations, without rescanning the table's rows. Batched exports require one table per API instance. Keep table data and column layout stable until export finishes. `batchRows` defaults to 4,096; `maxBatchCells` defaults to 65,536, can lower the effective row count per batch and must fit one row. Cancellation is checked between cells and yielding occurs during formatting and between batches. A synchronous DataTables renderer cannot be interrupted while it runs.

For an XLSX heading whose final physical row contains a merged or covered cell, use `headings: "structured", xlsx: { sheet: { autoFilter: false } }`. Native filters and Excel tables require an unmerged final row matching the declared column labels. Structured footer rows sit outside a native table. For CSV, choose `headings: "structured", csv: { spanMode: "repeat" }` to repeat each anchor across its covered positions, or `"anchors"` to leave covered fields empty. Each physical row is retained. Existing grouped defaults remain unchanged.

The PDF button accepts native `title`, `messageTop`, `messageBottom`, `pageSize`, `orientation`, `header: false` and `footer: false` settings. Omitted titles/messages inherit registration defaults; an explicit `null` clears that field. Text and filename resolution use the installed Buttons API. Supply writer options through `pdf` in registration defaults, or `officeimo: { pdf: ... }` on one button. Per-button PDF settings override only the supplied fields, retaining omitted fonts, metadata and layout settings; nested font profiles and footer objects replace their whole field. Native `header: false` and `footer: false` also suppress replacement table headings and footers. Native document-definition `customize` callbacks are rejected; use typed PDF options and portable `ExportCell` presentation. `pdf.footer` replaces the grid's footer when supplied. Unicode PDF exports require `pdf.fonts`, as described below.

Compressed PDF streams write output chunks progressively with an indirect length. Tagged reports serialize completed logical rows, paragraphs and graphics during export, retaining open content and compact references needed by the PDF object graph. ParentTree indexing uses bounded nodes. Total PDF memory still includes fonts, images, the current page, object offsets and retained parent references; Blob output additionally retains the completed file. Resource limits remain cumulative across completed structures.

Direct sink output avoids retaining the finished file. Blob exports and registered download buttons retain it; caller-owned sinks own partial bytes after failure. The grid still holds its input. Server-side full-data export needs an application-owned paged source passed directly to the writers with the server's filter/order contract.

DataTables and its DOM stay on the page. A host-owned worker can consume portable columns and bounded row batches from the source: request/acknowledge batches and output chunks rather than cloning the full matrix. Reconstruct `ExportCell` in the worker because structured cloning loses its brand. The [verification guide](../Build/BrowserExports/README.md#datatables-integration-and-comparisons) covers workers, cancellation, fallback and reproducible comparisons. Measurements describe the tested workload/browser, without a universal speed claim.

## CanopyX record exports

The optional `@evotecit/officeimo/integrations/canopyx` entry consumes CanopyX's immutable `prepareExport` capture. CanopyX owns the accepted query, column order, selection, source paging, revision checks and resolved presentation. OfficeIMO owns file generation. Neither package imports the other at runtime, and importing the main OfficeIMO package does not load the adapter.

```typescript
import { exportCanopy } from "@evotecit/officeimo/integrations/canopyx";

const capture = grid.prepareExport({ scope: "filtered", values: "raw" });
const reportOptions = {
  tones: {
    neutral: {},
    success: { background: "E2F0D9" },
    warning: { background: "FFF2CC", color: "9C6500" },
    danger: { background: "FFC7CE", color: "9C0006", bold: true }
  },
  columnOptions: {
    amount: { format: "#,##0.00", width: 18 },
    updated: { format: "yyyy-mm-dd hh:mm:ss.000", width: 42, wrapText: true }
  },
  xlsx: { sheet: { title: { text: "Inventory" }, table: {} } },
  limits: { maxRows: 250_000, maxCells: 5_000_000 }
};
const blob = await exportCanopy(capture, "xlsx", reportOptions);
// Deliver the Blob through the host's download or storage policy.
```

`exportCanopy(capture, format, options)` accepts `"xlsx"`, `"csv"` or `"pdf"`. `writeCanopyTo(capture, format, destination, options)` writes to a caller-owned sink or `WritableStream`; failed partial output remains caller-owned. Format options live in `xlsx`, `csv` and `pdf`, with common cancellation, limits and progress at the top level. Common limits override matching format limits. For Unicode PDF text, supply embedding-permitted TrueType fonts through `pdf.fonts`.

`createCanopyExport` retains `type: "number"` for native numeric columns in a raw capture, so grouped metrics can use its columns directly. Display captures keep their rendered strings. Datetime metadata remains unforced because the preservation profile can intentionally produce exact ISO text rather than a `Date`.

`columnOptions` is keyed by captured column ID. Settings for unselected columns are ignored, so one host format map can serve different visible/selected captures. `PdfFont` is exported from this entry as well as the PDF layer for constructing caller-supplied fonts.

| Captured contract | Output behavior |
| --- | --- |
| Column IDs, titles and order | Preserved without re-reading the viewport or UI settings |
| `values: "raw"` | Numeric, boolean and null values remain typed; CSV retains original ISO datetime strings |
| `values: "display"` | Captured display strings become literal exported values; CSV formula protection still applies |
| Resolved text | PDF uses it; XLSX retains raw values and uses explicit number/date formats |
| Semantic tones | Host-supplied portable styles; cell tone overrides overlapping row tone properties |
| Absolute HTTP(S)/mailto links | Native XLSX relationships and PDF annotations; CSV remains ordinary text |
| Datetime columns | XLSX defaults to typed dates at exact millisecond precision and literal ISO text for finer precision |
| CSS pixel width | Never treated as an Excel character width; `columnOptions.width` supplies explicit character widths and `pdf.columnWidths` supplies points |
| `recordCount` | Used for progress when known and checked against the actual enumeration; unknown counts remain unknown |

Excel/PDF reject unsupported presentation diagnostics and unmapped tones by default. The reporting runtime supplies semantic captures. For an ordinary grid, custom renderer or CSS-only presentation, choose `unsupportedPresentation: "text"` to accept the captured text fallback and optionally observe each diagnostic with synchronous `onDiagnostic`. That fallback uses capture data; it does not scrape custom DOM renderer output. CSV defaults to text because the format has no visual styling; choose `"reject"` explicitly when diagnostics should stop it. No unbounded diagnostic list is retained.

`datetime: "typed"` requires exact millisecond precision; `"text"` keeps the original ISO value. The default `"preserve"` retains finer timestamps and dates outside Excel's 1900–9999 range as text rather than rounding or dropping them. An ISO value converted to a typed Excel date must carry an explicit time zone. XLSX defaults to the capture's UTC/local clock choice; `xlsx.dateMode` can override it. PDF displays the captured text, and CSV raw mode retains the original string and its zone.

`createCanopyExport(capture, format, options)` exposes frozen columns, `rowCount` and fresh `rows()` enumerations for applications using the format writers directly. It retains no whole-table body. Each enumeration passes its cancellation signal into CanopyX's page requests, and destination backpressure stops reading ahead. CanopyX's source may retain identity/cursor bookkeeping for revision and duplicate detection; streamed document output does not imply constant memory for that entire source.

Host export callbacks can use the already-captured enumeration:

```typescript
import { writeCanopyTo } from "@evotecit/officeimo/integrations/canopyx";
import type { GridHostOptions } from "@evotecit/canopyx/reporting";

const onExport: GridHostOptions["onExport"] = async ({ capture, format, signal, reportProgress }) => {
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") {
    throw new Error("Unsupported export format");
  }
  const destination = await openDestination(format); // Host-owned delivery.
  await writeCanopyTo(capture, format, destination, {
    ...reportOptions, signal,
    onProgress: progress => reportProgress?.("generating", progress.rows)
  });
};
```

This entry exports grid records. Collapsed visual groups do not remove their records. The adapter exports records in the captured order; use the [grouped report writers](#grouped-analytical-reports) when an explicitly group-ordered capture needs summary rows and Excel outlines. Workers consume a host-owned bounded capture/row bridge; a live CanopyX grid remains on the page. Pass the writer's signal to the same capture enumeration and release the bridge when the host cancels it.

## PDF tables

`writePdf` and `writePdfTo` use the same `Column<T>` projection and `ExportCell` values as Excel and CSV. The PDF owner lays out and writes one page at a time, repeats grouped headings, splits oversized data rows across pages and appends a table footer. `writePdfTo` awaits the destination and retains page references, font mappings and the current page rather than the complete report body. Blob mode also retains the finished file.

```ts
import { writePdf, PdfFont, ExportCell, saveBlob } from "@evotecit/officeimo";

// Host-owned, embedding-permitted static TrueType font.
const response = await fetch("/fonts/report-regular.ttf");
if (!response.ok) throw new Error("Report font could not be loaded");
const regular = await PdfFont.create(new Uint8Array(await response.arrayBuffer()));
const report = await writePdf([
  { city: "Łódź", amount: 12.5 }, { city: "Gdańsk", amount: -2 }
], {
  title: "Sales report", fonts: { regular },
  columns: [
    { header: "City", key: "city", groups: ["Sales"] },
    { header: "Amount", key: "amount", groups: ["Sales"], alignment: "right",
      value: row => new ExportCell(row.amount, {
        text: row.amount.toFixed(2) + " USD",
        presentation: row.amount < 0 ? { color: "9C0006", background: "FFC7CE" } : {}
      }) }
  ],
  footer: { values: ["Total"], totals: { amount: "sum" } },
  pageFooter: "Confidential", pageSize: "A4", orientation: "portrait",
  limits: { maxRows: 100_000, maxPages: 5_000, maxOutputBytes: 128_000_000 }
});
saveBlob(report, "sales.pdf");
```

PDF uses `ExportCell.text` first, then a synchronous `formatValue(value, context)` callback, then scalar text. Dates default to UTC ISO strings; null and nonfinite numbers are blank. Empty report titles, messages and page decorations are omitted without reserving space or consuming the cell/text budget. Excel number-format strings and live conditional rules are not evaluated by the PDF writer. Resolve display text and highlights in the shared getter or DataTables `project` callback when they must appear in both formats. Footer totals operate on typed numeric values and support `sum`, `count`, `average`, `min` and `max`.

Lengths are points, with 72 points per inch. Defaults are A4 portrait, 36-point margins, 9-point text, 4-point cell padding and page numbers. Named sizes are A3, A4, A5, LETTER, LEGAL and TABLOID; custom `{ width, height }` sizes are accepted. `columnWidths` supplies point widths; otherwise `Column.width` uses the width of the font's zero character. `wideTable: "fit"` fits column widths within the printable width while retaining font size and wrapping text. `"reject"` rejects widths that exceed the page. A column that cannot fit one glyph fails explicitly.

`Column.groups` provides simple grouped headings. `headerRows` replaces them with a rectangular `TableSpanRows` matrix: each anchor declares `{ value, columnSpan?, rowSpan? }`, and every covered position is `null`. `footer.rows` uses the same model instead of `footer.values`/`totals`. Matrices support up to 16 rows and must cover the declared columns without overlap. Repeated headings must leave space for data; a structured footer must fit below the headings on one page. Titles/messages, page headers/footers and final `Page N of M` labels are separate from the table. Message-only continuation pages retain page decorations without repeating table headings. Page decoration callbacks receive the [page and section context](#compose-a-pdf-report) and return a synchronous string. Page-number space is measured from the selected font and `maxPages`; for large fonts or narrow pages, reduce that limit or set `pageNumbers: false`. Table and paragraph heights respect the selected font face's ascent and descent.

Without supplied fonts, PDF uses standard Helvetica with WinAnsi text. For Polish, Greek, Cyrillic, CJK or other supported Unicode scalars, supply an embedding-permitted static TrueType font containing those glyphs. `PdfFont` copies and validates bytes once and can be reused. For cooperative parsing, use `await PdfFont.create(bytes, { signal, maxBytes })`; its default intake limit is 16 MiB and can be raised up to the constructor's 64 MiB ceiling. The source snapshot is taken before yielding, so later caller mutation cannot change the font. The synchronous constructor remains available. PDF finalization also cooperates with cancellation while building the glyph subset. Optional `bold`, `italic` and `boldItalic` faces preserve their real glyphs; missing faces use synthetic emphasis. Used glyphs and composite dependencies are embedded as a subset, with Unicode extraction maps. Font permissions can require full embedding or prohibit embedding, which fails visibly. Collections, CFF/WOFF and variable fonts require conversion to static TrueType before use. The library neither fetches fonts nor reads installed fonts; font licensing belongs to the host.

Call `preflightPdfText(text, { fonts: { regular }, bold: true })` from `@evotecit/officeimo/pdf` to check a known scalar-layout run before export. It returns `{ supported: true }` or the first issue with `code`, `message`, and a zero-based UTF-16 `index` in the original input; `scalar` identifies the original character and is absent for an unpaired surrogate. Codes distinguish `embedded-font-required`, `missing-glyph`, `unsupported-layout`, `unsupported-control` and `unpaired-surrogate`. The check uses scalar font selection, line-break and tab rules; it does not invoke an optional shaping provider. It checks only the supplied string, leaves row sources untouched, and does not establish wrapping, document resource limits or accessibility.

The default writer handles scalar text layout for Latin, Greek, Cyrillic, Han, Hiragana, Katakana, precomposed Hangul and common symbols, including supplementary Unicode where the font provides it. Keep a supplementary character in one rich-text run when its font, color or link metadata differs from adjacent text; splitting its UTF-16 pair across different presentation fails explicitly. If a supplied TrueType font lacks U+202F (narrow no-break space) but provides an inkless U+2009 (thin space), the writer uses that thin-space glyph and its advance while preserving U+202F in the PDF's Unicode map. Other missing glyphs and malformed surrogates fail explicitly. Combining sequences, bidirectional text and scripts requiring shaping need the optional profile below. It does not provide general HTML/SVG rendering or PDF/A. Composed documents support the bounded image and vector contracts below. Long text is wrapped and continued across pages without truncation; CRLF/CR become line breaks and tabs expand to four spaces for display.

`PdfLimits` includes shared row/cell/text/output ceilings plus `maxPages` (default 10,000), `maxColumns` (1,024), `maxCellCharacters` (1,000,000 UTF-16 units), `maxRowLines` (100,000), `maxFontBytes` (16 MiB across unique supplied fonts), `maxPageBytes` (8 MiB of drawing commands and link dictionaries), and `maxHyperlinks` (100,000 native annotations). A continued linked cell creates one annotation per page fragment; repeated linked headings also consume that limit. PDF cell/text budgets count rendered strings, headings, totals, bookmark titles and page decorations. Captions count each painted occurrence; an unused continuation caption consumes no cell/text budget. `maxRows` counts source rows. Resource failures never silently remove rows, columns or text. Native deflate compresses streams and bounded batches of completed tagged structure objects when available. Tagged compaction uses PDF 1.5 object/xref streams inside the existing PDF 1.7 document profile; supported readers must understand that representation. `compression: false` or unavailable native deflate retains plain objects and classic cross-references. Object batches are bounded to 256 members and 32 KiB decoded bytes; page/object/font references and caller-owned output remain document-scale state.

Workers can call the same writers without a DOM. For a portable worker-to-page handoff, use `writePdfTo` and transfer byte chunks with acknowledgements; this also avoids WebKit worker Blob-read restrictions. Keep the row source and destination bridge bounded, and pass an `AbortSignal` to stop pending input or output. The caller owns disposal of partial bytes after failure, and the library releases borrowed stream locks without closing or aborting the destination.

### Readable wide tables

Use horizontal bands when an inventory has more columns than one page can display comfortably. The writer keeps each requested width instead of shrinking all columns to one page. Declare identifier columns to repeat at the left of every band:

```js
import { writePdf, saveBlob } from "@evotecit/officeimo";

const rows = [{ service: "WEB-000001", site: "Warsaw", owner: "Operations", environment: "Production",
  availability: 99.98, latency: 4.9, errors: 0, cpu: 21, memory: 48, disk: 35, requests: 12340, checks: 120 }];
const columns = [
  { header: "Service", key: "service", rowHeader: true }, { header: "Site", key: "site" },
  { header: "Owner", key: "owner" }, { header: "Environment", key: "environment" },
  { header: "Availability %", key: "availability", type: "number" },
  { header: "Latency ms", key: "latency", type: "number" },
  { header: "Errors", key: "errors", type: "number" }, { header: "CPU %", key: "cpu", type: "number" },
  { header: "Memory %", key: "memory", type: "number" }, { header: "Disk %", key: "disk", type: "number" },
  { header: "Requests", key: "requests", type: "number" }, { header: "Checks", key: "checks", type: "number" }
];
const report = await writePdf(rows, {
  columns,
  tagged: { language: "en-US" },
  wideTable: { mode: "bands", minimumColumnWidth: 90, repeatColumns: ["service", "site"] },
  limits: { maxRows: 100_000, maxPages: 10_000, maxTableBands: 16, maxTableWindowRows: 128 }
});
saveBlob(report, "service-inventory.pdf");
```

`minimumColumnWidth` defaults to 72 points for inferred widths. Explicit `columnWidths` remain exact point widths. `repeatColumns` accepts exact stable column keys or zero-based positions in this table's captured selection. Numeric-only lists must be sorted and unique; string `"0"` is a key, while numeric `0` is a position. Mixed selections resolve to declared column order. Missing, ambiguous or duplicate targets reject before reading rows. Hidden fields and original grid positions are never consulted. If the identifiers and one further column cannot fit, the writer rejects the layout rather than clipping or shrinking it.

Each bounded record window is emitted through its bands before the next window. Long records share one logical fragment window while each band paints only the height its visible cells need. A continuation band with no remaining non-identifier content is omitted; initial rows, structured spans and link obligations are retained. Repeated identifier cells must fit below the repeated headings and group context on a fresh page; taller identifiers are rejected. A short final fragment retains the full identifier and enough row height to display it. Sources are consumed once, values and formatters are captured once, and numeric totals are accumulated once. Repeated headings, identifiers, links and page decorations still consume the document's physical output budgets. `maxTableBands` defaults to 64 per table; `maxTableWindowRows` defaults to 256 retained row fragments per window. These bounds do not imply constant memory for fonts, page references or caller-owned destinations.

The same `wideTable` option applies to `pdfTable`, `writeGroupedPdf` and `pdfGroupedTable`. Grouped tables add a leading label column: repeat `[0, "service"]` to retain that label and the selected service key, or `[0, 1]` for the label and first data position. Continued analytical bands retain the active ancestor context. A deliberately small `maxTableWindowRows` can separate the last detail and closing subtotals even when they would fit on one page; the retained-fragment bound takes precedence. Structured headings and footers preserve their intersecting span labels in each band. An indivisible footer span matrix must fit the window bound and the physical page. With `tagged: { language }`, physical bands share one logical table: cells follow source-column order within each source row, and repeated visual identifiers do not create duplicate logical text.

### Display language

PDF and CSV accept `booleanText: { true: "Tak", false: "Nie" }`; defaults remain `True`/`False`. Explicit `ExportCell.text` or a column/format callback takes precedence. These are display choices: native XLSX booleans remain boolean cells. CSV applies normal quoting, text limits and formula protection to custom labels.

PDF page counters accept `pageNumberLabels: { prefix: "Strona ", separator: " z ", suffix: "" }`. The default is `Page 1 of 3`. Labels are captured once and use the same font, geometry and text limits as the counter. Number/date formatting belongs to the host's explicit formatter or `ExportCell.text`; schema keys, Excel formulas and invariant defaults do not depend on translated headings or the machine's language.

### Compose a PDF report

`writePdfDocument` combines sections, paragraphs, lists, independently paginated tables, images, vector paths, portable drawings and explicit page breaks. `writePdfDocumentTo` writes the same blocks to a caller-owned destination. `pdfTable` checks each table's columns and getters against its row type.

```ts
import { writePdfDocument, pdfTable, saveBlob } from "@evotecit/officeimo";
import type { PdfDocumentBlock, PdfListItem } from "@evotecit/officeimo";

const columns = [
  { header: "Service", key: "service" },
  { header: "Owner", key: "owner" },
  { header: "Status", key: "status" },
  { header: "Latency (ms)", key: "latency", alignment: "right" }
] as const;
async function* findings(): AsyncGenerator<PdfListItem> {
  yield { text: [{ text: "Review service latency. ", bold: true },
    { text: "Assign an owner and record the recovery action." }], keepTogether: true };
  yield { text: "Retain this inventory as the baseline for the next review." };
}
async function* inventory() {
  yield { service: "Directory", owner: "Operations", status: "Review", latency: 120 };
  yield { service: "Resolver", owner: "Network", status: "Healthy", latency: 8 };
}
async function* blocks(): AsyncGenerator<PdfDocumentBlock> {
  yield { kind: "section", title: "Service review", level: 1, headingStyle: "title" };
  yield { kind: "paragraph", text: "Two services were checked.", style: "body" };
  yield { kind: "paragraph", text: "Findings", headingLevel: 2, style: "heading",
    bookmark: { title: "Findings", level: 2 } };
  yield { kind: "list", ordered: true, itemStyle: "finding", items: findings() };
  yield { kind: "vector", width: 120, height: 40, fill: "217346", commands: [
    { op: "move", x: 1, y: 39 }, { op: "line", x: 119, y: 39 },
    { op: "line", x: 119, y: 1 }, { op: "close" }
  ], alternativeText: "Illustrative rising service-health indicator.",
    caption: { text: "Figure 1. Review indicator.", style: "caption" } };
  yield { kind: "section", title: "Inventory appendix", level: 2,
    headingStyle: "heading", orientation: "landscape" };
  yield { ...pdfTable(inventory(), { columns }),
    caption: { text: "Table 1. Checked services.", style: "caption" } };
  // Omitted geometry restores the document's portrait defaults.
  yield { kind: "section", title: "Next actions", level: 2, headingStyle: "heading" };
  yield { kind: "list", itemStyle: "finding", items: [
    { text: "Assign each open finding to an owner.", keepTogether: true }
  ] };
}
const report = await writePdfDocument(blocks(), {
  pageSize: "A4", orientation: "portrait", tagged: { language: "en" },
  paragraphStyles: {
    title: { fontSize: 20, presentation: { bold: true, color: "163C75" } },
    heading: { fontSize: 14, presentation: { bold: true }, keepWithNext: true },
    body: { fontSize: 10, spaceAfter: 10 },
    finding: { fontSize: 10, spaceAfter: 8 },
    caption: { fontSize: 9, presentation: { italic: true, color: "475569" } }
  },
  pageHeader: context => context.sectionTitle ?? "Service review",
  pageFooter: context => `Section ${context.sectionNumber} / ${context.sectionPageNumber}`
});
saveBlob(report, "service-review.pdf");
```

Document options own fonts, padding, default page geometry, decorations and limits. Each table owns its columns, repeated headings, widths, presentation, formatter and totals. The existing font and supported-script rules apply to every block. `writePdf` and `writePdfTo` remain convenient entry points for one table. The [executable service report](../OfficeIMO.JavaScript.Examples/report-engine-job.mjs) shows the same composition over a caller-owned selected row source and summary metadata.

A `section` is a flat boundary for following blocks. Its `title` and optional `level` (1–6, default 1) supply a visible heading and native bookmark by default. `headingStyle` selects the heading's named appearance; `includeHeading: false` hides that heading and `bookmark: false` omits its outline entry. A boundary starts a fresh page after visible content; first or consecutive unpainted sections do not create blank transition pages. Each section derives omitted `pageSize`, `orientation`, `margins`, `pageHeader`, `pageFooter` and `pageNumbers` from the captured document defaults, rather than the preceding section. Page numbers remain continuous. Header/footer callbacks receive `pageNumber`, one-based `sectionNumber` and `sectionPageNumber`, and optional `sectionTitle`; they must return text synchronously. A paragraph's `bookmark: { title, level? }` follows its actual first painted line, including a heading moved by a keep rule.

`paragraphStyles` is a captured record of names to `fontSize`, `presentation` (`bold`, `italic`, text `color` and horizontal `alignment`), spacing and keep rules. Paragraph `style` selects one definition; explicit fields override it, and explicit presentation members merge with the named presentation. Styles have no inheritance and do not imply semantic headings: use `headingLevel` explicitly. Unknown names fail. The registry is shared across sections.

A `list` consumes synchronous or asynchronous `items` once. Each item has `text: PdfText`, optional `style` and `keepTogether`; its style replaces the list's `itemStyle` name. `ordered: true` uses decimal markers, with optional positive safe-integer `start`; otherwise it uses bullets. `start` is valid only for ordered lists. `indent` and `markerGap` are point lengths for hanging body text. The marker stays with the first body line and is not repeated on continuation pages. Nested lists and other nested blocks are unsupported; styles resolving to `keepWithNext: true` are rejected for items. An empty list adds no content or page break. An empty document still produces one PDF page; an explicit `pageBreak` remains a real page break.

Attach `caption: { text, style?, fontSize?, presentation?, spaceBefore?, spaceAfter? }` to a table, image, vector or drawing block. It appears before the owned content and stays with the first table heading/data fragment or the whole graphic. Captioned graphics require nonempty `alternativeText` and cannot be decorative, including in untagged output. Captions and kept content must fit together on a fresh page; reduce a graphic's display dimensions if necessary. Labels such as “Table 1” are authored text: automatic numbering, after-content captions and cross-references are unsupported.

The document-wide composition limits are `maxSections` (default 1,000), `maxBookmarks` (10,000), `maxListItems` (100,000) and `maxParagraphStyles` (256). They accept nonnegative safe integers and share the existing cumulative text, cell, page, font and output budgets. Section boundaries do not reset these budgets. Bookmark titles consume the text and cell budgets once as outline metadata, including when their visible heading is omitted.

Paragraph `text` can also be an array of `PdfTextRun` values. Each run supplies text and optional `bold`, `italic`, RGB `color`, an explicit `font` face or an HTTP(S)/mailto `link`. Wrapping preserves continuous text across run boundaries and gives each visible link fragment its own annotation:

```ts
const invoice = await writePdfDocument([
  { kind: "group", blocks: [
    { kind: "paragraph", text: "Invoice", headingLevel: 1, fontSize: 18 },
    { kind: "paragraph", text: [
      { text: "Customer: " }, { text: "Contoso", bold: true }
    ] }
  ] },
  { kind: "paragraph", keepTogether: true, spaceAfter: 12, text: [
    { text: "View the " },
    { text: "payment details", color: "0563C1", link: { target: "https://example.com/invoices/42" } }
  ] },
  pdfTable([{ item: "Services", amount: 120 }], {
    columns: [{ header: "Item", key: "item" }, { header: "Amount", key: "amount" }],
    formatRuns: (value, cell) => [{
      text: cell.columnIndex === 1 && typeof value === "number" ? value.toFixed(2) : String(value ?? ""),
      bold: cell.columnIndex === 1
    }]
  })
], { fonts: { regular }, tagged: { language: "en-US" } });
saveBlob(invoice, "invoice.pdf");
```

`formatRuns(value, context)` is a synchronous PDF-only display callback for table data. It leaves raw values available to totals; `ExportCell.text`, including an explicit empty string, takes precedence. Whole-cell links and inline run links cannot overlap. The same static-font and supported-script rules apply to every run; styled text alone does not add shaping support. An explicit `font` is an exact `PdfFont` face. Bold/italic flags apply synthetic emphasis to that face; omit them when the supplied face already contains the desired emphasis. Distinct configured and explicit font identities share the document's `maxFontBytes` budget, including unused configured faces.

Paragraphs accept `keepTogether`, `keepWithNext`, `spaceBefore` and `spaceAfter`. Spacing defaults to zero points before and eight after. `keepTogether` moves an intact paragraph to a fresh page when needed; `keepWithNext` keeps it with the next visible paragraph line, first table fragment or complete graphic. A `group` contains paragraphs only and stays together by default. Nested groups and other block kinds are rejected; a final visible kept paragraph cannot reach outside its group. Content that cannot fit on a fresh page fails explicitly. `maxTextRuns` defaults to 10,000 per cell/paragraph and across a paragraph group or kept chain; `maxGroupBlocks` defaults to 256. Groups and kept chains also share `maxCellCharacters` so pagination cannot retain an unbounded report body.

Images accept `{ format: "png", data }`, `{ format: "jpeg", data }` or decoded `{ format: "rgb" | "rgba", width, height, data }` pixels. Raw pixels are row-major, 8 bits per channel, without row padding or premultiplied alpha. JPEG accepts 8-bit grayscale/RGB baseline, sequential and progressive frames. CMYK, ICC profiles, EXIF metadata and other component layouts fail explicitly; normalize color and orientation before supplying those JPEGs.

PNG intake accepts complete static, noninterlaced, 8-bit RGB/RGBA images with no color-profile metadata. It checks chunk CRCs, critical chunk order, scanline filters and the exact inflated byte length. RGB samples pass unchanged into the PDF's DeviceRGB image; straight alpha becomes a separate soft mask. Indexed/grayscale images, 16-bit samples, interlacing, animation, EXIF, `tRNS`, `sBIT`, and color/profile chunks (`sRGB`, `gAMA`, `cHRM`, `iCCP`, `cICP`, `mDCV`, `cLLI`) fail explicitly; normalize those inputs first. PNG decoding requires native `DecompressionStream("deflate")` and returns `code: "PLATFORM_UNAVAILABLE"` when unavailable. Arbitrary SVG bytes are not decoded.

An image's optional display `width` and `height` use points. One omitted dimension preserves the intrinsic aspect ratio; omitting both uses 96 DPI. The writer scales an oversized graphic down to the printable page and moves it intact to the next page when necessary. Source bytes are copied when the image is first consumed. Reusing that image object shares its PDF resource. Keep source objects and byte arrays stable until their block has been consumed.

Vector blocks use explicit top-left bounds and `move`, `line`, `cubic` and `close` commands. A block paints one path with optional six-digit RGB `fill` and `stroke`, `strokeWidth`, and `fillRule: "nonzero" | "evenodd"`. Coordinates must stay inside the declared bounds. Painting is clipped to those bounds. Images and vectors support `alignment: "left" | "center" | "right"`. This path contract does not interpret SVG markup, CSS, gradients, text or external resources.

Blocks and table rows are consumed once in order, with destination backpressure and cancellation. Pass a fresh source for each export. The result's `rows` sums the data rows across all tables; `columns` is the widest table. Row, cell, rendered-text, page, font, output and hyperlink limits cover the whole document. Graphics add `maxImageBytes` (32 MiB across distinct source image objects), `maxImagePixels` (16,777,216 pixels per image), and `maxVectorCommands` (100,000 commands across all vector blocks). Caller-owned destinations retain their normal partial-output and disposal contract.

Set optional `maxImageWorkingBytes` to reject an image before allocating its owned source snapshot, decoded planes and PNG scanlines. There is no default ceiling for this extra budget. PNG RGBA data is decoded directly into RGB and alpha planes with two scanlines; raw RGBA compacts its owned snapshot before writing the alpha mask. The budget excludes caller-owned input, finished output and native inflater/compressor allocations, so it does not bound total JavaScript heap or process memory. `maxImageBytes`, `maxImagePixels` and `maxOutputBytes` remain separate limits.

### Optional multilingual PDF shaping

The `@evotecit/officeimo/pdf/providers/harfbuzz` entry adapts host-loaded HarfBuzz and bidi engines. Importing it loads neither engine nor WASM. The default package and bundles retain no external runtime dependencies; applications choosing this profile install and load its engines explicitly:

```sh
npm install harfbuzzjs bidi-js
# TypeScript hosts also need the global type used by HarfBuzz's declarations:
npm install --save-dev @types/emscripten
```

```js
import * as harfbuzz from "harfbuzzjs";
import bidiFactory from "bidi-js";
import { PdfFont, writePdfDocument, saveBlob } from "@evotecit/officeimo/pdf";
import { createHarfBuzzTextShaper } from "@evotecit/officeimo/pdf/providers/harfbuzz";

async function loadFont(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error("Report font could not be loaded.");
  return PdfFont.create(new Uint8Array(await response.arrayBuffer()));
}
const [latin, arabic] = await Promise.all([
  loadFont("/fonts/Carlito-Regular.ttf"),
  loadFont("/fonts/NotoSansArabic-Regular.ttf")
]);
const shaper = createHarfBuzzTextShaper({ harfbuzz, bidi: bidiFactory(), language: "ar" });
try {
  const report = await writePdfDocument([{ kind: "paragraph", text: [
    { text: "Status: ", font: latin },
    { text: "السَّلَامُ عَلَيْكُمْ", font: arabic, color: "175F97" },
    { text: " (42)", font: latin }
  ] }], { fonts: { regular: latin }, textShaper: shaper, textDirection: "auto" });
  saveBlob(report, "multilingual-report.pdf");
} finally { shaper.dispose(); }
```

The qualified adapter uses harfbuzzjs 1.6.3 and bidi-js 1.1.0, with a Unicode 13 repertoire of Latin, Arabic, common and inherited characters. It supports static TrueType faces, combining marks, Arabic joining and mixed-direction horizontal text in composed paragraphs and table cells. Supply fonts covering each run; no font discovery or fallback occurs. Other scripts, later Unicode repertoire, vertical text, variable/color fonts and general Unicode conformance are outside this profile. The provider's versioned character table carries its Unicode data license notice. Its engines remain separately loaded MIT-licensed packages. Browser hosts serve their ESM/WASM assets over HTTP(S) with the required CSP; the classic local-file assets do not bootstrap these engines.

Shaping sees the whole physical paragraph and reorders each finalized line using paragraph direction. The writer owns word/cluster wrapping, pagination, glyph embedding, CID/Unicode mappings, links and logical structure. `textDirection` accepts `auto`, `ltr` or `rtl`. A presentation, font or link boundary inside an indivisible cluster fails explicitly. Top-level `title`, `messageTop`/`messageBottom`, page decorations, portable drawing text and `preflightPdfText` use scalar layout; use composed paragraphs for shaped headings or messages.

Each shaped physical paragraph is limited to 4,096 UTF-16 units and each shaping call to 32,768 glyphs. One adapter caches at most 16 exact font identities. Cancellation runs between bounded synchronous WASM calls, and provider errors stop the export. `dispose()` drops strong font/buffer references and rejects later calls; HarfBuzz 1.6.3 releases native objects through finalizers, so deterministic WASM-heap release requires terminating the host-owned worker.

Logical source clusters are retained in `ToUnicode` and `ActualText`; tagged content allocates its structure references in logical order before visual painting. Independent native extractors can apply their own bidi ordering to raw text. Copy/paste order, screen-reader behavior and PDF/UA acceptance remain separate qualification. TypeScript host assignability is checked with bundler resolution and the engines' own declarations; bidi-js's CommonJS default declaration has a separate NodeNext interop limitation.

`PdfTextShaper` is the engine-independent replacement contract. It accepts exact registered fonts and logical UTF-16 ranges and returns positioned glyphs, without PDF commands or HarfBuzz types. The [owned text-engine roadmap](../Docs/ROADMAP.md#officeimo-owned-text-engines) scopes staged replacements behind that boundary.

### Portable drawings

A `drawing` block paints an ordered collection of paths, fixed-baseline text and rectangular clip groups as one indivisible graphic. It uses the same page fitting, alignment and tagged-figure alternatives as images and vectors. `OfficeIMO.Drawing.OfficeDrawingPortableExporter.TryExportJson` produces this wire format from supported `OfficeDrawing` content. Pass the parsed JSON as the block's `drawing` value; failed producer exports return diagnostics and no partial drawing.

```ts
import { writePdfDocument } from "@evotecit/officeimo";
import type { PdfDrawing } from "@evotecit/officeimo";

const drawing: PdfDrawing = {
  version: 1, width: 200, height: 100,
  elements: [
    { kind: "path", stroke: "555555", lineCap: "square", commands: [
      { op: "move", x: 10, y: 10 }, { op: "line", x: 10, y: 90 },
      { op: "line", x: 190, y: 90 }
    ] },
    { kind: "group", clip: { x: 10, y: 10, width: 180, height: 80 }, elements: [
      { kind: "path", stroke: "217346", strokeWidth: 2, lineJoin: "round", commands: [
        { op: "move", x: 0, y: 95 }, { op: "line", x: 80, y: 60 },
        { op: "line", x: 200, y: 15 }
      ] }
    ] }
  ]
};
const report = await writePdfDocument([
  { kind: "drawing", drawing, alternativeText: "Sales increased through the quarter.",
    caption: { text: "Figure 1. Quarterly sales.", presentation: { italic: true } } }
], { tagged: { language: "en-US" } });
```

Version 1 accepts `path`, `text` and `group` elements only. Paths use the existing `move`, `line`, `cubic` and `close` commands, RGB fills/strokes, positive `strokeWidth`, `fillRule`, `lineCap: "butt" | "round" | "square"`, and `lineJoin: "miter" | "round" | "bevel"`. The miter limit is 10. Groups contain ordered `elements` and an optional rectangular `clip`. Every drawing is also clipped to its own bounds. Coordinates use points from the top left; finite values from -1,000,000 through 1,000,000 allow geometry to extend beyond a clip. Drawing dimensions are positive and at most 14,400 points; rectangular clip dimensions may reach 1,000,000 points.

Text elements declare `text`, `x`, `y`, `font`, `fontSize` and RGB `color`; `y` is the baseline. The font slot is `regular`, `bold`, `italic` or `boldItalic`, and that exact embedded static TrueType face must be supplied in the document's `fonts`. Use the same licensed static family in the .NET producer and the JavaScript writer. Missing faces and synthetic emphasis are rejected. Text stays at its authored position without wrapping or shaping; line breaks, control characters, unsupported scripts and missing glyphs fail explicitly under the existing scalar-text rules.

The wire format rejects unknown properties, versions, commands and element kinds. It does not parse SVG/CSS or load fonts and external resources. `maxDrawingElements` (default 100,000) counts paths, text and groups across drawing blocks; nesting is limited to 32 element levels. Drawing path commands share `maxVectorCommands` with vector blocks. Text shares cell/character budgets with the rest of the document. Serialized drawing content obeys `maxPageBytes`, and preflight and serialization cooperate with cancellation. Keep the supplied drawing stable until its block is consumed.

### Tagged PDF reports

Set `tagged: { language: "en-US" }` on a PDF writer to emit a logical structure tree. The language must be a locale identifier accepted by `Intl.getCanonicalLocales`. Composed paragraphs use `P` by default; `headingLevel: 1` through `6` selects `H1` through `H6` without changing the visual presentation. The document-level `title` uses `H1`; top/bottom messages use `P`.

Sections are `Sect` elements beneath one shared `Document`. Lists use `L` with ordered `LI` children, each containing a marker `Lbl` and a paragraph in `LBody`; a continued item retains that identity across pages. Captions are `Caption` children of their actual `Table` or `Figure`, before the owned content in reading order. Native bookmarks provide navigation destinations separately from these semantic roles.

```ts
import { writePdfDocument, pdfTable } from "@evotecit/officeimo";

const report = await writePdfDocument([
  { kind: "paragraph", text: "Sales report", headingLevel: 1,
    fontSize: 18, presentation: { bold: true } },
  { kind: "paragraph", text: "Services sold this quarter." },
  pdfTable([{ item: "Services", amount: 120 }], {
    columns: [{ header: "Item", key: "item" }, { header: "Amount", key: "amount" }]
  })
], { tagged: { language: "en-US" } });
```

Column headings use `TH` with column scope; empty heading cells use `TD`. Set `rowHeader: true` on a detail column to declare row-scope headings. Every ordinary cell references all declared headings in its row, including an identifier that follows a metric after column reordering. Multiple row headings are supported. This metadata does not select or physically repeat a column, infer roles from a caption, or change CSV/XLSX values and formatting. Column groups and `headerRows` preserve horizontal and vertical spans through `ColSpan`/`RowSpan`. Nonempty column headings have identifiers in the document's structure ID tree. Data and footer cells use explicit `Headers` associations to the applicable headings; a cell with no applicable heading has no such association. Structured `footer.rows` preserve their own spans. A table remains one `Table` with ordered `TR` and cell elements across pages and horizontal bands; each continued cell retains its logical identity. Repeated visual headings, identifiers, continuation captions, table borders/backgrounds, page headers/footers and page numbers are artifacts. Inline and whole-cell links use logical `Link` elements associated with their marked content and annotations across wrapping and page continuation; tagged links require a nonempty displayed label. Interactive links in artifact-only continuation captions have one logical alternative label while every visible copy retains its native annotation association.

`writeGroupedPdf` and composed `pdfGroupedTable` use nested group sections, visible group headings and native outline entries. Blank labels leave the section in place without an outline entry. `groupHeadingLevel` selects the first group's heading level (default 2); child groups increment it. Use 1 for a titleless report or 3 below an authored H2 section. A hierarchy extending beyond H6 is rejected. Each subtotal belongs to its own group and has an explicit row-heading label. A leaf subtotal shares its own detail table and header identities; an ancestor subtotal has a separate table under its ancestor section, and the grand total follows the groups. The last detail and leaf subtotal stay together when the physical page and retained-window limits permit. Labels and continuation text remain caller-authored; the engine never infers semantic roles from language, text prefixes or formatting.

Every image, vector or drawing block in tagged output requires either a nonempty `alternativeText` or `decorative: true`. Meaningful graphics become `Figure` elements with an alternative; decorative graphics are artifacts. Supplying both is an error. Figure alternatives count toward the cell/text budgets and obey `maxCellCharacters`. Choose alternatives and heading levels that express the report's meaning; the writer cannot infer them.

`maxStructureElements` (default 1,000,000) bounds the total logical elements created, including `Document`, sections, lists, captions, tables, rows, cells and figures. `maxMarkedContentReferences` (2,000,000) bounds the total references created for painted content, including continued cell lines. These cumulative limits include completed elements and apply across the document. Untagged output remains the default.

Selected reports were inspected in macOS Preview 11: tables, links and heading hierarchy were exposed, and group outline entries navigated to their destinations. In wide banded samples, the accessibility view exposed separate physical-page tables and fragmented repeated-heading text. These observations do not qualify logical table reading across bands, screen-reader speech or navigation.

This profile provides structure for supported reports. It does not claim accessible PDF or PDF/UA conformance. Assistive-technology behavior and author-supplied semantics require manual acceptance; see the [remaining accessibility work](../Docs/ROADMAP.md#officeimo-javascript).

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

Use `new Workbook(options)` and `addWorksheet(name, options)` for the advanced workbook API. `addWorksheet<RowType>` also accepts shared typed columns/getters and returns a `Worksheet<RowType>` whose `addRows` accepts that domain type. `workbook.worksheets` is a readonly snapshot. Each worksheet exposes its final sanitized `name` and data `rowCount`.

Rows can be synchronous or asynchronous iterables. Arrays are positional. Object rows use each column's literal `key`, or its `header` when no key is supplied; dots never traverse an object. Missing values produce empty cells and extra array values throw. Await each `addRows` call before appending to that worksheet or calling `toBlob`. A failed append prevents finalization. Repeated `toBlob` calls return the same output.

The public `StyleRegistry` owns fonts, fills, borders, number formats and cell-style indexes. `addFont`, `addFill`, `addBorder` and `addNumberFormat` register normalized definitions; `add` combines definitions or existing indexes. Equivalent definitions reuse their indexes. A column's `style` supplies a base, and `format`, `wrapText` and `alignment` override the corresponding fields. A `Cell` can select another registered style for a single value. Indexes belong to that workbook.

Cells accept strings, numbers, booleans, `Date`, `null` and `undefined`. Declared built-in types validate without coercion. Non-finite numbers and invalid dates become empty cells. Strings remain literal text, including `=...` and OOXML `_xHHHH_` sequences. XLSX never interprets a string as a formula.

Dates use Excel's 1900 system, including the fictitious leap day. `dateMode` defaults to local wall-clock fields; `"utc"` writes UTC clock fields. Cell dates carry no timezone. Property dates are UTC instants and are captured at construction. Valid cell date years are 1900 through 9999.

Excel limits are enforced: 1,048,576 rows including headings and the footer, 16,384 columns, 32,767 UTF-16 code units per cell, 65,530 hyperlinks per worksheet including oversized-text preservation links, 64,000 styles and widths from 0 through 255 characters. Names are trimmed, illegal characters become underscores, and names are shortened to 31 UTF-16 code units without splitting a surrogate pair. Blank names become `Sheet`, reserved `History` becomes `History_`, and case-insensitive duplicates receive a suffix.

`invalidCharacterPolicy` defaults to `"strip"`. XML 1.0-invalid controls, lone surrogates, U+FFFE and U+FFFF are removed. `"reject"` throws an `OfficeIMOError` with `code: "INVALID_XML"` for names, cell/header text, property text, font names and number formats. CR, LF, tabs, emoji, Polish and right-to-left text are preserved.

## Streamed output and resource limits

For one table, pass a native stream or `ByteSink` directly. The helper borrows the stream lock until it resolves or rejects; the caller closes or aborts the file and owns any partial bytes:

```ts
import { writeXlsxTo } from "@evotecit/officeimo";

async function exportSales(destination: WritableStream<Uint8Array>, signal: AbortSignal) {
  const result = await writeXlsxTo(loadSales(), destination, {
    columns, signal, dateMode: "utc",
    limits: { maxRows: 1_000_000, maxCells: 8_000_000, maxOutputBytes: 512_000_000 },
    onProgress: progress => console.log(progress.rows, progress.bytes)
  });
  // Close a file destination only after successful completion, using its host API.
  return result;
}
```

`writeCsvTo` and `writePdfTo` have the same destination and completion contract. Pass the signal into paged fetching and destination I/O too; rejecting an export does not cancel application-owned I/O automatically. Native Web Streams need no adapter. For Node writable streams, use the platform's `Writable.toWeb` adapter as shown below.

Select a caller-owned sink when constructing the workbook to deliver ZIP bytes during row production. Complete streamed worksheets in order: `close()` completes a worksheet explicitly, and starting the next worksheet closes its predecessor. Await each append before switching worksheets. A closed worksheet cannot receive more rows. `finish()` completes the ZIP directory and returns data-row, sheet and byte counts; repeated calls reuse the same result. The row count excludes headings, footers and preservation-sheet records. The sheet count includes the preservation sheet when present.

```ts
import { Workbook } from "@evotecit/officeimo/xlsx";
import type { Rows } from "@evotecit/officeimo/core";

async function exportRows(rows: Rows, destination: WritableStream<Uint8Array>) {
  const writer = destination.getWriter();
  try {
    const book = new Workbook({
      sink: { write: bytes => writer.write(bytes) }, dateMode: "utc",
      limits: { maxRows: 1_000_000, maxCells: 8_000_000, maxOutputBytes: 512_000_000 }
    });
    const sheet = book.addWorksheet("Data", {
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

Native XLSX charts add workbook-wide `maxCharts` (100 by default), `maxChartPoints` (100,000 numeric values across all series) and `maxChartTextCharacters` (1,000,000 UTF-16 units across chart metadata and category/series labels). Each chart's hidden source worksheet also counts toward `maxSheets`, and its cells obey the normal row, cell, text and output limits.

XLSX rejects oversized text by default. Set `oversizedText: "preserve"` to write a bounded preview with an internal link to the full value on a `Text overflow` worksheet. Each overflow record identifies the source sheet/cell, one-based part number, text and total part count. Concatenate its `Text` values in part order to reconstruct the complete XML-valid value. Chunks never split a supplementary Unicode character. The same strip/reject policy applies to invalid XML characters. Preserve mode does not replace an existing external link on that cell; that conflict throws.

ZIP entries cannot interleave, so preservation retains a bounded text spool until the report worksheets finish. `maxOverflowCharacters` defaults to 4,000,000 UTF-16 units. Width sampling has separate defaults of 100,000 retained cells and 1,000,000 UTF-16 units of encoded row XML, including markup, configurable through `maxBufferedCells` and `maxBufferedCharacters`. Automatic width sampling stops within these budgets; an explicitly requested sample that exceeds them fails visibly. Full cell values are preserved in both cases. CSV retains long text directly and is an alternative when Excel's cell/storage constraints do not suit the data.

### Node writable streams

`Writable.toWeb` connects Node file or HTTP response streams to `writeXlsxTo`, `writeCsvTo` and `writePdfTo`, including backpressure. End the Node stream after a successful export and await `finished` before using the file. Destroy it on failure; the caller decides whether to remove a partial file.

```ts
import { createWriteStream } from "node:fs";
import { Writable } from "node:stream";
import { finished } from "node:stream/promises";
import { writeXlsxTo } from "@evotecit/officeimo";

async function exportSales(path: string, signal: AbortSignal) {
  const output = createWriteStream(path, { signal });
  const completion = finished(output);
  void completion.catch(() => {}); // Observe I/O failures while the exporter is still writing.
  try {
    const result = await writeXlsxTo([{ customer: "Contoso", amount: 12.5 }], Writable.toWeb(output), {
      columns: [{ header: "Customer", key: "customer" }, { header: "Amount", key: "amount" }], signal
    });
    output.end();
    await completion;
    return result;
  } catch (error) {
    output.destroy();
    await completion.catch(() => {});
    throw error;
  }
}
```

## Workers and paged sources

Use `@evotecit/officeimo/execution` to send projected rows to an application-owned worker. The page keeps its grid and source I/O; the worker calls the normal format writer. `ExportCell` values retain their typed value, display text, presentation and links across the handoff, and Dates remain Dates.

The page creates the worker and owns its lifetime:

```ts
import { executeWorkerExport } from "@evotecit/officeimo/execution";
import { BlobByteSink } from "@evotecit/officeimo/core";
import { saveBlob } from "@evotecit/officeimo";

const worker = new Worker(new URL("./report.worker.js", import.meta.url), { type: "module" });
const controller = new AbortController();
const sink = new BlobByteSink();
try {
  await executeWorkerExport(worker, [["Contoso", 12.5]], sink, {
    signal: controller.signal,
    input: { columns: [{ header: "Customer" }, { header: "Amount", type: "number" }] }
  });
  saveBlob(sink.toBlob("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), "sales.xlsx");
} catch (error) {
  sink.discard();
  throw error;
} finally {
  worker.terminate();
}
```

The worker installs its chosen operation. Keep application callbacks here; functions cannot cross the messaging boundary:

```ts
// report.worker.ts
import { handleWorkerExports } from "@evotecit/officeimo/execution";
import { writeXlsxTo } from "@evotecit/officeimo";
import type { Column, ExportValue } from "@evotecit/officeimo";

const handler = handleWorkerExports(self, ({ input, rows, sink, signal, onProgress }) => {
  const options = input as { columns: readonly Column<readonly ExportValue[]>[] };
  return writeXlsxTo(rows, sink, { columns: options.columns, signal, onProgress });
});
// Call handler.dispose() if this worker stops accepting exports but remains alive.
```

`executeWorkerExport` returns `{ rows, columns, bytes }` and accepts a `ByteSink` or `WritableStream<Uint8Array>`. Use a destination rather than `BlobByteSink` when the completed file should not stay in memory. The worker operation can call CSV or PDF writers in the same way. One endpoint accepts one export at a time and can be reused after completion, failure or cancellation.

The default handoff sends at most 64 rows and 65,536 selected cells per batch, with 1,000,000 UTF-16 units across values, display text, links and presentation strings. `batchRows`, `maxBatchCells` and `maxBatchCharacters` change these ceilings; one row must fit. Cell/text budgets can leave one captured lookahead row on the host. Output sends one chunk of at most 65,536 bytes at a time, controlled by `maxChunkBytes`, and awaits the destination before acknowledging it. Progress coalesces while an acknowledgement is pending and preserves its final update. These are message bounds, not a whole-process memory guarantee.

`input` accepts plain records, arrays, Dates and `Uint8Array` bytes, with default `maxSetupBytes` of 16 MiB, `maxSetupNodes` of 100,000 and nesting up to 32 levels. The byte budget counts UTF-16 string storage and binary bytes. Functions, accessors, cycles, live grid objects and `PdfFont` instances are rejected. Send font bytes when needed and reconstruct them with `await PdfFont.create(bytes, { signal })` inside the operation. Configuration and row/output data are copied through structured cloning; caller-owned buffers are never detached.

Pass an `AbortSignal` to the handoff and the same signal to host-owned page fetching and destination I/O. Cancellation rejects pending source/sink requests, invokes producer cleanup and removes transport listeners. The destination owns partial output. `handler.dispose()` cancels its active operation and removes its listeners; the host still owns worker termination and any object URL it created. Remote OfficeIMO errors retain their `code` and unsupported-feature identifier; a host callback failure or host cancellation preserves its original local reason.

`createDataTablesExport(...).rows` and `createCanopyExport(...).rows()` already produce portable array rows. Supply those rows and their captured columns to the handoff while leaving selection, paging, revisions and source disposal with the owning integration. For an application API, yield pages from an async generator and convert JSON date strings before yielding. Getters that select domain fields run before the worker handoff, or are defined in the worker with its own source. The [worker qualification](../Build/BrowserExports/execution-contracts.js) exercises classic and module workers, backpressure, cancellation and disposal in Chromium, Firefox and WebKit.

The table helpers and DataTables adapter own their workbook. Use portable `ExportCell` values and style definitions in report patches. Workbook-local column/header style IDs, numeric font/fill/border/format references and advanced `Cell` values require the `Workbook` API, where the caller can register those definitions. The convenience writers reject them instead of interpreting indexes from another workbook. Use `XlsxColumn<T>` for advanced columns with a registered style or `Cell` keys/getter results; `Column<T>` is the portable projection shared with CSV. Literal keys select scalar, nullable, date or `ExportCell` fields. Nested objects or arrays need explicit scalar getters for each selected column.

## Resolved values, grouped headings, totals and print layout

`ExportCell` captures a typed value, optional display text and portable presentation before an export begins. Presentation accepts the declared scalar fields only: string colors/number formats, Boolean emphasis/wrapping flags and supported alignment values. Invalid field types are rejected; undeclared fields are omitted. A report producer can resolve highlighting once and reuse the same cells for XLSX, CSV and PDF. XLSX uses the typed value; CSV uses it by default and uses supplied display text with `valueMode: "display"`. PDF prefers supplied display text and applies portable presentation. Formula protection and quoting still run after display-text selection and CSV formatting.

```ts
import { Workbook, ExportCell } from "@evotecit/officeimo/xlsx";
import { writeCsv } from "@evotecit/officeimo/csv";

const columns = [
  { header: "Name", key: "name", groups: ["Identity"], width: 28 },
  { header: "Latency", key: "latency", groups: ["Metrics"], type: "number", format: "0.000" }
] as const;
const rows = [{ name: "Łódź", latency: new ExportCell(125.75, {
  text: "125.750 ms", presentation: { background: "FCE4D6", bold: true }
}) }];
const book = new Workbook();
const sheet = book.addWorksheet("Report", {
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
const regions = book.addWorksheet("Summary", {
  columns: [{ header: "A" }, { header: "B" }], includeHeader: false,
  mergedCells: ["A1:B2", "A3:B3"]
});
await regions.addRows([["Report summary", null], [null, null], ["End", null]]);
```

`sheet.headerRows` replaces inferred group/leaf headings with a rectangular `TableSpanRows` matrix. Each anchor carries a typed `value` and optional `columnSpan`/`rowSpan`; covered positions are `null`. `sheet.footer.rows` uses the same shape after the data and is mutually exclusive with `footer.values`/`footer.totals`. Both matrices are captured before consuming the source. Heading/footer rows do not increase data-row progress or completion counts. `includeHeader: false` suppresses replacement headings too.

```ts
const columns = [{ header: "ID" }, { header: "Asset" }, { header: "State" }, { header: "Score", type: "number" }] as const;
const headerRows = [
  [{ value: "Identity", columnSpan: 2 }, null, { value: "Assessment", columnSpan: 2 }, null],
  columns.map(column => ({ value: column.header }))
];
const footerRows = [
  [{ value: "Source", rowSpan: 2 }, { value: "Operational snapshot", columnSpan: 3 }, null, null],
  [null, { value: "Owner reviewed", columnSpan: 3 }, null, null]
];
const book = new Workbook();
const sheet = book.addWorksheet("Assessment", {
  columns, headerRows, footer: { rows: footerRows }, table: {}, freezeHeader: true
});
await sheet.addRows([["0001", "Controller", "Healthy", 99.3]]);
const workbook = await book.finish();
```

The final heading row above is unmerged and matches the leaf labels, so a native table/filter can use it. Other rectangular headings are supported when table/filtering is disabled. Explicit span matrices retain separate adjacent groups even when their labels are equal.

`footer.values` supplies explicit typed footer cells in column order. `footer.totals` maps unambiguous column keys (or headers without keys) to `sum`, `count`, `average`, `min` or `max`. Aggregation uses visible finite numbers written to Excel, including converted date serials and column-writer results; manually hidden rows and cells containing nested `SUBTOTAL` or `AGGREGATE` functions are excluded. The writer retains only the state needed by each column's operation. Count means numeric count and uses General format rather than a column's date format. Formulas use `SUBTOTAL`, with cached values for readers; averages/minima/maxima remain blank when no visible numeric value exists, including when filtering hides every numeric row. Their guarded expressions are registered as custom table totals so native Excel sees consistent formula metadata. Totals respond to Excel filtering; strings supplied as ordinary values never become formulas. A sum outside JavaScript's finite numeric range fails visibly.

`autoSize` inspects a leading sample of up to 100 rows by default, then starts the worksheet. Automatic sampling shortens the sample to fit the cell and encoded-text buffer budgets, including on wide tables. An explicit `sampleRows` can reach Excel's row capacity and fails if the sample exceeds a buffer budget. Inspecting a complete table requires enough `maxBufferedCells` and `maxBufferedCharacters` capacity for its serialized rows; keep the default leading sample when early streamed output and bounded memory matter. Sampling can span append calls. Each sampled row is validated and serialized once; custom value/style callbacks do not run again when the sample is written. Widths use an approximate character count, preferring `ExportCell.text`, clamped to `minWidth`/`maxWidth`; column `width` wins. This is bounded sizing rather than font measurement.

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
  cellStyle: ({ value, columnIndex }) => columnIndex === 1 && typeof value === "number" && value > 100
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

Style callbacks are synchronous. Their `rowIndex` and `columnIndex` are zero-based data indexes. `worksheetRow` is the one-based Excel row, including titles and headings. Row `values` follow the declared export order and unwrap `Cell` values. A cell callback sees the value after a custom column writer. Headers use `headerStyle` separately. Alternating patches start on the second data row and remain consistent across appends. Heights are points, positive and at most 409; `freezeColumns` freezes leading columns alongside an optional header.

Native tables default unspecified column widths to 20 characters so date/time and numeric columns have useful space. Set `defaultColumnWidth` for another fallback or a column's `width` for an individual override. Worksheets without a table retain Excel's default width unless a width is supplied. These are fixed widths, not font-measured autofit.

Hyperlinks can also be supplied through `SheetOptions.hyperlinks`. Each link addresses one exported cell and uses an absolute HTTP, HTTPS or mailto URL. HTTP credentials, unsafe schemes and duplicate link cells are rejected. Links preserve the existing cell value and style; use a style patch for blue/underlined link text when desired.

`sheet.addImage({ data: pngBytes, row: 6, column: 1, width: 640, height: 320, description: "Latency chart" })` places a PNG, such as a rendered chart, without changing table data. Supply a `Uint8Array` containing the complete PNG; the library checks its signature/IHDR and copies the bytes. Anchors are one-based; display dimensions are CSS pixels at 96 DPI. The writer embeds the image and a one-cell drawing anchor. It does not render charts or convert SVG. Multiple worksheets can each contain a table, links and images.

## Native editable Excel charts

`sheet.addChart` writes a native Excel chart with editable source cells, cached values, labels and axes. Choose `columnClustered`, `barClustered`, `line` or `pie`. PNGs supplied through `addImage` retain their rendered appearance; native charts use Excel's chart drawing and editing. PDF reports can use the [portable drawing contract](#portable-drawings) for already composed graphics.

```ts
import { Workbook, saveBlob } from "@evotecit/officeimo/xlsx";

const rows = [
  { site: "Łódź", actual: 12.5, target: 15 },
  { site: "Warsaw", actual: 25, target: 15 },
  { site: "Gdańsk", actual: 8, target: 15 }
];
const book = new Workbook();
const sheet = book.addWorksheet("Controllers", {
  columns: [
    { header: "Site", key: "site" },
    { header: "Actual", key: "actual", type: "number" },
    { header: "Target", key: "target", type: "number" }
  ],
  table: {}
});
await sheet.addRows(rows);
sheet.addChart({
  type: "columnClustered", title: "Controllers by site",
  data: {
    categories: rows.map(row => row.site),
    series: [
      { name: "Actual", values: rows.map(row => row.actual), color: "2563EB" },
      { name: "Target", values: rows.map(row => row.target), color: "94A3B8" }
    ]
  },
  row: 6, column: 1, width: 640, height: 360,
  legend: "bottom", dataLabels: { showValue: true, position: "outsideEnd" },
  categoryAxis: { title: "Site" },
  valueAxis: { title: "Controllers", minimum: 0, maximum: 30, numberFormat: "0.0" }
});
saveBlob(await book.toBlob(), "controllers.xlsx");
```

A snapshot chart captures its categories, series and options at `addChart`. Later caller mutations cannot change the export. The snapshot form above gets one hidden source worksheet with a workbook-unique name, separate from the report table. Unhide that worksheet in Excel to edit the chart's data; changing the visible report cells does not update this snapshot. Hidden source sheets count in the result's `sheets`; source rows do not count as report `rows` or row progress. Images and native charts can share the same worksheet, in Blob or streamed output.

Categories are strings; every series has the same number of finite numeric values. Series colors accept opaque six-digit RGB hex. Legends support `right` (the default), `bottom`, `left`, `top` or `false`. Column/bar/line charts support category/value axis titles, value bounds and number formats. Pie charts require one nonnegative series with at least one positive value and reject axes. Percent labels are available only for pie charts. Unsupported chart kinds, options and label positions fail explicitly; stacking, secondary axes, trendlines and arbitrary chart styles are outside this contract.

## Editable analysis workbooks

The advanced `Workbook` API writes authored numeric formulas, inline-list editing rules and native charts linked to visible worksheet data. Excel recalculates formulas when the workbook opens or values change; OfficeIMO writes the initial numeric caches supplied by the report producer.

```ts
import { Workbook, FormulaCell, saveBlob } from "@evotecit/officeimo/xlsx";

const statuses = ["Healthy", "Attention", "Offline"] as const;
const counts = { Healthy: 2, Attention: 0, Offline: 1 };
const book = new Workbook();
const services = book.addWorksheet("Services", {
  columns: [{ header: "Status", key: "status" }],
  table: { name: "Services" },
  dataValidation: [{ type: "list", column: "status", items: statuses, allowBlank: false }]
});
await services.addRows([{ status: "Healthy" }, { status: "Offline" }, { status: "Healthy" }]);
await services.close();
const range = services.dataRange("status");

const overview = book.addWorksheet("Overview", {
  columns: [{ header: "Status", key: "status" }, { header: "Services", key: "count", type: "number" }]
});
overview.addChart({
  type: "columnClustered", title: "Service health",
  data: { kind: "range", sheet: overview, rowCount: 3, categoryColumn: "status", series: [{ column: "count" }] },
  row: 6, column: 1, width: 640, height: 340, legend: false
});
await overview.addRows(statuses.map(status => ({ status,
  count: new FormulaCell(range ? `COUNTIF(${range},"${status}")` : "0", counts[status])
})));
saveBlob(await book.toBlob(), "service-analysis.xlsx");
```

Changing a service status in Excel updates the summary and its chart. If Excel uses manual calculation, recalculate to refresh them. The executable [service report job](../OfficeIMO.JavaScript.Examples/report-engine-job.mjs) computes initial caches during the same pass that writes the selected source rows. Its counts cover all exported records; filtering those records later in Excel does not change the `COUNTIF` summaries.

`FormulaCell(formula, cachedValue, style?)` accepts an explicit invariant XLSX storage expression, a finite numeric initial result or `null` for a blank cached result, and an optional workbook-local style index. An optional leading `=` is removed. Expressions contain 1–8,192 UTF-16 units and reject XML-invalid characters even under the text-strip policy. Direct external-workbook and DDE references are unsupported. Use English function names, commas and any required [future-function prefix](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf), for example `_xlfn.AGGREGATE(9,0,A2:A10)`. The writer does not translate localized Excel entry syntax or add function prefixes. The author owns the expression's Excel semantics and cache accuracy; this is formula authoring, not a formula evaluator or sandbox. Ordinary strings remain literal values. Formula cells are supported in advanced worksheet rows, getters, custom-writer results and `footer.values`; portable table/CSV/PDF inputs and structured heading/footer matrices continue to accept resolved values. An explicit formula input cannot also pass through a custom column writer; return the formula from that writer instead.

`worksheet.dataRange(column)` returns a quoted, absolute, finite A1 reference after the worksheet closes, or `undefined` for empty data. Select a unique declared column key or one-based column number. Titles, heading matrices and footers are excluded. The reference describes the exported rows and does not expand when users add rows later.

A chart's `data: { kind: "range", sheet, rowCount, categoryColumn, series }` binds existing chart types to a visible worksheet in the same workbook. Register it while the source is idle, before its first selected row is serialized. Optional `firstRow` is a zero-based data-row offset; column numbers are one-based. Categories require strings and numeric series require finite numbers or numeric formula caches. The writer captures only the declared cache during its ordinary row pass. It emits references to actual source cells and creates no duplicate hidden source sheet. A series name uses the emitted matching leaf heading when available, or a literal supplied name. Bound charts omit hidden/filtered source points in Excel; snapshot charts retain their separate hidden-data behavior. Missing, truncated, foreign or late-bound sources fail explicitly.

Charts omit language metadata by default. Set `language: "pl-PL"` on a chart when its text needs an explicit language tag; valid tags are canonicalized without inferring the machine or report language.

`dataValidation` contains inline-list rules targeting a declared data column. Each rule supplies `type: "list"`, `column` and nonempty `items`. Items cannot contain commas, and the quoted list expression must fit Excel's 255 UTF-16-unit inline-source limit. `allowBlank` defaults to true; `showErrorMessage` defaults to true and `errorStyle` to `"stop"`. Optional `errorTitle` and `errorMessage` fit 32 and 255 UTF-16 units respectively. Rules follow the final exported data range, excluding headings and totals; empty sheets emit no validation elements. Excel validation is an editing aid, not a security boundary.

Workbook limits include `maxFormulaCharacters` (1,000,000 authored UTF-16 units), `maxDataValidations` (1,000 captured rules) and `maxValidationTextCharacters` (1,000,000 captured UTF-16 units). Bound caches share `maxChartPoints` and `maxChartTextCharacters`; formula expressions also consume the general text budget. These limits apply across the workbook, including editing rules declared on empty sources.

## Grouped analytical reports

`createGroupedReport` supplies one shared, single-use analytical stream to XLSX and PDF. It consumes rows already ordered into contiguous groups, with one through four levels from outermost to innermost. Every visible column has a unique stable `key`; translated or reordered headings never choose the grouping field or a metric. A closed group recurring under the same parent is rejected; the same child key under another parent is valid. The engine does not sort or retain all report rows.

```ts
import { createGroupedReport, writeGroupedXlsx, saveBlob } from "@evotecit/officeimo/xlsx";

const columns = [
  { header: "Site", key: "site" },
  { header: "Service", key: "service" },
  { header: "Latency", key: "latency", type: "number", format: "0.00" }
] as const;
const rows = [
  { site: "Berlin", service: "Directory", latency: 12 },
  { site: "Berlin", service: "DNS", latency: 18 },
  { site: "Warsaw", service: "Directory", latency: 6 }
];
const report = createGroupedReport(rows, {
  columns, groupBy: "site", totals: { latency: "average" },
  groupLabel: value => String(value ?? ""),
  limits: { maxRows: 100000, maxCells: 300000 }
});
const file = await writeGroupedXlsx(report, {
  labelColumn: { header: "Summary", width: 30 },
  subtotalLabel: group => `${group.label}: ${group.rows} services`,
  grandTotalLabel: summary => `All ${summary.rows} services`,
  sheet: { name: "Service analysis", freezeHeader: true },
  collapseGroups: true,
  limits: { maxOutputBytes: 96000000 }
});
saveBlob(file, "grouped-services.xlsx");
```

The leading label column preserves every original typed column, including a numeric first column. `sum`, `count`, `average`, `min` and `max` operate on explicitly numeric columns and accept finite numbers or null values. `count` counts numeric cells; `group.rows` and `summary.rows` count all records. Empty sum/count results are zero; empty average/min/max results are blank. Every ancestor and the grand summary aggregate accepted numeric records directly, so unequal child sizes do not distort averages.

XLSX emits native outline controls and finite `SUBTOTAL` formulas with initial caches. Excel recalculates them after detail edits; manually collapsed rows remain included. Grouped sheets do not enable an Excel table or autofilter: filtering/selection belongs to the host capture before export. `includeGrandTotal: false` omits the final summary. Group labels, subtotal labels and grand labels are host-authored strings, so applications supply their own translations and terminology.

For PDF, create a fresh source and call `writeGroupedPdf` or `writeGroupedPdfTo` from the same `/pdf` entry. These convenience writers use the same composed-document owner as `pdfGroupedTable`. The required `labelColumn`, `subtotalLabel` and `grandTotalLabel` remain host-authored; ordinary PDF fonts, layout and display options apply. Continued pages repeat the full active ancestor path and leaf headings. `continuationLabel` supplies display text for each active group. Empty reports retain requested headings with or without a grand total.

Untagged output keeps the last detail row and its closing subtotals together when the bundle fits a fresh page; oversized bundles use ordinary row splitting under repeated context. The [tagged profile](#tagged-pdf-reports) nests group sections and gives nonblank labels a visible heading and navigation entry. A leaf subtotal shares its detail table and column headings; an ancestor subtotal uses a separate summary table under that ancestor's section, and the grand total follows the groups. Each summary row has a row-heading label associated with its values. Separate ancestor and grand-summary tables show the label and declared aggregate columns, including zero or blank results, with their own headings; detail-only fields do not create empty summary bands. The opening path is painted once by group headings and existing ancestor context, then repeated in full on later bands and pages. This presentation can use more space than the compact untagged layout. `groupHeadingLevel` sets the outer group's heading level, defaulting to 2; child groups increment it. Use 1 for a titleless report or 3 beneath an authored H2 section. An opened hierarchy extending beyond H6 fails explicitly.

`writeGroupedXlsxTo` and `writeGroupedPdfTo` await caller-owned destinations. Source and writer cancellation reach pending reads and writes; progress/result `rows` count data records. Generated group and summary rows still consume the format writer's row/cell/text budgets. Source `limits` cover only rows, cells and text; set `maxOutputBytes` on the selected writer. `maxGroups` defaults to 1,000 across all opened nodes, and `maxGroupKeyCharacters` to 1,000,000 retained UTF-16 units. The engine retains parent-scoped identifiers for recurrence detection, depth-bounded aggregate state and one boundary lookahead, rather than group bodies.

Group identity is locale independent: strings compare exactly, numbers and booleans keep their types, dates compare UTC milliseconds, and null/undefined share one empty group. Default date labels use ISO UTC text; use a level descriptor's `label` for localized display, or the single-level `groupLabel` shorthand. Use the factory and writer from the same package entry or standalone bundle. A source is consumed once; create another source, with a new iterable when necessary, for a second format. CSV remains a flat typed-record export.

### Compose a hierarchical analysis

`groupBy: ["region", "site"]` selects projected values by stable column key. For a grouping field that should not appear in the detail columns, supply `{ key, value, label? }`. Its synchronous getter receives the original record and `{ rowIndex, level, key }`; an existing visible column reuses its captured projection and cannot also declare a grouping getter. Each label callback receives the typed key. Levels require unique keys; `groupLabel` applies only to a single level.

```ts
import { Workbook, createGroupedReport, addGroupedXlsx, saveBlob } from "@evotecit/officeimo";

const selected = [
  { portfolio: "Core", environment: "Production", service: "Directory", latency: 12 },
  { portfolio: "Core", environment: "Production", service: "DNS", latency: 18 },
  { portfolio: "Core", environment: "Staging", service: "Directory", latency: 6 }
];
const source = createGroupedReport(selected, {
  columns: [
    { header: "Service", key: "service" },
    { header: "Latency", key: "latency", type: "number", format: "0.00" }
  ],
  groupBy: [
    { key: "portfolio", value: row => row.portfolio },
    { key: "environment", value: row => row.environment }
  ],
  totals: { latency: "average" }
});
const book = new Workbook();
await book.addWorksheet("Context", { columns: [{ header: "Scope" }] })
  .addRows([["Selected services; grouping fields are metadata."]]);
const analysis = await addGroupedXlsx(book, source, {
  labelColumn: { header: "Group", width: 30 },
  subtotalLabel: group => `${group.label}: ${group.rows} services`,
  grandTotalLabel: summary => `All ${summary.rows} selected services`,
  sheet: { name: "Analysis", freezeHeader: true },
  conditionalFormats: [{ type: "cellIs", range: { column: "latency" },
    operator: "greaterThan", value: 15, style: { font: { bold: true, color: "991B1B" } } }]
});
const overview = book.addWorksheet("Overview", {
  columns: [{ header: "Portfolio" }, { header: "Services", type: "number" }]
});
await overview.addRows(analysis.groups.filter(group => group.level === 0)
  .map(group => [group.label, group.rows]));
saveBlob(await book.toBlob(), "selected-services.xlsx");
```

`addGroupedXlsx` closes its new analysis worksheet and leaves the existing workbook open for further sheets and finalization. It shares that workbook's styles, destination, cancellation and cumulative limits. Its result contains `worksheet`, accepted `rows`, output `columns`, `grandSummary`, physical `detailRanges`, per-group identity/path/rows/totals and `startRow`, `firstDetailRow`, `lastDetailRow`, `summaryRow`, plus optional `grandTotalRow`. Coordinates are one-based physical worksheet positions. `Worksheet.rowCount` includes generated analysis rows; workbook result/progress row counts exclude those generated rows and still include ordinary sibling-sheet data. On a grouped worksheet, `dataRange()` returns its single contiguous detail run or `undefined` for no details. Multiple detail runs require the explicit `detailRanges`; a single rectangular range would also include generated totals and is rejected.

Top-level grouped `conditionalFormats` and `dataValidation` apply only to the returned detail runs. Select a stable data-column key or one-based visible data-column number; the added label column does not change those selectors. A1 rule ranges and generated rows are rejected. Inline validation retains the ordinary Excel limits. Disjoint editing references are bounded to 32,767 UTF-16 units; split a worksheet if its group count would exceed that ceiling. Native tables and autofilters remain unsupported for interleaved detail and summary rows.

For a PDF document, create a fresh source and insert `pdfGroupedTable(source, labelsAndTableOptions)` alongside sections, paragraphs, charts, ordinary tables and recommendations. The block inherits fonts, page geometry, cancellation and resource policy from `writePdfDocument`/`writePdfDocumentTo`. Block options contain table presentation and labels; document options belong on the document writer. A grand total moved or split onto a fresh page repeats the metric headings unless `includeHeader: false` was requested. Ownership of a source begins when its block is reached, and a completed, failed or cancelled block releases its iterator and listeners. An unreached block does not consume its source.

Group events expose zero-based `level` and a captured `path` containing all active `{ groupIndex, key, label, level }` identities. `source.groupBy` is a readonly array of stable level keys. A composed report can use a supplied opening overview, or create a computed overview after consuming the detail. It does not silently replay a source or retain its complete body to compute an earlier summary. Totals describe the actual exported rows, including selected/filtered captures; a grid's full-query summary is a different scope.

The [complete analytical report example](../OfficeIMO.JavaScript.Examples/analytical-report-job.mjs) combines hidden grouping fields, detail-only editing rules, an editable overview chart and multi-section PDF output. Its worker input is ordinary flat projected rows: visible values followed by explicitly selected grouping scalars. Those metadata fields consume the same worker cell/text budgets; callbacks are reconstructed in the worker. It has explicit English and Polish example resources, while the engine's keys, grouping and numeric semantics remain language independent.

The [HtmlForgeX operational inventory example](../OfficeIMO.JavaScript.Examples/OperationalInventoryExample.cs) generates a typed table with selection, filtering and column reordering. Its host captures the current view and sends portable values and explicit column metadata to a worker using the public grouped XLSX/PDF writers. Stable service keys identify repeated PDF columns and row headings; the table's presentation remains under the host's control.

## Live Excel conditional formatting

`conditionalFormats` writes native Excel rules that recalculate when users edit workbook values. Use these alongside resolved row/cell highlights when the report needs both an exported appearance and live spreadsheet behavior:

```ts
import { Workbook } from "@evotecit/officeimo/xlsx";

const book = new Workbook({ limits: { maxConditionalFormats: 100, maxDifferentialStyles: 50 } });
const sheet = book.addWorksheet("Latency", {
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
| `cellIs` | `equal`, `notEqual`, `greaterThan`, `greaterThanOrEqual`, `lessThan` or `lessThanOrEqual` with one finite numeric `value`; `equal` and `notEqual` also accept literal text. `between` or `notBetween` with an ordered pair of numeric `values`. Dates can be compared using Excel serial numbers or an expression. |
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

`uploadChunk` above stands for an application's asynchronous byte destination; its promise supplies backpressure. The library performs no network access itself. Both CSV entry points use RFC 4180 quoting, comma/semicolon/tab delimiters, CRLF by default, optional UTF-8 BOM and a final line ending. Alternative endings are LF and CR. Booleans default to `True`/`False`; `booleanText: { true: "Tak", false: "Nie" }` supplies explicit display labels. Valid dates are UTC ISO 8601 with milliseconds. Numbers use JavaScript's invariant string representation.

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

`quote` defaults to `"minimal"`; `"all"` quotes every field and `"strings"` always quotes string values. Required delimiter, quote and newline escaping applies in every mode. `nullValue` replaces null/undefined values and receives the same protection and quoting as other strings. Formatters run on data only; their context contains zero-based data `rowIndex` and `columnIndex`, column definition and original projected `values`. Both Blob and caller-owned sink APIs share these options. CSV retains long text that exceeds Excel's per-cell text limit.

`headerRows` and `footer: { rows }` accept the same rectangular matrices as XLSX. CSV preserves their physical rows and requires `spanMode` whenever an emitted matrix spans cells: `"repeat"` writes the anchor value in every covered field; `"anchors"` writes it once and leaves covered fields empty. `includeHeader: false` suppresses the heading matrix. Display selection, quoting and formula protection still apply; column value formatters run only on data rows. Metadata rows do not increase data-row progress or completion counts.

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

A `ByteSink` implements `write(Uint8Array): void | Promise<void>`. It must accept the bytes before resolving. `BlobByteSink` copies input buffers, so callers can reuse them. `ChunkedTextSink` emits bounded UTF-8 batches and preserves surrogate pairs at append/chunk boundaries. Text buffering and grid projection share a cooperative yield after roughly 32 ms of work, using the browser task scheduler when available. This amortizes task scheduling without blocking for a complete large export. A caller can still supply one large source string or a slow synchronous callback; the budget does not interrupt that synchronous work. `writeBytes` connects a byte iterable to a sink.

Choose CSV for a values-only exchange and XLSX when readers need typed dates/numbers, report layout, highlighting, filtering or formulas. Blob helpers suit downloads whose completed file fits in memory. For large exports, pass a destination to `writeCsvTo` or `writeXlsxTo`, consume a paged async source and await each write; a host-owned worker can keep generation away from the page. The writers yield between batches so page events and cancellation can run. Source getters, formatters and destination callbacks still run in the caller's context and should finish promptly.

Pass an `AbortSignal` and synchronous `onProgress` to the writers. CSV/XLSX progress counts exclude headings and footers. XLSX `rows` is the workbook-wide data total; row events additionally identify `sheetName` and `sheetRows`. DataTables adds its captured `totalRows`. Completion reports the accepted output bytes. Throwing from a callback fails the export. Cancellation returns producer iterators and rejects pending input or sink operations. On failure, writers invoke producer cleanup and observe its result; a pending or rejected cleanup does not delay or replace the export error. A cancelled `finish()` or `toBlob()` releases owned output and rejects its returned promise with the original signal reason, including cancellation between calls. Give I/O producers and destinations the same signal so they can release their own resources.

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

`Workbook.addWorksheet` and `addChart` capture their configuration during registration. Nested registration from an option, column or style getter fails with `OfficeIMOError` code `INVALID_STATE`; prepare dependent inputs before registering them. A rejected nested call does not prevent later registration.

Column writers synchronously convert a domain column type to a supported scalar or styled `Cell`. They receive the column, zero-based `rowIndex`/`columnIndex`, one-based `worksheetRow` and `sheetName`. They cannot inject raw cell XML. Extra parts can carry custom XML, a valid theme or other schema-owned content. An optional internal relationship defaults to `/xl/workbook.xml`; supply `source` for another existing part. Part/relationship definitions are copied when registered; producer iterables remain caller-owned until consumed.

`dataValidation` supports the [inline-list editing profile](#editable-analysis-workbooks). There is no DOCX, PPTX, reader or formula evaluation engine in this package's current writer contract. Grid-specific selection and rendering belong to optional integrations such as the DataTables entry. Row/cell callbacks resolve highlighting during export; `conditionalFormats` separately writes the supported native rules for Excel to evaluate. The [roadmap](../Docs/ROADMAP.md#officeimo-javascript) records the remaining work.

## Portable classic scripts and the .NET asset package

The archive exports `@evotecit/officeimo/classic.js`, `/xlsx.js`, `/csv.js`, `/pdf.js` and optional `/execution.js` for copying into a portable report. Load the chosen asset with a plain script tag:

```html
<script src="officeimo.js"></script>
<script>
  const workbook = new OfficeIMO.xlsx.Workbook();
  const sink = new OfficeIMO.core.BlobByteSink();
</script>
```

The combined bundle exposes all seven namespaces on global `OfficeIMO` plus the established root helpers. The XLSX bundle exposes `core`, `zip`, `xml`, `opc`, `xlsx`; CSV exposes `core`, `csv`; PDF exposes `core`, `pdf`. Load `officeimo-execution.js` beside the chosen writer to add `executeWorkerExport` and `handleWorkerExports` to `OfficeIMO`. The standalone bundles compose in either order. Classic scripts work from `file://` and in host-owned workers without a module loader. ES modules use normal import paths on HTTP(S) and in Node; browser local-file module policy can block `file://` imports.

[OfficeIMO.JavaScript.Assets](../OfficeIMO.JavaScript.Assets/README.md) embeds the same committed TypeScript-built bundles in .NET. Its `JavaScriptAsset` API returns exact content and content-hashed filenames; consuming .NET applications require no Node tooling. Standalone `.mjs` assemblies preserve the asset package's module API. The npm subpaths remain genuine separate `tsc` modules with generated declarations.

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

The browser command uses the repository's existing isolated HtmlTinkerX/Playwright runner and pinned .NET SDK. It checks Chromium, Firefox and WebKit, classic `file://` scripts, ESM imports, workers, downloads and the shared corpus. [C# conformance tests](../OfficeIMO.JavaScript.Assets.Tests/JavaScriptConformanceTests.cs) independently open produced workbooks with OfficeIMO.Excel/OfficeIMO.Reader.Excel and validate them with the Open XML SDK. The [architecture guide](../Docs/officeimo.javascript-architecture.md) describes API stability, format ownership and adding a module.
