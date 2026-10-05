# OfficeIMO.Browser

OfficeIMO.Browser writes value-only tabular XLSX workbooks and UTF-8 CSV in the browser. It consumes the rows and columns supplied by the host, so a portable report can export its current filter, sort and visible column order without a server. Full-dataset workbooks generated with a report remain the job of OfficeIMO.Excel in .NET.

The JavaScript has no runtime dependencies. The .NET asset package embeds the same scripts and exposes their content and SHA-256-based names. It does not execute JavaScript or require a browser on the report-building machine.

```js
import { createWorkbook, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv } from "@evotecit/officeimo/csv";

const columns = [
  { header: "Name", key: "name", width: 28 },
  { header: "Last seen", key: "seen", type: "date", format: "yyyy-mm-dd hh:mm" },
  { header: "Healthy", key: "healthy", type: "boolean" }
];
const rows = [{ name: "DC01", seen: new Date(), healthy: true }];
const workbook = createWorkbook({ creator: "Report", dateMode: "utc" });
const sheet = workbook.addSheet("Domain controllers", { columns, freezeHeader: true, autoFilter: true });
await sheet.addRows(rows);
saveBlob(await workbook.toBlob(), "domain-controllers.xlsx");
saveBlob(await writeCsv(rows, { columns, bom: true }), "domain-controllers.csv");
```

Each export consumes a synchronous or asynchronous iterable once. Arrays are positional; object rows use each column's `key`, or its `header` when no key is declared. Extra array values are rejected; missing values produce empty cells. Await each `addRows()` call before appending again or finalizing. A failed append invalidates that sheet, preventing partial workbooks from being downloaded. Start a new workbook after an error or cancellation.

XLSX cells accept strings, finite numbers, booleans, `Date`, `null` and `undefined`. Declared column types validate values without coercion. `NaN`, infinities and invalid dates become empty cells. Strings stay text, including strings beginning with `=`. XML 1.0-invalid controls, lone surrogates, U+FFFE and U+FFFF are stripped. CR, LF, tabs, emoji and right-to-left text are preserved. Literal OOXML `_xHHHH_` text is escaped for Excel.

Dates use Excel's 1900 date system, including the fictitious leap day between February 28 and March 1. Cell dates default to local wall-clock fields; `dateMode: "utc"` writes UTC clock fields. Neither mode stores a timezone in the workbook. Workbook property dates are UTC instants. Valid cell date years are 1900 through 9999.

Excel permits 1,048,576 rows per sheet, including the header, 16,384 columns, 32,767 UTF-16 code units per cell and widths from 0 through 255 characters. Exceeding a limit throws; the host may split rows into separate sheets. Names are trimmed, illegal characters become underscores, names are limited to 31 UTF-16 code units without splitting a surrogate pair, and case-insensitive duplicates receive a suffix. Blank names become `Sheet`; reserved `History` becomes `History_`. Read `sheet.name` for the final name. Classic ZIP output rejects entries or archives reaching 4 GiB and workbooks above 65,527 sheets.

CSV uses RFC 4180 quoting, CRLF by default, comma/semicolon/tab delimiters, optional UTF-8 BOM and a final line ending. Booleans are `True`/`False`; dates use UTC ISO 8601 with milliseconds. Formula protection defaults to true and follows OfficeIMO.CSV exactly: skip leading ASCII spaces, then prefix a string with an apostrophe when its next character is `=`, `+`, `-`, `@`, tab, CR or LF. Typed negative numbers remain numeric. `formulaInjectionProtection: false` preserves original strings for trusted non-spreadsheet consumers.

Rows are encoded in bounded text batches and the main thread yields during long exports. XLSX streams each sheet through platform `CompressionStream("deflate-raw")` when supported; `compression: "store"` forces uncompressed ZIP, also the automatic fallback on older browsers. Output `Blob` chunks still occupy memory proportional to the final file. There is no shared-string dictionary, persistent storage, network access, formula engine or JavaScript PDF writer.

Pass `signal` and `onProgress` to `createWorkbook` or `writeCsv`. Progress counts exclude headers; XLSX row notifications identify the sheet and completion reports the total. A callback exception fails the export. Cancellation also stops pending asynchronous iterator input; producers waiting on I/O must receive the same signal to release their own resources. Hosts may use their own Blob delivery policy instead of `saveBlob`, especially inside sandboxed iframes.

```csharp
using OfficeIMO.Browser;

string inlineScript = BrowserAssets.Script.Content;
string fileName = BrowserAssets.Script.HashedFileName;
```

For `file://`, use a classic script (`globalThis.OfficeIMO`) or embed `BrowserAssets.Script.Content` in a script element. ES modules can be blocked by local-file origin policy. XLSX, CSV and combined module/classic assets are available individually through `BrowserAssets`. Hashes describe exact UTF-8 bytes without BOM.

Contributor checks use `npm run build`, `npm run check` and `npm test` in this folder. The deterministic asset assembler concatenates the native source; the library build uses no TypeScript compiler, bundler or minifier. Generated files in `Assets` are checked in so building the .NET package requires only the pinned .NET SDK.
