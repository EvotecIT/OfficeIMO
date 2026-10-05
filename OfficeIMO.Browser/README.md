# OfficeIMO.Browser

OfficeIMO.Browser writes value-only tabular XLSX workbooks and UTF-8 CSV in the browser. It consumes the rows and columns supplied by the host, so a portable report can export its current filter, sort and visible column order without a server. Full-dataset workbooks generated with a report remain the job of OfficeIMO.Excel in .NET.

The JavaScript has no runtime dependencies. The .NET asset package embeds the same scripts and exposes their content and SHA-256-based names. It does not execute JavaScript or require a browser on the report-building machine.

## Install and export

From this package's source folder, create an npm archive and install it into the report's JavaScript project:

```sh
npm run build
npm pack --pack-destination /path/to/local-packages
# In the report's JavaScript project:
npm install /path/to/local-packages/evotecit-officeimo-0.1.0.tgz
```

The npm package provides combined, XLSX and CSV ES modules. The entry points are `@evotecit/officeimo`, `@evotecit/officeimo/xlsx` and `@evotecit/officeimo/csv`. Hand-written declarations support strict TypeScript consumers; the library build itself uses plain JavaScript.

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

The host supplies the current view in export order. The writer does not know about a grid's filters, hidden columns, pagination or cell renderers. Supply displayed strings, or supply typed values with an appropriate XLSX number format. Use a fresh iterable for each export when the input is a one-shot generator.

## XLSX options and limits

| Scope | Options |
| --- | --- |
| Workbook | `creator`, `title`, `created`, `modified`, `dateMode`, `compression`, `signal`, `onProgress` |
| Sheet | `columns`, `includeHeader` (default true), `freezeHeader`, `autoFilter`, `boldHeader` (default true), `headerFill` (RGB or ARGB hex) |
| Column | `header`, `key`, `width`, `type`, `format`, `wrapText`, `alignment` |

Horizontal alignment accepts `left`, `center`, `right`, `fill`, `justify` and `distributed`. Header fills are optional. Multiple sheets share the workbook's style registry, but each sheet has its own declared projection and append lifecycle.

XLSX cells accept strings, finite numbers, booleans, `Date`, `null` and `undefined`. Declared column types validate values without coercion. `NaN`, infinities and invalid dates become empty cells. Strings stay text, including strings beginning with `=`. XML 1.0-invalid controls, lone surrogates, U+FFFE and U+FFFF are stripped. CR, LF, tabs, emoji and right-to-left text are preserved. Literal OOXML `_xHHHH_` sequences are split across inline text runs so Excel and readers that concatenate the runs preserve the same literal text.

Dates use Excel's 1900 date system, including the fictitious leap day between February 28 and March 1. Cell dates default to local wall-clock fields; `dateMode: "utc"` writes UTC clock fields. Neither mode stores a timezone in the workbook. Workbook property dates are UTC instants. Valid cell date years are 1900 through 9999.

Excel permits 1,048,576 rows per sheet, including the header, 16,384 columns, 32,767 UTF-16 code units per cell and widths from 0 through 255 characters. Exceeding a limit throws; the host may split rows into separate sheets. Names are trimmed, illegal characters become underscores, names are limited to 31 UTF-16 code units without splitting a surrogate pair, and case-insensitive duplicates receive a suffix. Blank names become `Sheet`; reserved `History` becomes `History_`. Read `sheet.name` for the final name. Classic ZIP output rejects entries or archives reaching 4 GiB and workbooks above 65,527 sheets.

## CSV

CSV uses RFC 4180 quoting, CRLF by default, comma/semicolon/tab delimiters, optional UTF-8 BOM and a final line ending. Booleans are `True`/`False`; dates use UTC ISO 8601 with milliseconds. Formula protection defaults to true and follows OfficeIMO.CSV exactly: skip leading ASCII spaces, then prefix a string with an apostrophe when its next character is `=`, `+`, `-`, `@`, tab, CR or LF. Typed negative numbers remain numeric. `formulaInjectionProtection: false` preserves original strings for trusted non-spreadsheet consumers.

`writeCsv` requires `columns` and accepts `delimiter`, `lineEnding`, `bom`, `includeHeader`, `formulaInjectionProtection`, `signal` and `onProgress`. Alternative line endings are LF and CR. Column `key` and `header` select and label data; XLSX-only widths, styles and declared types do not change CSV formatting. The shared [OfficeIMO.CSV vectors](../OfficeIMO.TestAssets/CSV/browser-exports.json) qualify identical .NET/JavaScript bytes for quoting, delimiters, BOM and formula protection.

## Streaming, progress and cancellation

Rows are encoded into text batches targeting 32 KiB, and the main thread yields after about 8 ms of work. A single CSV field can exceed the batch target. XLSX compresses each sheet during `addRows()` through platform `CompressionStream("deflate-raw")` when supported; `compression: "store"` forces uncompressed ZIP, also the automatic fallback when the platform lacks raw deflate. Output `Blob` chunks still occupy memory proportional to the final file, with browser-dependent native buffering. Stored fallback files can be much larger than compressed output.

Strings use `inlineStr`. In the reproducible 10,000-row × 4-column XML-payload comparison, a shared-string representation saves 1.6% with 80 distinct strings, but increases compressed payload by 56.9% with 40,000 unique strings. Inline strings avoid a growing dictionary and let platform compression handle repeated text. The comparison uses the same deflater for both representations and excludes ZIP metadata; run [the measurement tool](../Build/BrowserExports/measure.mjs) to reproduce it.

There is no persistent storage, network access, formula engine or JavaScript PDF writer. Hosts can run the same classic asset in their own Blob worker and return output through `postMessage`; download helpers belong on the page. This avoids requiring module loading from a local-file origin. Worker creation still depends on the host's Content Security Policy.

Pass `signal` and `onProgress` to `createWorkbook` or `writeCsv`. Progress counts exclude headers; XLSX row notifications identify the sheet and completion reports the total. A callback exception fails the export. Cancellation also stops pending asynchronous iterator input; producers waiting on I/O must receive the same signal to release their own resources. Hosts may use their own Blob delivery policy instead of `saveBlob`, especially inside sandboxed iframes.

## Embedded assets and offline reports

Reference `OfficeIMO.Browser.csproj` from a source checkout, or pack the .NET assets into a local NuGet feed and install that archive:

```sh
dotnet pack OfficeIMO.Browser/OfficeIMO.Browser.csproj -c Release -o ./packages
dotnet add /path/to/report/Report.csproj package OfficeIMO.Browser --version 3.4.4 --source ./packages
```

The package targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2. Generated scripts are embedded resources; consuming applications need neither Node nor a JavaScript build tool.

```csharp
using OfficeIMO.Browser;

string inlineScript = BrowserAssets.Script.Content;
string fileName = BrowserAssets.Script.HashedFileName;
```

For `file://`, use a classic script (`globalThis.OfficeIMO`) or embed `BrowserAssets.Script.Content` in a script element. ES modules can be blocked by local-file origin policy. XLSX, CSV and combined module/classic assets are available individually through `BrowserAssets`. Hashes describe exact UTF-8 bytes without BOM.

Each `BrowserAsset` exposes `FileName`, `Content`, `ContentHash` (the first 16 lowercase hex digits of SHA-256) and `HashedFileName`. Write `Content` as UTF-8 without BOM when delivering the hashed filename. The classic npm assets are also exported as `@evotecit/officeimo/classic.js`, `@evotecit/officeimo/xlsx.js` and `@evotecit/officeimo/csv.js` for copying into an offline report.

The [minimal HtmlForgeX example](../OfficeIMO.Browser.Examples/README.md) generates a single HTML file and an HTML/JavaScript bundle. Both export the filtered, sorted, reordered and visible table from `file://` with no server. Host sandbox and download policies still apply; SharePoint/UltimateHtmlViewer and grid-specific wiring require qualification in those hosts.

Excel and OfficeIMO's .NET reader preserve the qualified date serials and cell text. In the LibreOffice 24.2.7 spot check, dates before March 1900 display one day earlier, and conversion to CSV normalizes a cell's CRLF to LF. The writer retains Excel's standard serial dates and the original XML text rather than changing them for that application behavior.

## Contributing

Use `npm run build`, `npm run check` and `npm test` in this folder. The deterministic asset assembler concatenates native source; the library build uses no TypeScript compiler, bundler or minifier. Generated files in `Assets` are checked in with LF line endings so building the .NET package requires only the pinned .NET SDK.

[Browser export verification](../Build/BrowserExports/README.md) runs all three browser engines, the .NET reader, Open XML validation, strict declarations and isolated npm/NuGet consumers. Scale and size measurements are explicit opt-in evidence runs, separate from CI correctness checks. Browser and HTML-generation tools are confined to the verification/example projects and do not enter either distributed package.
