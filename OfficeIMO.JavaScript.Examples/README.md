# Offline current-view exports

This minimal HtmlForgeX example generates a single HTML file with the OfficeIMO runtime embedded, plus an HTML/JavaScript bundle using the .NET asset's content-hashed filename. Open either page directly from disk. It needs no server or internet connection.

From the repository root:

```sh
dotnet run --project OfficeIMO.JavaScript.Examples/OfficeIMO.JavaScript.Examples.csproj -c Release -- /path/to/example-output
```

Open `current-view.html` or `current-view-bundle.html` in that output directory. Filter the three controllers, reverse their name order, hide Site or move it to the first/last column, then export Excel or CSV. Exports contain the current displayed rows and visible columns in their current order. Excel uses a native table, a blue header, frozen row/column, numeric latency values with two decimal places and warning highlighting above 20 ms. CSV uses a UTF-8 BOM and the package's default formula protection. An empty view exports just the visible headers.

The page uses HtmlForgeX's existing default styles and a plain typed table. Its small script projects the displayed DOM using stable column fields, so translated headings do not change data types, sorting or formatting; a data-grid integration should supply the grid's own current-view iterable, especially when the grid virtualizes rows. The [OfficeIMO JavaScript guide](../OfficeIMO.JavaScript/README.md) owns the writer API and limits.

This example stays outside the default solution. HtmlForgeX is an example-only dependency and is absent from the OfficeIMO.JavaScript.Assets NuGet and npm artifacts.

## Grouped service reports

`report-engine-job.mjs` is the shared worker job for the website's report demonstration. Its `writeOperationalReport` entry accepts `layout: "analysis"` or `"grouped"`. `operationalRowsBySite` supplies an explicitly site-ordered lazy source for grouping; a real grid host supplies its own filtered/selected capture ordered by the selected grouping field.

Grouped mode writes XLSX outline controls and editable numeric subtotals, or PDF group sections with continued-page context. It preserves typed raw values and derives metrics from the entire exported capture. The example's explicit `language: "en"` and `"pl"` resource profiles own headings, summary labels, page counters and display formatting; stable column keys choose all semantic roles. CSV retains the flat analysis layout.

The grouped source permits 100,000 data rows, 700,000 cells and 1,000 groups. Writer limits also cover generated summaries and cap output at 96 MiB. PDF uses host-loaded, embedding-permitted Carlito regular/bold font bytes. Grouping, metrics, formulas, pagination and byte writing remain in the OfficeIMO engine; the job supplies data, labels and report appearance.

## Wide analytical reports

`wide-analytical-report-job.mjs` extends the hierarchical service example to 16 typed fields, including request volume, error rate, CPU, memory, storage, certificate expiry, version and incident dates. `wideAnalyticalRows(count)` produces the ordered data lazily; `writeWideAnalyticalReport` accepts the same worker context and explicit English/Polish resources as the analytical example. XLSX retains all fields in one editable detail sheet. PDF selects the engine's horizontal-band policy with repeated service IDs and ancestor context. The example owns no band planner, source replay or PDF renderer. Its finite physical-work limits also cover repeated output.
