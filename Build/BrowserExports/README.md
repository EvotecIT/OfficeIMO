# Browser export verification

## CanopyX captures

The optional CanopyX adapter is qualified against an explicitly selected native source checkout. It introduces no CanopyX package into normal builds or shipped artifacts:

```powershell
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/canopy" "--canopy=$canopySource"
# Opt-in native paging: 10,000/100,000 rows, four/twenty columns; --full also adds 250,000/1,000,000 rows at four columns.
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/canopy-scale" "--canopy=$canopySource" --canopy-scale --full
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run test:canopy-types -- "$evidence/npm" "$canopySource"
```

The normal lane reads actual ordinary and reporting assets, records their hashes, and rejects a candidate that changes during the run. Each browser qualifies raw/display values, column order and visibility, filtered and selected-query captures after UI changes, unknown-count cursors, revision mismatch, in-flight source cancellation, the prepared host callback capture, a clicked export button and explicit diagnostics for custom renderers and relative links. Classic, compression-fallback and bounded worker paths produce independently read CSV/XLSX/PDF artifacts. XLSX validation checks schema, typed values, formats, tone precedence and hyperlinks; PDF validation checks ordered resolved text and link actions. Wide and compact screenshots complement the assertions.

The opt-in scale lane generates native source pages on demand, checks every CSV/XLSX cell using the shared independent oracle, and awaits a 64 KiB file bridge. Ten-thousand-row cases use offset pages, a delayed sink and compression fallback; larger cases exercise unknown-count cursors and offset pages. Large completed files are removed after readback, retaining compact reports. Native source ID/cursor bookkeeping still grows with the record count. Diagnostic durations include native capture and acknowledged file delivery; they are not controlled performance comparisons. PDF scale and full report preservation use the separate format qualification lane.

PDF table qualification uses the same HtmlTinkerX browser owner and a separate managed PDF reader:

```powershell
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/pdf" --pdf "--comparison-assets=$assets"
# Use --stack=bundled for the older installed DataTables/Buttons pair.
# Opt-in scale adds 10,000/100,000-row plain/styled tables, four/twenty columns.
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/pdf-scale" --pdf --scale
```

The PDF lane checks Chromium, Firefox and WebKit, native/uncompressed/fallback streams, Polish text, Japanese text, supplementary symbols, repeated and spanned headings, multirow footers, totals, oversized text, paged input, slow sinks, cancellation and resource ceilings. Worker output transfers byte chunks to avoid WebKit worker Blob-read restrictions. Every document must open without repair in `OfficeIMO.Pdf`; ordered row IDs, complete long text, headings and expected values are independently checked. DataTables button scenarios run when verified comparison assets are provided. The ordinary workflow runs both installed pairs without host-timing gates. `--scale` records diagnostic durations and validates every cell in plain/styled tables at 10,000 and 100,000 rows with four/twenty columns; large PDFs are deleted immediately after readback. Rendered-page inspection with an independent PDF application remains separate visual proof.

The verification runner uses the repository's pinned .NET SDK, Node 18 or newer and PowerShell 7 for the wrapper commands. Browser installation, sessions and captures use the existing test-only HtmlTinkerX/Playwright package. The runner is outside the normal solution and is not packable; neither distributed OfficeIMO package acquires browser tooling. Excel and LibreOffice spot checks remain optional independent application proof.

The `--links` correctness lane qualifies portable cell links in all three browser engines. It checks numeric/date XLSX values and relationship coordinates through both managed readers and the Open XML SDK, then reads PDF URI actions, Unicode tooltips and page-fragment rectangles independently. Classic and compression-fallback lanes produce Blobs; the worker lane writes to a byte destination and transfers chunks to avoid host-specific worker Blob read limitations. Each lane also checks raw/display CSV semantics. Use `--validate-links <fixture-directory>` to read the captured small link fixtures without launching a browser.

Build and run the TypeScript package before the C# tests:

```powershell
$evidence = '/path/to/task-evidence'
npm ci --ignore-scripts --prefix OfficeIMO.JavaScript
npm --prefix OfficeIMO.JavaScript --script-shell pwsh test
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run test:consumer
dotnet test OfficeIMO.Browser.Tests/OfficeIMO.Browser.Tests.csproj -c Release
# The CSV contract is also part of the ordinary CSV suite.
dotnet test OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj -c Release --filter FullyQualifiedName~CsvBrowserExportVectorsTests
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run test:browser -- "$evidence/browser"
./Build/BrowserExports/Test-NpmConsumer.ps1 -EvidenceDirectory "$evidence/npm"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- --validate-directory "$evidence/npm"
./Build/BrowserExports/Test-NugetConsumer.ps1 -EvidenceDirectory "$evidence/nuget"
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run sizes -- "$evidence/sizes.json"
```

`npm test` runs the strict compilation, committed-bundle check and Node tests. The separate consumer check imports every public subpath with strict compiler options and positive/negative type contracts. The npm consumer installs the real archive into an isolated application, uses the package's pinned test compiler, executes all subpaths and checks that no runtime dependencies were installed. The NuGet consumer restores the actual asset archive into a task-local package directory and verifies all ten byte/hash contracts on .NET 8 and .NET 10.

The shared XLSX manifest is `OfficeIMO.TestAssets/JavaScript/xlsx-writer.json`; CSV vectors are in `OfficeIMO.TestAssets/CSV/browser-exports.json`. A C# test invokes the Node fixture producer and validates every generated file using the shared `JavaScriptWorkbookContract`. Both OfficeIMO.Excel and OfficeIMO.Reader.Excel open the workbooks, expected cells/styles/parts are checked, and the Open XML SDK validates each document. The CSV contract compares TypeScript bytes with independently generated OfficeIMO.CSV bytes.

The report PNG fixture comes from the published MIT-licensed ChartForgeX 1.8.2 package: `Chart.Create().WithTitle("Report latency").WithSize(420, 240).WithTheme(ChartTheme.ReportLight()).WithXAxis("Sample").WithYAxis("Milliseconds").AddBar("Latency", new[] { new ChartPoint(1, 12), new ChartPoint(2, 8), new ChartPoint(3, 15) }).ToPng()`. The manifest records its producer and SHA-256. Verification consumes those fixed bytes; it does not install ChartForgeX. The browser report lane places two native report tables, free-form merged regions and the chart worksheet in one workbook, in both Blob and streamed output and both compression modes. Independent ZIP traversal compares the embedded PNG with the producer bytes.

The browser command owns a temporary server bound to `127.0.0.1` for the actual ESM graph. It waits for its child verification process and closes the server in `finally`. Classic scripts and report downloads are tested through `file://`. Chromium, Firefox and WebKit cover public layers, shared fixtures in both compression modes, escaping, worker exports, cancellation, limits and the two existing offline examples. Append `--engine=WebKit` for a focused browser reproduction. The command's normal lane always runs all three engines and the row-limit checks.

Every browser-produced XLSX, including worker and example downloads, is opened and SDK-validated. Every positive Node test artifact can also be captured and qualified:

```powershell
$env:OFFICEIMO_JS_TEST_OUTPUT = "$evidence/node-workbooks"
npm --prefix OfficeIMO.JavaScript --script-shell pwsh test
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- --validate-directory "$evidence/node-workbooks"
```

`OFFICEIMO_JS_EVIDENCE_DIR` optionally retains the ordinary C# test's generated corpus beneath a caller-owned evidence directory. Without it, the test uses and deletes a unique temporary directory. Browser reports record engine/version, layer assertions, workbooks, CSV counts, worker/limit coverage and runtime errors. Example screenshots capture wide/compact states; screenshots complement the download assertions.

Host-dependent measurements stay explicit. Comparison runs require the native baseline and at least one OfficeIMO lane; select `-Qualification` for OfficeIMO-only measurements:

```powershell
node Build/BrowserExports/measure.mjs "$evidence/measurements"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/scale" --scale
```

The measurement script compares actual inline-string worksheet XML with an equivalent shared-string representation using the same compressor, and measures the generated bundles. The scale runner writes 100,000 rows by 20 columns for XLSX/CSV, validates output and records responsiveness and available heap metrics. These timing/memory measurements do not become ordinary CI correctness envelopes.

`Check-Excel.ps1` and `check-libreoffice.py` preserve optional application spot checks. They do not install, ship or invoke either application as a product dependency.

## DataTables integration and comparisons

The optional grid integration is qualified against the version pairs in `comparison-assets.json`. The downloaded MIT-licensed assets are test-only and hash-checked; they never enter npm, NuGet or product build dependencies. The `bundled` pair identifies the assessed HtmlForgeX asset versions, while `current` identifies the pinned current comparison pair. Updating this manifest is an intentional comparison-input change.

```powershell
node Build/BrowserExports/fetch-comparison.mjs "$evidence/comparison-assets"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/datatables" --datatables "--comparison-assets=$evidence/comparison-assets"
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run test:datatables-types -- "$evidence/npm"
```

The interoperability lane covers actual search/order/selected-row defaults, hidden and reordered columns, grouped headings, footers, presentation, button delivery/completion, server-side loaded scope, cancellation and bounded host-owned workers in Chromium, Firefox and WebKit. Workers request at most 64 portable rows at a time and acknowledge output chunks of at most 64 KiB with a delayed sink; auto/stored compression and cancellation run separately. Every completed workbook is opened by both C# readers and the Open XML validator. Wide/compact captures and runtime errors are retained. Type consumers install the actual npm archive and upstream declarations in isolated folders; newer upstream declaration errors are recorded separately from strict consumer assignability.

The comparison suite uses PSPublishModule/PowerForge for warmups, rotated ordering, measurements and JSON/CSV/Markdown artifacts. Build the runner first and supply its DLL. Pass an affinity mask for a discovered processor domain; repeat the complete selected matrix on each domain when interpreting a heterogeneous host. Setup, grid creation, output transfer and validation are excluded. Each measured operation completes one Blob export. The inner `BrowserExportMs` metric excludes the small control roundtrip. `GatherMs` measures calls to Buttons' `exportData()` only: batched mode gathers headings there and includes its later body projection in `BrowserExportMs`, so `GatherMs` is not a comparable measure of all data collection. Host process allocation/working set metrics describe PowerShell, not the browser.

```powershell
./Build/BrowserExports/Run-DataTablesComparison.ps1 -Binary /path/to/OfficeIMO.Browser.Interop.dll -Assets "$evidence/comparison-assets" -OutputRoot "$evidence/comparison" -Rows 10000 -Columns 20,100 -Plan
./Build/BrowserExports/Run-DataTablesComparison.ps1 -Binary /path/to/OfficeIMO.Browser.Interop.dll -Assets "$evidence/comparison-assets" -OutputRoot "$evidence/comparison" -Rows 10000,100000 -Columns 20 -WarmupCount 1 -IterationCount 3
# XLSX qualification produces diagnostic timings and no cross-library ratios.
./Build/BrowserExports/Run-DataTablesComparison.ps1 -Binary /path/to/OfficeIMO.Browser.Interop.dll -Assets "$evidence/comparison-assets" -OutputRoot "$evidence/xlsx-qualification" -Rows 10000,100000 -Columns 20 -Formats xlsx -Lanes compatibility,batched -Qualification
# A full-table width scan is required for comparative XLSX measurements.
./Build/BrowserExports/Run-DataTablesComparison.ps1 -Binary /path/to/OfficeIMO.Browser.Interop.dll -Assets "$evidence/comparison-assets" -OutputRoot "$evidence/xlsx-comparison" -Rows 1000,10000 -Columns 4,20 -Formats xlsx -Lanes native,batched -FullWidthScan
./Build/BrowserExports/Run-DataTablesComparison.ps1 -Binary /path/to/OfficeIMO.Browser.Interop.dll -Assets "$evidence/comparison-assets" -OutputRoot "$evidence/pdf-comparison" -Rows 1000,10000 -Columns 4,20 -Formats pdf -Lanes native,batched
```

Cross-library comparisons accept CSV, PDF and XLSX with an explicit `-FullWidthScan`. Native Buttons, OfficeIMO compatibility and OfficeIMO batched lanes use the same grid and ordered values: typed numbers and repeated literal Unicode strings, no title/footer and quoted CSV. `-Unique` selects unique strings. PDF comparisons use the same supplied Carlito fonts, A3 landscape pages, font size, column widths, grid and repeated headings. `-Styled` adds alternating body bands to PDF. Completion waits for the actual PDF Blob; validation reconstructs wrapped cells across lines and pages and checks every value, heading, font size and printable boundary. PDF comparison workloads support up to twenty readable columns.

Native Excel scans the complete table for approximate column widths. Comparative XLSX therefore asks OfficeIMO to sample every data row and explicitly raises its cell and serialized-character buffer budgets for that workload. Both exporters clamp widths to 54; their approximate sizing algorithms can produce different widths. Validation checks full-table sizing and readable bounds. This opt-in comparison retains the serialized sample in memory and does not measure OfficeIMO's default bounded leading sample or streamed memory behavior. Fixed-width XLSX is available through `-Qualification` with explicit OfficeIMO lanes and no comparison ratios; it validates width 20, no filter, every value, cell type, coordinate and count. XLSX styles are SDK-validated at every size; workbooks up to 10,000 rows additionally use both C# readers and the full SDK validator when styles conform, with an explicit 128 MiB per-part XML character budget for the bounded 100-column workload. Larger worksheets use forward-only validation rather than a whole-document DOM.

The default `-TextProfile unicode` includes supplementary emoji. `-TextProfile bmp` keeps accented Unicode text without supplementary characters and records that distinct input contract in the metadata. Run both when assessing interoperability; an invalid output in one profile remains a failure, even when the other profile passes. PDF uses text supported by the supplied fonts in either profile.

The runner retains each operation's measurement, content hash and validation result before deleting the large output. Any conformance failure leaves its lane failed and the wrapper exits with an error after collecting all cases. A native exporter may preserve values while failing the schema check: its retained timings are diagnostic, not a qualified speed ranking. Do not suppress that failure or repair the baseline's generated document to improve its status. Responsiveness is the largest sampled timer gap; sampled Chromium JavaScript heap excludes native/browser-process memory and is unavailable in Firefox/WebKit. Keep environment, source/binary/asset hashes, affinity, browser versions and failed lanes alongside any interpretation.

For bottleneck investigation, the JSON-line `--datatables-session` protocol accepts `"profile": true` on a `run` request after `prepare`. Chromium writes the native DevTools CPU profile to `export.cpuprofile` in the session's evidence directory and detaches the profiling session before returning. Deliver and validate that export through the usual `validate` request. Profiling perturbs timing; keep these captures separate from controlled comparison runs and copy a profile before another capture replaces it.

`"diagnosticYields": true` instruments fallback timer delays and public cell-render/index/text-stripping calls in the same isolated browser session. The profile reports native task-scheduler availability. Instrumentation perturbs timing and is excluded from controlled rankings. Browser globals and API methods are restored after the operation, and output follows the same independent validation route.

## Representative export qualification

The opt-in matrix crosses plain/styled XLSX and CSV at 10,000, 100,000, 250,000 and 1,000,000 rows with four/twenty columns and repeated/unique strings. It exercises the shared table writers and nested-object column getters, checking one projection per selected cell, completion counts, typed numeric/boolean/date values throughout and long Unicode in the 10,000-row matrix. Sources fetch and consume bounded 256-row pages. Additional 10,000-row cases exercise delayed page delivery, cancellation while a second page fetch is pending, workers, stored compression fallback, slow sinks, a hung sink cancelled through the export signal, source cancellation and row ceilings. Worker and worker/fallback cases also export 100,000 rows by twenty columns. Every case runs in Chromium, Firefox and WebKit; execution modes are these targeted lanes, rather than a full cross-product with every matrix size. `--columns=4` or `--columns=20` narrows the main matrix; targeted cases keep their stated widths.

```powershell
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/qualification" --qualify --full
# A smaller reproduction; --keep retains artifacts deliberately.
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/smoke" --qualify --rows=10000 --keep
```

The caller-owned sink sends at most 64 KiB to the host per delivery. Each artifact is independently checked before removal: CSV uses `OfficeIMO.CSV.OpenDataReader`; XLSX uses forward-only ZIP/XML traversal to verify every typed cell, coordinate, count and cached total. Preservation chunks reconstruct the complete source text. The 10,000-row XLSX cases additionally run the Open XML SDK validator and both OfficeIMO readers; the rich adapter uses an explicit 64 MiB per-part XML budget. Larger cases avoid whole-worksheet DOM validation and report that distinction. Failure cases must leave an unfinished archive, with partial bytes owned by the destination. Reports retain hashes, browser versions, validation results, output bytes, first-byte/source counts, timer gaps and heap samples where the engine exposes them.

Artifacts are deleted after successful validation unless `--keep` is selected; compact JSON reports remain. `--matrix-only` omits the extra failure/worker cases. These are qualification runs with diagnostic durations, not controlled cross-library speed rankings. CPU placement, warmup and rotated comparison policy belong to the shared PowerForge benchmark runner; sampled Chromium JS heap excludes native compression/browser allocations, and missing Firefox/WebKit heap metrics are recorded as unavailable. No measurement becomes an ordinary CI correctness threshold.

Add `--conditional` to qualify live XLSX conditional formats across the same row sizes and four/twenty-column plain/styled layouts. This lane selects XLSX with repeated strings in the main matrix and retains unique-string worker/paged cases from the targeted lanes. Every completed workbook must contain exactly four rules over its final data rows and two differential styles, independent of row count; the streaming validator checks their types, priorities, range boundaries and bounded XML size alongside every typed value. Slow sinks, compression fallback and cancellation use the same rules. The shared `conditional-report` corpus separately qualifies two/three-color scales, formula thresholds, explicit font resets, partial borders, format overrides, empty data ranges and report layout through the SDK and readers. Native Excel edit/recalculate/render evidence remains a separate application check.

```powershell
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/conditional" --qualify --full --conditional
```
