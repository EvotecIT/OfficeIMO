# Browser export verification

The verification runner uses the repository's pinned .NET SDK, Node 18 or newer and PowerShell 7 for the wrapper commands. Browser installation, sessions and captures use the existing test-only HtmlTinkerX/Playwright package. The runner is outside the normal solution and is not packable; neither distributed OfficeIMO package acquires browser tooling. Excel and LibreOffice spot checks remain optional independent application proof.

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
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- --validate "$evidence/npm/packed-consumer.xlsx"
./Build/BrowserExports/Test-NugetConsumer.ps1 -EvidenceDirectory "$evidence/nuget"
npm --prefix OfficeIMO.JavaScript --script-shell pwsh run sizes -- "$evidence/sizes.json"
```

`npm test` runs the strict compilation, committed-bundle check and Node tests. The separate consumer check imports every public subpath with strict compiler options and positive/negative type contracts. The npm consumer installs the real archive into an isolated application, uses the package's pinned test compiler, executes all subpaths and checks that no runtime dependencies were installed. The NuGet consumer restores the actual asset archive into a task-local package directory and verifies all six byte/hash contracts on .NET 8 and .NET 10.

The shared XLSX manifest is `OfficeIMO.TestAssets/JavaScript/xlsx-writer.json`; CSV vectors are in `OfficeIMO.TestAssets/CSV/browser-exports.json`. A C# test invokes the Node fixture producer and validates every generated file using the shared `JavaScriptWorkbookContract`. Both OfficeIMO.Excel and OfficeIMO.Reader.Excel open the workbooks, expected cells/styles/parts are checked, and the Open XML SDK validates each document. The CSV contract compares TypeScript bytes with independently generated OfficeIMO.CSV bytes.

The browser command owns a temporary server bound to `127.0.0.1` for the actual ESM graph. It waits for its child verification process and closes the server in `finally`. Classic scripts and report downloads are tested through `file://`. Chromium, Firefox and WebKit cover public layers, shared fixtures in both compression modes, escaping, worker exports, cancellation, limits and the two existing offline examples. Append `--engine=WebKit` for a focused browser reproduction. The command's normal lane always runs all three engines and the row-limit checks.

Every browser-produced XLSX, including worker and example downloads, is opened and SDK-validated. Every positive Node test artifact can also be captured and qualified:

```powershell
$env:OFFICEIMO_JS_TEST_OUTPUT = "$evidence/node-workbooks"
npm --prefix OfficeIMO.JavaScript --script-shell pwsh test
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- --validate-directory "$evidence/node-workbooks"
```

`OFFICEIMO_JS_EVIDENCE_DIR` optionally retains the ordinary C# test's generated corpus beneath a caller-owned evidence directory. Without it, the test uses and deletes a unique temporary directory. Browser reports record engine/version, layer assertions, workbooks, CSV counts, worker/limit coverage and runtime errors. Example screenshots capture wide/compact states; screenshots complement the download assertions.

Host-dependent measurements stay explicit:

```powershell
node Build/BrowserExports/measure.mjs "$evidence/measurements"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/scale" --scale
```

The measurement script compares actual inline-string worksheet XML with an equivalent shared-string representation using the same compressor, and measures the generated bundles. The scale runner writes 100,000 rows by 20 columns for XLSX/CSV, validates output and records responsiveness and available heap metrics. These timing/memory measurements do not become ordinary CI correctness envelopes.

`Check-Excel.ps1` and `check-libreoffice.py` preserve optional application spot checks. They do not install, ship or invoke either application as a product dependency.
