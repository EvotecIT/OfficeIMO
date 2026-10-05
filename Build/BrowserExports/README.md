# Browser export verification

The verification harness checks browser-produced XLSX with OfficeIMO.Excel's editable model, its public streaming reader and the Open XML SDK validator. It also compares browser CSV bytes with the same vectors used by OfficeIMO.CSV.Tests. HtmlTinkerX owns browser installation and session lifetime; all three engines run sequentially and each session is disposed before the next one starts.

Use the repository's pinned .NET SDK, Node 24 and PowerShell 7. The harness uses the published HtmlTinkerX package only for testing. It is outside the normal solution, is not packable and adds no dependency to OfficeIMO.Browser. Browser runtimes are installed on demand through HtmlTinkerX; allow network access for first-run installation. Microsoft Excel and LibreOffice are optional application spot checks, not runtime requirements.

## Correctness and package checks

From the repository root, choose an evidence directory outside the checkout:

```powershell
$evidence = '/path/to/browser-export-evidence'
node OfficeIMO.Browser/Build/assets.mjs --check
node --test OfficeIMO.Browser/tests/*.test.mjs
dotnet test OfficeIMO.Browser.Tests/OfficeIMO.Browser.Tests.csproj -c Release
dotnet test OfficeIMO.CSV.Tests/OfficeIMO.CSV.Tests.csproj -c Release --filter 'Category!=Performance'
dotnet run --project OfficeIMO.Browser.Examples/OfficeIMO.Browser.Examples.csproj -c Release -- "$evidence/example"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/interop" "--example=$evidence/example" --limits
./Build/BrowserExports/Test-NpmConsumer.ps1 -EvidenceDirectory "$evidence/npm"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- --validate "$evidence/npm/packed-consumer.xlsx"
./Build/BrowserExports/Test-NugetConsumer.ps1 -EvidenceDirectory "$evidence/nuget"
```

The browser fixtures cover typed values, formats, widths, bold/fill/wrap/alignment, properties, frozen panes, autofilters, multiple sanitized sheet names, Unicode and XML escaping, literal OOXML escapes, empty/one-cell sheets, maximum cell length and column count, cancellation, local/UTC dates, and both kinds of unavailable raw-deflate support. A host-owned classic Blob worker checks the same exports with cloned dates and returned Blobs from a local-file page. `--limits` actually attempts 1,048,576 data rows with a header and verifies rejection and failed finalization at the sheet limit. Node's independent ZIP reader verifies CRC-32 and inflates compressed payloads.

The example check opens inline and bundle pages through `file://`, exercises current-view XLSX and CSV downloads, verifies empty-view output, captures wide/compact screenshots and fails on external network requests or page errors. Inspect the screenshots as well as the automated results. The existing .NET workflow runs these correctness and package checks without performance thresholds.

The npm consumer installs the actual archive and test-only TypeScript 5.9.3 in an isolated folder. It checks strict positive and negative declarations, readonly domain records, async arrays, package export paths and real writes. The library package contains no runtime or development dependencies. The NuGet consumer restores the actual archive into a task-local package directory and verifies all six embedded assets and their SHA-256 names on .NET 8 and .NET 10.

## Opt-in measurements

```powershell
node Build/BrowserExports/measure.mjs "$evidence/measurements"
dotnet run --project Build/BrowserExports/OfficeIMO.Browser.Interop.csproj -c Release -- . "$evidence/interop" "--example=$evidence/example" --scale
```

The size tool reports raw bytes, gzip level 9 and Brotli quality 11 for every module/classic build. It compares actual inline worksheet XML with an equivalent indexed worksheet and shared-string table, using the same deflater for both. The comparison excludes ZIP metadata; dictionary sizes count text only, without Map/object overhead. It is a measurement prototype, not another production writer.

`--scale` writes 100,000 rows × 20 columns in every browser, mixing numbers, repeated strings and unique strings. It records elapsed time, Blob bytes, a 10 ms timer's largest gap and Long Tasks entries where supported. Chromium enables precise JavaScript heap reporting and samples it at progress notifications. Firefox and WebKit do not expose that heap API. Samples exclude native stream/Blob memory, can miss brief peaks and must not be described as a whole-process peak-memory bound. Long Tasks data is unavailable in engines that do not support that observer. Each scale XLSX receives full Open XML validation and a streaming-reader check of its final row.

Measurement output has no CI timing or memory threshold. Retain compact JSON reports and the fixtures needed for review; remove superseded scale output and isolated consumers after verification.

## Optional application spot checks

On Windows with Excel installed, run:

```powershell
./Build/BrowserExports/Check-Excel.ps1 -WorkbookPath "$evidence/interop/chromium/rich-auto.xlsx" -EvidenceDirectory $evidence
```

The script opens the file read-only in an isolated hidden Excel instance, verifies literal text, numeric/boolean values, early-1900 serial dates, frozen headers and styling, then fits the preview's columns/rows, exports a PDF and closes Excel without saving. The preview fitting makes unspecified column widths readable; it does not add an autofit feature to the writer. The JSON report records the original declared width, formatted date and application version. The fixture file stays unchanged.

For LibreOffice, convert the same fixture with an isolated task-owned profile:

```sh
libreoffice -env:UserInstallation=file:///absolute/evidence/libreoffice/profile --headless \
  --convert-to 'csv:Text - txt - csv (StarCalc):44,34,76,1' \
  --outdir /absolute/evidence/libreoffice /absolute/evidence/interop/chromium/rich-auto.xlsx
python Build/BrowserExports/check-libreoffice.py /absolute/evidence/libreoffice
```

The report separates preserved values from date/newline differences. LibreOffice 24.2.7 normalizes CRLF inside a cell to LF when converting to CSV and displays pre-March-1900 dates one day earlier. These are disclosed application interoperability limits; Excel's date-system contract and the original XML text remain the writer's output. Remove the isolated profile when the conversion process has exited.
