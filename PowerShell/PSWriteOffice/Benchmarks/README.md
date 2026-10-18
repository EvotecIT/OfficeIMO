# PSWriteOffice Benchmarks

PSWriteOffice benchmarks use the PSPublishModule/PowerForge benchmark DSL. The
suite is split by file format so workbook behavior is compared with workbook
tools and CSV behavior is compared with CSV tools.

## Excel

`Compare-ExcelPerformance.ps1` measures OfficeIMO-backed PSWriteOffice Excel
cmdlets against PowerShell-facing Excel alternatives:

- `PSWriteOffice`
- `ImportExcel`
- `ExcelFast` for the workbook lanes it supports

Excel smoke, standard, and full suites use five measured iterations; large and
super-large suites use three. The grouped rotation keeps comparable workbook
engines adjacent and alternates which engine runs first. PowerForge performs an
explicit managed-memory cleanup after setup and data creation, outside the timed
operation, so allocations from one engine do not become another engine's GC bill.

Every read comparison uses the same PSWriteOffice-produced workbook shape for
the selected row count. The competing readers do not benchmark files created by
their own writers.

The current ExcelFast 0.0.1-alpha16 writer places an empty drawing element in an
invalid worksheet position. Its write lane therefore includes ExcelFast's own
`Get-Workbook` and `Save-Workbook` normalization inside the timed operation.
This keeps ExcelFast in the comparison without reporting malformed-workbook
speed as a valid result.

ExcelFast's mixed-object writer currently stores booleans, dates, and numbers
as culture-formatted text cells. Mixed and wide typed write lanes are therefore
not equivalent and remain excluded. The ExcelFast write comparison uses the
text-only default shape, while read lanes use the same PSWriteOffice-produced
typed workbook fixture as every other reader.

Runs that include PSWriteOffice build directly against a complete OfficeIMO
source checkout. With sibling PSWriteOffice and OfficeIMO checkouts, set the
source root once for the current shell:

```powershell
$env:OfficeIMORoot = (Resolve-Path ..\OfficeIMO).Path
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Smoke
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 -Suite Standard -Plan
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -RowCount 1000,5000,10000 `
    -RepeatCount 2 `
    -Engine PSWriteOffice,ImportExcel,ExcelFast `
    -UpdateReadme
```

### Streaming `IDataReader` tables

The `datareader-table` scenario measures the SQL-facing path: a forward-only
`IDataReader` is written as a real XLSX table and then validated for typed values
and Open XML package correctness. It is a PSWriteOffice capability lane, not a
cross-library ranking, because the compared PowerShell modules do not expose an
equivalent `IDataReader` workbook export.

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -RowCount 25000 `
    -RepeatCount 31 `
    -Scenario datareader-table `
    -Engine PSWriteOffice
```

A focused 2026-08-10 before/after run used 25,000 rows, twelve warmups, 31
measurements, High process priority, and affinities `0x1` and `0x10000` on an
AMD Ryzen 9 9950X3D2. The package-native path reduced observed mean export time
from 127.74 to 82.21 ms on `0x1` and from 84.95 to 47.31 ms on `0x10000`.
Managed allocation fell from about 12.37 MB to 3.00 MB in both processor
domains. These are workstation observations; the paired processor runs are kept
visible because the faster CPU domain can change by sub-operation.

<!-- BENCHMARK:ExcelComparison:START -->
| Scenario | Rows | PSWriteOffice | ExcelFast | ImportExcel | Result |
| --- | ---: | ---: | ---: | ---: | --- |
| append-existing-table | 1000 | 100.2 ms (1.00x) | Skipped | 340.0 ms (3.39x slower) | PSWriteOffice fastest |
| append-existing-table | 5000 | 323.2 ms (1.00x) | Skipped | 1.14 s (3.52x slower) | PSWriteOffice fastest |
| append-existing-table | 10000 | 741.3 ms (1.00x) | Skipped | 2.17 s (2.93x slower) | PSWriteOffice fastest |
| chart-only-workbook | 1000 | 91.0 ms (1.00x) | Skipped | 285.2 ms (3.13x slower) | PSWriteOffice fastest |
| chart-only-workbook | 5000 | 348.1 ms (1.00x) | Skipped | 1.07 s (3.09x slower) | PSWriteOffice fastest |
| chart-only-workbook | 10000 | 949.7 ms (1.00x) | Skipped | 2.01 s (2.12x slower) | PSWriteOffice fastest |
| csv-to-excel | 1000 | 73.4 ms (1.00x) | Skipped | 422.1 ms (5.75x slower) | PSWriteOffice fastest |
| csv-to-excel | 5000 | 270.0 ms (1.00x) | Skipped | 1.65 s (6.12x slower) | PSWriteOffice fastest |
| csv-to-excel | 10000 | 571.4 ms (1.00x) | Skipped | 3.40 s (5.95x slower) | PSWriteOffice fastest |
| datatable-default | 1000 | 21.0 ms (1.00x) | Skipped | 263.1 ms (12.52x slower) | PSWriteOffice fastest |
| datatable-default | 5000 | 27.1 ms (1.00x) | Skipped | 777.1 ms (28.62x slower) | PSWriteOffice fastest |
| datatable-default | 10000 | 36.1 ms (1.00x) | Skipped | 1.50 s (41.72x slower) | PSWriteOffice fastest |
| import-default-full | 1000 | 20.7 ms (1.00x) | 42.6 ms (2.06x slower) | 121.5 ms (5.86x slower) | PSWriteOffice fastest |
| import-default-full | 5000 | 96.4 ms (1.00x) | 106.3 ms (1.10x slower) | 276.7 ms (2.87x slower) | PSWriteOffice fastest |
| import-default-full | 10000 | 189.6 ms (1.00x) | 226.6 ms (1.19x slower) | 520.5 ms (2.75x slower) | PSWriteOffice fastest |
| import-default-range | 1000 | 18.0 ms (1.00x) | 48.2 ms (2.67x slower) | 125.0 ms (6.93x slower) | PSWriteOffice fastest |
| import-default-range | 5000 | 74.1 ms (1.00x) | 115.6 ms (1.56x slower) | 276.7 ms (3.73x slower) | PSWriteOffice fastest |
| import-default-range | 10000 | 110.7 ms (1.00x) | 224.6 ms (2.03x slower) | 399.2 ms (3.61x slower) | PSWriteOffice fastest |
| many-small-sheets | 1000 | 129.5 ms (1.00x) | Skipped | 373.1 ms (2.88x slower) | PSWriteOffice fastest |
| many-small-sheets | 5000 | 499.4 ms (1.00x) | Skipped | 1.22 s (2.44x slower) | PSWriteOffice fastest |
| many-small-sheets | 10000 | 989.5 ms (1.00x) | Skipped | 2.94 s (2.97x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 1000 | 91.8 ms (1.00x) | Skipped | 284.5 ms (3.10x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 5000 | 405.4 ms (1.00x) | Skipped | 1.04 s (2.57x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 10000 | 801.6 ms (1.00x) | Skipped | 2.00 s (2.50x slower) | PSWriteOffice fastest |
| named-range-workbook | 1000 | 85.7 ms (1.00x) | Skipped | 289.6 ms (3.38x slower) | PSWriteOffice fastest |
| named-range-workbook | 5000 | 328.0 ms (1.00x) | Skipped | 1.06 s (3.25x slower) | PSWriteOffice fastest |
| named-range-workbook | 10000 | 792.0 ms (1.00x) | Skipped | 2.00 s (2.53x slower) | PSWriteOffice fastest |
| objects-default | 1000 | 45.4 ms (1.00x) | Skipped | 302.2 ms (6.65x slower) | PSWriteOffice fastest |
| objects-default | 5000 | 101.0 ms (1.00x) | Skipped | 1.06 s (10.52x slower) | PSWriteOffice fastest |
| objects-default | 10000 | 226.3 ms (1.00x) | Skipped | 2.11 s (9.33x slower) | PSWriteOffice fastest |
| objects-no-table | 1000 | 25.6 ms (1.00x) | Skipped | 280.9 ms (10.98x slower) | PSWriteOffice fastest |
| objects-no-table | 5000 | 62.3 ms (1.00x) | Skipped | 1.05 s (16.88x slower) | PSWriteOffice fastest |
| objects-no-table | 10000 | 238.4 ms (1.00x) | Skipped | 2.29 s (9.60x slower) | PSWriteOffice fastest |
| objects-table | 1000 | 29.4 ms (1.00x) | Skipped | 318.9 ms (10.86x slower) | PSWriteOffice fastest |
| objects-table | 5000 | 63.7 ms (1.00x) | Skipped | 995.8 ms (15.63x slower) | PSWriteOffice fastest |
| objects-table | 10000 | 175.2 ms (1.00x) | Skipped | 2.05 s (11.68x slower) | PSWriteOffice fastest |
| objects-table-autofit | 1000 | 47.4 ms (1.00x) | Skipped | 299.2 ms (6.31x slower) | PSWriteOffice fastest |
| objects-table-autofit | 5000 | 70.7 ms (1.00x) | Skipped | 1.11 s (15.67x slower) | PSWriteOffice fastest |
| objects-table-autofit | 10000 | 203.8 ms (1.00x) | Skipped | 2.01 s (9.85x slower) | PSWriteOffice fastest |
| objects-title-freeze | 1000 | 89.5 ms (1.00x) | Skipped | 278.2 ms (3.11x slower) | PSWriteOffice fastest |
| objects-title-freeze | 5000 | 349.2 ms (1.00x) | Skipped | 970.0 ms (2.78x slower) | PSWriteOffice fastest |
| objects-title-freeze | 10000 | 892.7 ms (1.00x) | Skipped | 1.78 s (1.99x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 1000 | 59.7 ms (1.00x) | Skipped | 288.2 ms (4.83x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 5000 | 182.5 ms (1.00x) | Skipped | 1.04 s (5.71x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 10000 | 408.0 ms (1.00x) | Skipped | 2.08 s (5.10x slower) | PSWriteOffice fastest |
| read-named-range-metadata | 1000 | 17.4 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-named-range-metadata | 5000 | 19.1 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-named-range-metadata | 10000 | 18.1 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-no-header-range | 1000 | 18.0 ms (1.00x) | 46.0 ms (2.56x slower) | 126.2 ms (7.02x slower) | PSWriteOffice fastest |
| read-no-header-range | 5000 | 84.9 ms (1.00x) | 97.6 ms (1.15x slower) | 252.2 ms (2.97x slower) | PSWriteOffice fastest |
| read-no-header-range | 10000 | 110.7 ms (1.00x) | 220.1 ms (1.99x slower) | 427.8 ms (3.86x slower) | PSWriteOffice fastest |
| read-table-metadata | 1000 | 15.8 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-table-metadata | 5000 | 45.7 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-table-metadata | 10000 | 15.3 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 1000 | 19.0 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 5000 | 42.3 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| read-used-range-datatable | 10000 | 75.3 ms (1.00x) | Skipped | Skipped | PSWriteOffice fastest |
| report-workbook | 1000 | 100.9 ms (1.00x) | Skipped | 309.1 ms (3.06x slower) | PSWriteOffice fastest |
| report-workbook | 5000 | 374.6 ms (1.00x) | Skipped | 1.07 s (2.85x slower) | PSWriteOffice fastest |
| report-workbook | 10000 | 889.0 ms (1.00x) | Skipped | 2.09 s (2.35x slower) | PSWriteOffice fastest |
| summary-formulas | 1000 | 86.4 ms (1.00x) | Skipped | 272.5 ms (3.15x slower) | PSWriteOffice fastest |
| summary-formulas | 5000 | 298.0 ms (1.00x) | Skipped | 1.05 s (3.53x slower) | PSWriteOffice fastest |
| summary-formulas | 10000 | 748.6 ms (1.00x) | Skipped | 1.98 s (2.64x slower) | PSWriteOffice fastest |
| text-objects-default | 1000 | 73.2 ms (1.00x) | 194.1 ms (2.65x slower) | 424.1 ms (5.79x slower) | PSWriteOffice fastest |
| text-objects-default | 5000 | 99.5 ms (1.00x) | 427.8 ms (4.30x slower) | 1.71 s (17.14x slower) | PSWriteOffice fastest |
| text-objects-default | 10000 | 160.8 ms (1.00x) | 790.5 ms (4.92x slower) | 3.17 s (19.71x slower) | PSWriteOffice fastest |
| update-existing-workbook | 1000 | 170.8 ms (1.00x) | Skipped | 333.1 ms (1.95x slower) | PSWriteOffice fastest |
| update-existing-workbook | 5000 | 913.2 ms (1.00x) | Skipped | 1.38 s (1.51x slower) | PSWriteOffice fastest |
| update-existing-workbook | 10000 | 2.39 s (1.00x) | Skipped | 2.65 s (1.11x slower) | PSWriteOffice fastest |
| wide-objects-default | 1000 | 73.7 ms (1.00x) | Skipped | 227.7 ms (3.09x slower) | PSWriteOffice fastest |
| wide-objects-default | 5000 | 281.9 ms (1.00x) | Skipped | 827.0 ms (2.93x slower) | PSWriteOffice fastest |
| wide-objects-default | 10000 | 298.3 ms (1.00x) | Skipped | 1.52 s (5.09x slower) | PSWriteOffice fastest |
| workbook-package-merge | 1000 | 266.0 ms (1.00x) | Skipped | 653.3 ms (2.46x slower) | PSWriteOffice fastest |
| workbook-package-merge | 5000 | 857.1 ms (1.00x) | Skipped | 1.56 s (1.82x slower) | PSWriteOffice fastest |
| workbook-package-merge | 10000 | 1.45 s (1.00x) | Skipped | 2.66 s (1.83x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:ExcelComparison:END -->

## CSV

`Compare-CsvPerformance.ps1` measures PSWriteOffice CSV cmdlets against native
PowerShell CSV import/export:

- `PSWriteOffice`
- `NativeCsv`

The CSV release gate uses 25 measured iterations for suites that include the
short 1,000-row lanes. Large and super-large suites use 11 and 7 iterations,
respectively. PowerForge keeps each scenario comparison together and alternates
engine order between iterations, so neither engine keeps the first or second
position. An explicit managed-memory cleanup runs after setup and data creation,
outside the timed operation, for every engine. Use `-RepeatCount` for an
intentional diagnostic override; keep the defaults when recording release or
README results.

The CSV suite also includes dbatools-shaped read scenarios over the same
generated CSV shape used by `dataplat/dbatools.library/benchmarks/CsvBenchmarks`.
`csv-dbatools-quick-single-column` and `csv-dbatools-quick-all-columns` cover the
upstream `QuickTest` pattern and the small/medium/large reader shape by changing
`-RowCount`. The explicit `csv-dbatools-wide-*` and `csv-dbatools-quoted-*`
lanes cover the upstream wide 50-column and quote-all files.

The `csv-read-datatable-*` lanes compare `Import-OfficeCsv -AsDataTable` with a
native PowerShell baseline that reads via `Import-Csv` and fills a `DataTable`.
That keeps the database/table workflow visible instead of comparing only
PowerShell object materialization.

The `csv-write-gzip-*` and `csv-read-gzip-datatable-*` lanes keep compressed
CSV in the same benchmark matrix. The native baseline uses `GZipStream` with
`ConvertTo-Csv` / `ConvertFrom-Csv`, while PSWriteOffice uses
`Export-OfficeCsv -CompressionType GZip` and
`Import-OfficeCsv -CompressionType GZip`.

CSV runs that include PSWriteOffice use the same direct OfficeIMO source root:

```powershell
$env:OfficeIMORoot = (Resolve-Path ..\OfficeIMO).Path
```

### CSV Capability Coverage

The CSV benchmarks sit next to feature coverage on purpose: the target is fast
and correct, not a narrow parser trick. They measure the CSV reader and writer
paths used by PSWriteOffice, including streaming, compression, typed values, and
round-trip validation.

| Capability | Current PSWriteOffice / OfficeIMO.CSV support | Benchmark visibility |
| --- | --- | --- |
| Streaming reads | `Import-OfficeCsv` streams rows by default; `-AsDataTable` materializes one table when the next hop is tabular | Object read lanes and DataTable read lanes |
| Streaming writes | `Export-OfficeCsv` accepts objects, `DataTable`, `DataView`, and `IDataReader` input | Object/DataTable write lanes; DbaClientX export benchmark covers direct `IDataReader` SQL export |
| Compression | `Export-OfficeCsv` / `Import-OfficeCsv` expose `-CompressionType`; OfficeIMO.CSV supports none, auto, GZip, Deflate, and runtime-supported Brotli/ZLib | GZip write/read DataTable lanes |
| Cancellation | OfficeIMO.CSV accepts a cancellation token; PSWriteOffice cancels it from `StopProcessing`, so Ctrl+C can stop long reads | Contract coverage, not timed by default |
| Progress | OfficeIMO.CSV exposes progress callbacks; PSWriteOffice exposes `-ProgressInterval` and writes PowerShell progress records | Contract coverage, not timed by default |
| Schema and typed tables | OfficeIMO.CSV supports explicit/inferred schema, `DataTable`, and `IDataReader` schema tables; PSWriteOffice exposes `Import-OfficeCsv -AsDataTable -InferSchema` and `Import-OfficeCsv -AsDataReader -ColumnType` | DataTable read lanes; DbaClientX round-trip uses schema-aware reader handoff |
| CSV dialects | Culture/list separator, delimiter detection, multi-character delimiters, no-header/custom headers, comments, W3C headers, null values, date formats, quote modes, and selected quote fields | Default benchmark dialect plus focused command tests |
| Robust parsing | Duplicate-header policy, row-length mismatch policy, strict/lenient quotes, parse-error collection/skip-row, max field length, decompression limits, smart-quote normalization, and string interning | Command/core tests; timed lanes stay on clean generated files |
| Platform shape | PSWriteOffice targets Windows PowerShell/.NET Framework and PowerShell 7/.NET; Brotli/ZLib depend on runtime support, and comparison tools such as `bcp`/FastBCP/dbatools must be installed locally | Benchmarks skip unavailable optional engines |

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 -Suite Smoke
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 -Suite Standard -Plan
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Standard `
    -RowCount 1000,5000,10000 `
    -RepeatCount 25 `
    -Engine PSWriteOffice,NativeCsv `
    -UpdateReadme
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 100000 `
    -Scenario csv-dbatools-quick-single-column,csv-dbatools-quick-all-columns `
    -Engine PSWriteOffice,NativeCsv
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 100000 `
    -Scenario csv-dbatools-wide-single-column,csv-dbatools-wide-all-columns,csv-dbatools-quoted-single-column,csv-dbatools-quoted-all-columns `
    -Engine PSWriteOffice,NativeCsv
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 10000,100000 `
    -Scenario csv-write-datatable `
    -Engine PSWriteOffice,NativeCsv
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 10000,100000 `
    -Scenario csv-read-datatable-mixed,csv-read-datatable-wide `
    -Engine PSWriteOffice,NativeCsv
```

```powershell
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-CsvPerformance.ps1 `
    -Suite Smoke `
    -RowCount 10000,100000 `
    -Scenario csv-write-gzip-mixed,csv-read-gzip-datatable-mixed,csv-write-gzip-wide,csv-read-gzip-datatable-wide `
    -Engine PSWriteOffice,NativeCsv
```

Focused local dbatools QuickTest-shaped run, `20260708-053630-00691048`, 100,000
rows, 10 columns, three measured iterations:

| Scenario | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | --- |
| First column read | 341.3 ms, 294,869 rows/s | 347.3 ms, 285,460 rows/s | PSWriteOffice fastest |
| All columns read | 1.08 s, 92,385 rows/s | 1.12 s, 89,942 rows/s | PSWriteOffice fastest |

Focused local DataTable CSV write run, `20260708-060903-2e761e92`, five
measured iterations:

| Rows | PSWriteOffice median | NativeCsv median | Result |
| ---: | ---: | ---: | --- |
| 10,000 | 59.8 ms | 42.2 ms | NativeCsv fastest |
| 100,000 | 43.9 ms | 301.5 ms | PSWriteOffice fastest |

Focused local dbatools wide/quoted run, `20260708-072202-eb2e9470`, 10,000
rows, three measured iterations:

| Scenario | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | --- |
| Quoted all columns | 119.8 ms, 81,573 rows/s | 127.0 ms, 76,753 rows/s | PSWriteOffice fastest |
| Quoted first column | 44.8 ms, 222,418 rows/s | 48.7 ms, 205,266 rows/s | PSWriteOffice fastest |
| Wide all columns | 643.9 ms, 15,344 rows/s | 651.9 ms, 15,291 rows/s | PSWriteOffice fastest |
| Wide first column | 99.6 ms, 98,934 rows/s | 124.5 ms, 80,758 rows/s | PSWriteOffice fastest |

Focused local DataTable CSV read run, `20260708-100615-e9dbaa80`, 100,000
rows, three measured iterations. The native baseline reads with `Import-Csv`
and fills a `DataTable`, so this compares the table/database workflow rather
than only PowerShell object output:

<!-- BENCHMARK:CsvDataTableComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-read-datatable-mixed | 100000 | 219.6 ms (1.00x) | 2.57 s (11.72x slower) | PSWriteOffice fastest |
| csv-read-datatable-wide | 100000 | 1.60 s (1.00x) | 10.85 s (6.78x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:CsvDataTableComparison:END -->

Focused local GZip CSV run, generated by the benchmark updater. The write lanes
validate by decompressing and counting rows; the read lanes validate the returned
DataTable row count.

<!-- BENCHMARK:CsvGZipComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-read-gzip-datatable-mixed | 10000 | 77.5 ms (1.00x) | 205.2 ms (2.65x slower) | PSWriteOffice fastest |
| csv-read-gzip-datatable-wide | 10000 | 45.2 ms (1.00x) | 706.0 ms (15.63x slower) | PSWriteOffice fastest |
| csv-write-gzip-mixed | 10000 | 86.7 ms (1.00x) | 192.1 ms (2.22x slower) | PSWriteOffice fastest |
| csv-write-gzip-wide | 10000 | 118.8 ms (1.00x) | 397.7 ms (3.35x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:CsvGZipComparison:END -->

The 100,000-row values pool two independent 20-iteration runs for each
operation.

<!-- BENCHMARK:CsvComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-read-source-mixed | 100000 | 471.7 ms (1.00x) | 554.4 ms (1.18x slower) | PSWriteOffice fastest |
| csv-write-mixed | 100000 | 209.0 ms (1.00x) | 225.2 ms (1.08x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:CsvComparison:END -->

The same 100,000-row object import comparison across other file shapes:

| Shape | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | --- |
| Multiline | 472.1 ms | 546.7 ms | 16% faster |
| Quoted | 502.5 ms | 594.2 ms | 18% faster |
| Wide, 40 columns | 1.84 s | 2.27 s | 24% faster |

## Options

The wrappers build PSWriteOffice in `Release` mode by default and import local
development binaries when a selected run includes `PSWriteOffice`. Use
`-PSWriteOfficeConfiguration Debug` for diagnostics or `-SkipPSWriteOfficeBuild`
when reusing an existing local build. Quick and focused runs leave this
README unchanged unless `-UpdateReadme` is specified.

PSWriteOffice benchmark lanes require the current OfficeIMO source tree. Pass
`-OfficeIMORoot`, or set the `OfficeIMORoot` environment variable, so the build
uses direct project references. The wrappers stop instead of falling back to a
published OfficeIMO package when the source path is missing:

```powershell
$evotecRoot = if ($env:EVOTEC_GITHUB_ROOT) { $env:EVOTEC_GITHUB_ROOT } else { 'C:\Support\GitHub' }
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\Benchmarks\Compare-ExcelPerformance.ps1 `
    -Suite Standard `
    -OfficeIMORoot (Join-Path $evotecRoot 'OfficeIMO')
```

## Output

Benchmark artifacts are written under `Ignore\Benchmarks\ExcelPerformance` and
`Ignore\Benchmarks\CsvPerformance`:

- `samples.json` / `samples.csv`
- `summary.json` / `summary.csv`
- `comparison.json` / `comparison.md`
- `metadata.json`
- `run-report.json`

Start with `comparison.md` or the generated README table for the comparison,
then use `samples.csv` when diagnosing individual failures.
