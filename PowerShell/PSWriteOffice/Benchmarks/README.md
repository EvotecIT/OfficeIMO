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

`Export-OfficeExcel` and `Add-OfficeExcelTable` fit column widths by default,
while plain `Export-Excel` and ExcelFast's `Export-Workbook` do not. Lanes whose
competitors do not fit columns run the PSWriteOffice side with `-NoAutoFit` so
both sides do the same work. `objects-table-autofit` and `report-workbook` keep
fitting on both sides (`Export-Excel -AutoSize`) and are the like-for-like
column-fitting comparison. PSWriteOffice still applies its table style and
frozen header row in every lane.

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

Runs that include PSWriteOffice build the module and OfficeIMO engines from this
checkout.

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

The Excel table below comes from PowerShell 7 runs of the Standard suite at
1,000, 5,000 and 10,000 rows. ExcelFast appears only in the lanes it supports.

<!-- BENCHMARK:ExcelComparison:START -->
| Scenario | Rows | PSWriteOffice | ExcelFast | ImportExcel | Result |
| --- | ---: | ---: | ---: | ---: | --- |
| append-existing-table | 1000 | 75.5 ms (1.00x) | n/a | 309.3 ms (4.09x slower) | PSWriteOffice fastest |
| append-existing-table | 5000 | 325.1 ms (1.00x) | n/a | 1.44 s (4.43x slower) | PSWriteOffice fastest |
| append-existing-table | 10000 | 523.7 ms (1.00x) | n/a | 1.68 s (3.21x slower) | PSWriteOffice fastest |
| chart-only-workbook | 1000 | 65.8 ms (1.00x) | n/a | 219.1 ms (3.33x slower) | PSWriteOffice fastest |
| chart-only-workbook | 5000 | 280.1 ms (1.00x) | n/a | 759.4 ms (2.71x slower) | PSWriteOffice fastest |
| chart-only-workbook | 10000 | 595.0 ms (1.00x) | n/a | 1.45 s (2.45x slower) | PSWriteOffice fastest |
| csv-to-excel | 1000 | 46.1 ms (1.00x) | n/a | 391.0 ms (8.48x slower) | PSWriteOffice fastest |
| csv-to-excel | 5000 | 52.1 ms (1.00x) | n/a | 1.28 s (24.62x slower) | PSWriteOffice fastest |
| csv-to-excel | 10000 | 95.7 ms (1.00x) | n/a | 2.44 s (25.45x slower) | PSWriteOffice fastest |
| datatable-default | 1000 | 6.9 ms (1.00x) | n/a | 201.7 ms (29.43x slower) | PSWriteOffice fastest |
| datatable-default | 5000 | 11.0 ms (1.00x) | n/a | 578.2 ms (52.71x slower) | PSWriteOffice fastest |
| datatable-default | 10000 | 17.2 ms (1.00x) | n/a | 1.16 s (67.61x slower) | PSWriteOffice fastest |
| import-default-full | 1000 | 28.5 ms (1.00x) | 67.7 ms (2.38x slower) | 110.9 ms (3.90x slower) | PSWriteOffice fastest |
| import-default-full | 5000 | 34.9 ms (1.00x) | 117.2 ms (3.36x slower) | 169.7 ms (4.86x slower) | PSWriteOffice fastest |
| import-default-full | 10000 | 92.3 ms (1.00x) | 259.6 ms (2.81x slower) | 319.7 ms (3.46x slower) | PSWriteOffice fastest |
| import-default-range | 1000 | 16.8 ms (1.00x) | 23.1 ms (1.37x slower) | 114.3 ms (6.80x slower) | PSWriteOffice fastest |
| import-default-range | 5000 | 63.7 ms (1.00x) | 119.2 ms (1.87x slower) | 222.7 ms (3.50x slower) | PSWriteOffice fastest |
| import-default-range | 10000 | 120.8 ms (1.00x) | 267.1 ms (2.21x slower) | 364.7 ms (3.02x slower) | PSWriteOffice fastest |
| many-small-sheets | 1000 | 41.1 ms (1.00x) | n/a | 225.4 ms (5.49x slower) | PSWriteOffice fastest |
| many-small-sheets | 5000 | 181.2 ms (1.00x) | n/a | 770.3 ms (4.25x slower) | PSWriteOffice fastest |
| many-small-sheets | 10000 | 443.2 ms (1.00x) | n/a | 1.52 s (3.44x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 1000 | 119.6 ms (1.00x) | n/a | 269.2 ms (2.25x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 5000 | 343.3 ms (1.00x) | n/a | 901.0 ms (2.62x slower) | PSWriteOffice fastest |
| multi-sheet-regions | 10000 | 358.0 ms (1.00x) | n/a | 1.46 s (4.06x slower) | PSWriteOffice fastest |
| named-range-workbook | 1000 | 36.6 ms (1.00x) | n/a | 200.3 ms (5.48x slower) | PSWriteOffice fastest |
| named-range-workbook | 5000 | 108.5 ms (1.00x) | n/a | 699.9 ms (6.45x slower) | PSWriteOffice fastest |
| named-range-workbook | 10000 | 289.3 ms (1.00x) | n/a | 1.42 s (4.90x slower) | PSWriteOffice fastest |
| objects-default | 1000 | 12.4 ms (1.00x) | n/a | 208.8 ms (16.85x slower) | PSWriteOffice fastest |
| objects-default | 5000 | 31.5 ms (1.00x) | n/a | 693.9 ms (22.05x slower) | PSWriteOffice fastest |
| objects-default | 10000 | 57.9 ms (1.00x) | n/a | 1.26 s (21.84x slower) | PSWriteOffice fastest |
| objects-no-table | 1000 | 12.1 ms (1.00x) | n/a | 210.5 ms (17.39x slower) | PSWriteOffice fastest |
| objects-no-table | 5000 | 29.3 ms (1.00x) | n/a | 672.5 ms (22.99x slower) | PSWriteOffice fastest |
| objects-no-table | 10000 | 55.8 ms (1.00x) | n/a | 1.32 s (23.70x slower) | PSWriteOffice fastest |
| objects-table | 1000 | 23.3 ms (1.00x) | n/a | 211.1 ms (9.07x slower) | PSWriteOffice fastest |
| objects-table | 5000 | 66.8 ms (1.00x) | n/a | 738.7 ms (11.06x slower) | PSWriteOffice fastest |
| objects-table | 10000 | 54.6 ms (1.00x) | n/a | 1.27 s (23.28x slower) | PSWriteOffice fastest |
| objects-table-autofit | 1000 | 60.7 ms (1.00x) | n/a | 194.4 ms (3.20x slower) | PSWriteOffice fastest |
| objects-table-autofit | 5000 | 251.9 ms (1.00x) | n/a | 770.3 ms (3.06x slower) | PSWriteOffice fastest |
| objects-table-autofit | 10000 | 505.3 ms (1.00x) | n/a | 1.34 s (2.65x slower) | PSWriteOffice fastest |
| objects-title-freeze | 1000 | 10.8 ms (1.00x) | n/a | 203.2 ms (18.85x slower) | PSWriteOffice fastest |
| objects-title-freeze | 5000 | 31.4 ms (1.00x) | n/a | 693.6 ms (22.08x slower) | PSWriteOffice fastest |
| objects-title-freeze | 10000 | 58.7 ms (1.00x) | n/a | 1.26 s (21.40x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 1000 | 37.9 ms (1.00x) | n/a | 204.4 ms (5.39x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 5000 | 128.9 ms (1.00x) | n/a | 794.6 ms (6.16x slower) | PSWriteOffice fastest |
| pivot-only-workbook | 10000 | 268.5 ms (1.00x) | n/a | 1.31 s (4.89x slower) | PSWriteOffice fastest |
| read-no-header-range | 1000 | 17.5 ms (1.00x) | 18.0 ms (1.03x slower) | 108.9 ms (6.21x slower) | PSWriteOffice and ExcelFast within 5% |
| read-no-header-range | 5000 | 83.4 ms (1.00x) | 109.1 ms (1.31x slower) | 213.3 ms (2.56x slower) | PSWriteOffice fastest |
| read-no-header-range | 10000 | 170.6 ms (1.00x) | 231.0 ms (1.35x slower) | 365.0 ms (2.14x slower) | PSWriteOffice fastest |
| report-workbook | 1000 | 75.7 ms (1.00x) | n/a | 251.6 ms (3.33x slower) | PSWriteOffice fastest |
| report-workbook | 5000 | 364.7 ms (1.00x) | n/a | 760.5 ms (2.09x slower) | PSWriteOffice fastest |
| report-workbook | 10000 | 722.8 ms (1.00x) | n/a | 1.42 s (1.96x slower) | PSWriteOffice fastest |
| summary-formulas | 1000 | 23.5 ms (1.00x) | n/a | 186.7 ms (7.96x slower) | PSWriteOffice fastest |
| summary-formulas | 5000 | 84.6 ms (1.00x) | n/a | 704.6 ms (8.33x slower) | PSWriteOffice fastest |
| summary-formulas | 10000 | 224.3 ms (1.00x) | n/a | 1.28 s (5.73x slower) | PSWriteOffice fastest |
| text-objects-default | 1000 | 27.3 ms (1.00x) | 275.4 ms (10.10x slower) | 295.6 ms (10.84x slower) | PSWriteOffice fastest |
| text-objects-default | 5000 | 31.4 ms (1.00x) | 326.7 ms (10.41x slower) | 1.12 s (35.63x slower) | PSWriteOffice fastest |
| text-objects-default | 10000 | 59.8 ms (1.00x) | 512.3 ms (8.56x slower) | 2.23 s (37.27x slower) | PSWriteOffice fastest |
| update-existing-workbook | 1000 | 101.4 ms (1.00x) | n/a | 373.3 ms (3.68x slower) | PSWriteOffice fastest |
| update-existing-workbook | 5000 | 325.9 ms (1.00x) | n/a | 993.2 ms (3.05x slower) | PSWriteOffice fastest |
| update-existing-workbook | 10000 | 707.6 ms (1.00x) | n/a | 1.85 s (2.62x slower) | PSWriteOffice fastest |
| wide-objects-default | 1000 | 25.8 ms (1.00x) | n/a | 162.0 ms (6.29x slower) | PSWriteOffice fastest |
| wide-objects-default | 5000 | 82.9 ms (1.00x) | n/a | 609.2 ms (7.35x slower) | PSWriteOffice fastest |
| wide-objects-default | 10000 | 180.3 ms (1.00x) | n/a | 1.22 s (6.75x slower) | PSWriteOffice fastest |
| workbook-package-merge | 1000 | 80.9 ms (1.00x) | n/a | 548.2 ms (6.78x slower) | PSWriteOffice fastest |
| workbook-package-merge | 5000 | 270.7 ms (1.00x) | n/a | 1.21 s (4.48x slower) | PSWriteOffice fastest |
| workbook-package-merge | 10000 | 488.2 ms (1.00x) | n/a | 2.03 s (4.16x slower) | PSWriteOffice fastest |
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

The `csv-read-datatable-*` lanes compare `Import-OfficeCsv -As DataTable` with a
native PowerShell baseline that reads via `Import-Csv` and fills a `DataTable`.
That keeps the database/table workflow visible instead of comparing only
PowerShell object materialization.

The `csv-convertto-*` and `csv-convertfrom-*` lanes cover in-memory conversion.
`ConvertTo-OfficeCsv` is compared with `ConvertTo-Csv` over the same objects;
both emit one string per CSV record, and the lane validates the record count and
every value. `Import-OfficeCsv -Text` is compared with `ConvertFrom-Csv` over the
same CSV text string and validates every returned value.

Windows PowerShell 5.1 has no `-UseQuotes`, so its native writers quote every
field. Under 5.1 the native write lanes and the native-generated source files for
the read lanes are therefore fully quoted; PowerShell 7 uses `-UseQuotes AsNeeded`.
Every engine still reads the same source file within one run.

The `csv-write-gzip-*` and `csv-read-gzip-datatable-*` lanes keep compressed
CSV in the same benchmark matrix. The native baseline uses `GZipStream` with
`ConvertTo-Csv` / `ConvertFrom-Csv`, while PSWriteOffice uses
`Export-OfficeCsv -CompressionType GZip` and
`Import-OfficeCsv -CompressionType GZip`.

CSV runs that include PSWriteOffice use the same direct OfficeIMO source root:

```powershell
```

### CSV Capability Coverage

The CSV benchmarks sit next to feature coverage on purpose: the target is fast
and correct, not a narrow parser trick. They measure the CSV reader and writer
paths used by PSWriteOffice, including streaming, compression, typed values, and
round-trip validation.

| Capability | Current PSWriteOffice / OfficeIMO.CSV support | Benchmark visibility |
| --- | --- | --- |
| Streaming reads | `Import-OfficeCsv` streams rows by default; `-As DataTable` materializes one table when the next hop is tabular | Object read lanes and DataTable read lanes |
| Streaming writes | `Export-OfficeCsv` accepts objects, `DataTable`, `DataView`, and `IDataReader` input | Object/DataTable write lanes; DbaClientX export benchmark covers direct `IDataReader` SQL export |
| Compression | `Export-OfficeCsv` / `Import-OfficeCsv` expose `-CompressionType`; OfficeIMO.CSV supports none, auto, GZip, Deflate, and runtime-supported Brotli/ZLib | GZip write/read DataTable lanes |
| Cancellation | OfficeIMO.CSV accepts a cancellation token; PSWriteOffice cancels it from `StopProcessing`, so Ctrl+C can stop long reads | Contract coverage, not timed by default |
| Progress | OfficeIMO.CSV exposes progress callbacks; PSWriteOffice exposes `-ProgressInterval` and writes PowerShell progress records | Contract coverage, not timed by default |
| Schema and typed tables | OfficeIMO.CSV supports explicit/inferred schema, `DataTable`, and `IDataReader` schema tables; PSWriteOffice exposes `Import-OfficeCsv -As DataTable -InferSchema` and `Import-OfficeCsv -As DataReader -ColumnType` | DataTable read lanes; DbaClientX round-trip uses schema-aware reader handoff |
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
| csv-read-datatable-mixed | 100000 | 131.3 ms (1.00x) | 1.44 s (10.99x slower) | PSWriteOffice fastest |
| csv-read-datatable-wide | 100000 | 556.7 ms (1.00x) | 5.16 s (9.27x slower) | PSWriteOffice fastest |
| csv-write-datatable | 100000 | 37.6 ms (1.00x) | 235.7 ms (6.28x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:CsvDataTableComparison:END -->

Focused local GZip CSV run, generated by the benchmark updater. The write lanes
validate by decompressing and counting rows; the read lanes validate the returned
DataTable row count.

<!-- BENCHMARK:CsvGZipComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-read-gzip-datatable-mixed | 10000 | 11.5 ms (1.00x) | 134.6 ms (11.71x slower) | PSWriteOffice fastest |
| csv-read-gzip-datatable-wide | 10000 | 42.2 ms (1.00x) | 449.6 ms (10.65x slower) | PSWriteOffice fastest |
| csv-write-gzip-mixed | 10000 | 23.8 ms (1.00x) | 128.1 ms (5.38x slower) | PSWriteOffice fastest |
| csv-write-gzip-wide | 10000 | 63.1 ms (1.00x) | 233.1 ms (3.69x slower) | PSWriteOffice fastest |
<!-- BENCHMARK:CsvGZipComparison:END -->

The 100,000-row values pool two independent 20-iteration runs for each
operation.

<!-- BENCHMARK:CsvComparison:START -->
| Scenario | Rows | PSWriteOffice | NativeCsv | Result |
| --- | ---: | ---: | ---: | --- |
| csv-convertfrom-mixed | 100000 | 348.9 ms (1.00x) | 445.8 ms (1.28x slower) | PSWriteOffice fastest |
| csv-convertto-mixed | 100000 | 160.7 ms (1.00x) | 182.6 ms (1.14x slower) | PSWriteOffice fastest |
| csv-read-source-mixed | 100000 | 381.4 ms (1.00x) | 456.1 ms (1.20x slower) | PSWriteOffice fastest |
| csv-write-mixed | 100000 | 138.0 ms (1.00x) | 172.6 ms (1.25x slower) | PSWriteOffice fastest |
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

PSWriteOffice benchmark lanes use the engines in the same OfficeIMO checkout.
Pass `-OfficeIMORoot` only when comparing an explicitly selected alternate engine
checkout. A missing source tree fails validation instead of selecting a public
package.

```powershell
pwsh -NoLogo -NoProfile -File ./Benchmarks/Compare-ExcelPerformance.ps1 -Suite Standard
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
