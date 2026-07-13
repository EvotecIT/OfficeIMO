# OfficeIMO CSV Benchmarks

This project compares raw .NET CSV paths without PowerShell object overhead. Use it beside the PSWriteOffice benchmark scoreboard, not as a replacement for it.

## Current generated headline comparison

This compact table is selected from the same BenchmarkDotNet artifacts used by
the detailed investigations below and is refreshed through PSPublishModule.
Lower is faster; short-run results vary by machine, runtime, data, and options.

<!-- officeimo-csv-benchmark-table:start -->
| Scenario | Variables | Host | Operation | OfficeIMO.CSV | Sep | Sylvan.Data.Csv | Result |
| --- | --- | --- | --- | ---: | ---: | ---: | --- |
| Prevalidated text-row write upper bound | Contract=OfficeIMO caller-prevalidated text, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet short, Shape=wide, Snapshot=2026-07-13 | .NET 8 | Write rows | 1.00x (16ms) | 1.54x (24ms) | 1.73x (27ms) | OfficeIMO.CSV fastest |
| Wide field-span CSV read | Contract=field spans, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet short, Shape=wide, Snapshot=2026-07-13 | .NET 8 | Read every field | 1.00x (5ms) | 0.47x (2ms) | 1.89x (9ms) | OfficeIMO.CSV slower than Sep |
| Wide projected-row CSV write | Contract=projected values, Format=CSV, Rows=25,000, Runner=BenchmarkDotNet short, Shape=wide, Snapshot=2026-07-13 | .NET 8 | Write projected rows | 1.00x (37ms) | 0.66x (24ms) | 0.74x (27ms) | OfficeIMO.CSV slower than Sep |
<!-- officeimo-csv-benchmark-table:end -->

## Run

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter *Csv*Benchmarks*
```

Refresh the compact package-README comparison locally with one command:

```powershell
.\Build\Benchmarks\Update-BenchmarkReadmes.ps1 -Run Csv
```

The script runs only the focused equivalent lanes, calls PSPublishModule's
`Import-BenchmarkResult` and `Update-BenchmarkDocument`, and replaces the
marker-delimited tables. It is a deliberate local maintainer command; benchmark
execution is not scheduled in CI.

For a write-focused competitor pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*Write*" --job short --warmupCount 1 --iterationCount 3
```

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV trusted text-row writing, OfficeIMO.CSV direct IDataReader writing, OfficeIMO.CSV reusable reads, OfficeIMO.CSV field-span reads, OfficeIMO.CSV in-memory and streaming DataTable materialization with string and inferred-schema columns, OfficeIMO.CSV direct DbDataReader consumption and DbDataReader-to-DataTable loading, CsvHelper typed/projected writes, CsvHelper raw/typed reads, Sylvan raw/string/span field reads and DataTable loading, Dataplat.Dbatools.Csv reader/writer/DataTable paths, and Sep strict reader/writer paths.

Read lanes intentionally touch each field value and return a small checksum based on field count and text length. DataTable lanes materialize the table and then traverse the cells for the same checksum. Direct DbDataReader lanes traverse the public reader contract without first materializing a DataTable, while DbDataReader-to-DataTable lanes keep the ADO.NET table-loading path visible. This keeps the comparison honest: a lane cannot win by only counting rows or skipping the field payload.

For a SQL-shaped DataTable materialization pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadDataTable*" --job short --warmupCount 1 --iterationCount 3
```

For the streaming `CsvDocument.ToDataTable` paths used by thin consumers such as PSWriteOffice:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadStreamingDataTable*" --job short --warmupCount 1 --iterationCount 3
```

For a dbatools.library-shaped CSV reader pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvDbatoolsLibraryParityBenchmarks*" --job short --warmupCount 1 --iterationCount 3
```

`CsvDbatoolsLibraryParityBenchmarks` mirrors the published dbatools.library CSV benchmark layout from [dataplat/dbatools.library `benchmarks/CsvBenchmarks`](https://github.com/dataplat/dbatools.library/tree/main/benchmarks/CsvBenchmarks), specifically `CsvReaderBenchmarks.Benchmarks.cs` and `QuickTest.cs`: small, medium, large, wide, quoted, modern medium/large, all-values, and quick-test-style single-column/all-column read lanes. It keeps OfficeIMO in the same file-path reader shape beside Dataplat.Dbatools.Csv, LumenWorks, Sep, Sylvan, and CsvHelper so the raw parser comparison is apples-to-apples. Each parity lane validates the expected row count and deterministic field-length checksum for its input file, so a lane cannot win by silently under-reading or skipping field materialization. The broader `CsvBenchmarks` and `CsvWideBenchmarks` lanes still touch every field and return checksums for stricter payload validation.

Parity check: the class includes all 20 upstream `CsvReaderBenchmarks` methods by benchmark description plus all 10 QuickTest read lanes, then adds matching OfficeIMO lanes beside them. The extra `OfficeIMO-DataReader-QuickTest-GetValues` lane keeps the SQL/bulk-copy-shaped `DbDataReader.GetValues` path visible at the same 100k-row QuickTest size. Dataplat remains the BenchmarkDotNet baseline in this parity class to preserve the upstream comparison frame. `TypeConverterBenchmarks` is intentionally out of scope here because it measures dbatools vector conversion rather than CSV parser throughput, not CSV reader throughput.

CsvHelper, Sylvan.Data.Csv, Dataplat.Dbatools.Csv, LumenWorksCsvReader2, and Sep are benchmark-only dependencies in this project. They should not be added to `OfficeIMO.CSV` unless a future design decision intentionally changes the runtime dependency model.

## Current dbatools.library Parity Snapshot

Fresh local short-job run on 2026-07-09 using the 14 QuickTest-shaped parity lanes with row-count and field-length checksum validation enabled. Treat these as direction-finding numbers; run a longer BenchmarkDotNet job before release claims or performance gates.

QuickTest single-column/all-column read lanes:

| Method | Single column mean | All columns mean | Allocated |
| --- | ---: | ---: | ---: |
| OfficeIMO span reader | 4.57 ms | 4.70 ms | 770 KB / 770 KB |
| OfficeIMO streaming DataReader | 19.68 ms | 21.33 ms | 40.6 MB |
| SEP | 7.14 ms | 15.59 ms | 3.1 MB / 39.4 MB |
| Sylvan | 8.23 ms | 16.34 ms | 3.1 MB / 39.6 MB |
| CsvHelper | 30.15 ms | 45.63 ms | 3.1 MB / 39.6 MB |
| Dataplat.Dbatools.Csv | 26.13 ms | 29.38 ms | 39.9 MB |
| LumenWorks | 118.19 ms | 33.17 ms | 1.58 GB / 39.7 MB |

All-values read lane:

| Method | Mean | Allocated |
| --- | ---: | ---: |
| OfficeIMO span reader | 4.60 ms | 770 KB |
| OfficeIMO streaming DataReader | 22.64 ms | 41.3 MB |
| Dataplat.Dbatools.Csv DataReader | 26.70 ms | 39.9 MB |
| LumenWorks | 32.53 ms | 39.7 MB |

Additional guardrail: `OfficeIMO-DataReader-QuickTest-GetValues` reads the same
100k-row QuickTest file through `DbDataReader.GetValues`; the latest 2026-07-09
short run measured 23.10 ms and 40.6 MB. Keep this lane visible when optimizing
the SQL/bulk-copy-shaped reader path.

The span-reader result is the fastest raw parser shape. The streaming DataReader result is the SQL/bulk-copy-shaped path; it now reads reusable parser string rows directly and is faster than Dataplat's DataReader in these short runs, with Dataplat still holding a small allocation edge.

## Current Typed DataReader Snapshot

Fresh local short-job runs on 2026-07-09 using the 25,000-row, 40-column wide payload. Every lane traverses every value. The file lane includes file decoding and uses the public `CsvDocument.CreateDataReader(path, ...)` API used by PSWriteOffice and DbaClientX.

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*DataReader*Schema*" --job short --warmupCount 5 --iterationCount 10
```

| Input | Schema | Mean | Allocated |
| --- | --- | ---: | ---: |
| CSV file | Explicit 40-column schema | 103.11 ms | 101.49 MB |
| CSV text | Explicit 40-column schema | 91.41 ms | 66.95 MB |
| CSV text | Inferred from 25,000 rows | 135.78 ms | 66.97 MB |

Explicit typed readers parse numbers, booleans, dates, and GUIDs directly from source spans. Inferred readers inspect spans without retaining sampled rows, then replay the immutable text through the typed reader. String-only file readers stay on the lower-memory streaming path.

## Current Write Snapshot

Fresh local short-job run on 2026-07-08:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*Write*" --job short --warmupCount 1 --iterationCount 3
```

The table shows the fastest method per shape/row-count lane. Treat this as a quick comparison snapshot; run a longer BenchmarkDotNet job before making release claims or budget gates.

| Shape | Rows | Fastest method | Mean |
| --- | ---: | --- | ---: |
| Mixed | 1000 | OfficeIMO_WriteTrustedTextRows | 0.08 ms |
| Mixed | 10000 | OfficeIMO_WriteTrustedTextRows | 0.98 ms |
| Mixed | 25000 | OfficeIMO_WriteTrustedTextRows | 2.28 ms |
| Multiline | 1000 | OfficeIMO_WriteTrustedTextRows | 0.12 ms |
| Multiline | 10000 | OfficeIMO_WriteTrustedTextRows | 1.37 ms |
| Multiline | 25000 | OfficeIMO_WriteTrustedTextRows | 3.79 ms |
| Quoted | 1000 | OfficeIMO_WriteTrustedTextRows | 0.14 ms |
| Quoted | 10000 | OfficeIMO_WriteTrustedTextRows | 1.27 ms |
| Quoted | 25000 | OfficeIMO_WriteTrustedTextRows | 4.22 ms |
| Wide | 1000 | OfficeIMO_WriteTrustedTextRows | 0.28 ms |
| Wide | 10000 | OfficeIMO_WriteTrustedTextRows | 3.41 ms |
| Wide | 25000 | OfficeIMO_WriteTrustedTextRows | 9.45 ms |

## Current Wide IDataReader Write Snapshot

Fresh local short-job run on 2026-07-07:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*Write*" --job short --warmupCount 1 --iterationCount 3
```

The table keeps the SQL-shaped writer path visible. `OfficeIMO_WriteDataReader` writes through the public `IDataReader` bridge, `Dataplat_WriteFromReader` uses the dbatools reader bridge, and the trusted text lane shows the faster path available when the caller already supplies culture-aware formatting and schema validation.

| Rows | OfficeIMO IDataReader | dbatools reader | Sylvan reader | SEP projected | OfficeIMO trusted text |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 1000 | 1.68 ms | 1.76 ms | 1.12 ms | 1.21 ms | 0.67 ms |
| 10000 | 18.36 ms | 20.68 ms | 14.82 ms | 14.48 ms | 6.26 ms |
| 25000 | 47.57 ms | 78.13 ms | 35.75 ms | 33.28 ms | 16.02 ms |

In this run the OfficeIMO `IDataReader` bridge stays ahead of dbatools' reader bridge at every tested row count while allocating about one-third as much memory. Sylvan and SEP remain faster in the projected-row lanes, and the OfficeIMO trusted text lane shows the upper bound when the caller already supplies formatting and schema validation.

## Current Wide Read Snapshot

Fresh local short-job run on 2026-07-07:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*Read*FieldSpan*" --job short --warmupCount 1 --iterationCount 3
```

The table shows the fastest raw field-span read method per wide row-count lane. These lanes touch every field and compare OfficeIMO.CSV against SEP and Sylvan without PowerShell object overhead.

| Shape | Rows | Fastest method | Mean | SEP span read | Sylvan span read |
| --- | ---: | --- | ---: | ---: | ---: |
| Wide | 1000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.06 ms | 0.08 ms | 0.11 ms |
| Wide | 10000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.67 ms | 0.87 ms | 1.05 ms |
| Wide | 25000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 1.73 ms | 2.09 ms | 2.79 ms |
