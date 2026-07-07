# OfficeIMO CSV Benchmarks

This project compares raw .NET CSV paths without PowerShell object overhead. Use it beside the PSWriteOffice benchmark scoreboard, not as a replacement for it.

## Run

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter *Csv*Benchmarks*
```

For a write-focused competitor pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*Write*" --job short --warmupCount 1 --iterationCount 3
```

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV trusted text-row writing, OfficeIMO.CSV direct IDataReader writing, OfficeIMO.CSV reusable reads, OfficeIMO.CSV field-span reads, OfficeIMO.CSV string and inferred-schema DataTable materialization, OfficeIMO.CSV direct DbDataReader consumption and DbDataReader-to-DataTable loading, CsvHelper typed/projected writes, CsvHelper raw/typed reads, Sylvan raw/string/span field reads and DataTable loading, Dataplat.Dbatools.Csv reader/writer/DataTable paths, and Sep strict reader/writer paths.

Read lanes intentionally touch each field value and return a small checksum based on field count and text length. DataTable lanes materialize the table and then traverse the cells for the same checksum. Direct DbDataReader lanes traverse the public reader contract without first materializing a DataTable, while DbDataReader-to-DataTable lanes keep the ADO.NET table-loading path visible. This keeps the comparison honest: a lane cannot win by only counting rows or skipping the field payload.

For a SQL-shaped DataTable materialization pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadDataTable*" --job short --warmupCount 1 --iterationCount 3
```

For a dbatools.library-shaped CSV reader pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvDbatoolsLibraryParityBenchmarks*" --job short --warmupCount 1 --iterationCount 3
```

`CsvDbatoolsLibraryParityBenchmarks` mirrors the published dbatools.library CSV benchmark layout from `benchmarks/CsvBenchmarks/CsvReaderBenchmarks.Benchmarks.cs` and `benchmarks/CsvBenchmarks/QuickTest.cs`: small, medium, large, wide, quoted, modern medium/large, all-values, and quick-test-style single-column/all-column read lanes. It keeps OfficeIMO in the same file-path reader shape beside Dataplat.Dbatools.Csv, LumenWorks, Sep, Sylvan, and CsvHelper so the raw parser comparison is apples-to-apples. The broader `CsvBenchmarks` and `CsvWideBenchmarks` lanes still touch every field and return checksums for stricter payload validation.

Parity check: the class includes all 20 `CsvReaderBenchmarks` descriptions from `dataplat/dbatools.library` plus all 10 QuickTest read lanes, then adds matching OfficeIMO lanes beside them. Dataplat remains the BenchmarkDotNet baseline in this parity class to preserve the upstream comparison frame. `TypeConverterBenchmarks` is intentionally out of scope here because it measures dbatools vector conversion rather than CSV parser throughput, not CSV reader throughput.

CsvHelper, Sylvan.Data.Csv, Dataplat.Dbatools.Csv, LumenWorksCsvReader2, and Sep are benchmark-only dependencies in this project. They should not be added to `OfficeIMO.CSV` unless a future design decision intentionally changes the runtime dependency model.

## Current dbatools.library Parity Snapshot

Fresh local short-job runs on 2026-07-08. Treat these as direction-finding numbers; run a longer BenchmarkDotNet job before release claims or performance gates.

QuickTest single-column/all-column read lanes:

| Method | Single column mean | All columns mean | Allocated |
| --- | ---: | ---: | ---: |
| OfficeIMO span reader | 4.57 ms | 4.66 ms | ~771 KB |
| OfficeIMO streaming DataReader | 15.29 ms | 20.51 ms | 41.3 MB |
| SEP | 6.44 ms | 16.76 ms | 3.1 MB / 39.4 MB |
| Sylvan | 11.12 ms | 19.21 ms | 3.1 MB / 39.6 MB |
| CsvHelper | 32.68 ms | 47.22 ms | 3.1 MB / 39.6 MB |
| Dataplat.Dbatools.Csv | 29.32 ms | 28.98 ms | 39.9 MB |
| LumenWorks | 104.36 ms | 36.89 ms | 1.58 GB / 39.7 MB |

All-values read lane:

| Method | Mean | Allocated |
| --- | ---: | ---: |
| OfficeIMO span reader | 4.68 ms | 772 KB |
| OfficeIMO streaming DataReader | 18.75 ms | 41.3 MB |
| Dataplat.Dbatools.Csv DataReader | 26.38 ms | 39.9 MB |
| LumenWorks | 38.56 ms | 39.7 MB |

The span-reader result is the fastest raw parser shape. The streaming DataReader result is the SQL/bulk-copy-shaped path; it now reads reusable parser string rows directly and is faster than Dataplat's DataReader in these short runs, with Dataplat still holding a small allocation edge.

## Current Write Snapshot

Fresh local short-job run on 2026-07-07:

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

The table keeps the SQL-shaped writer path visible. `OfficeIMO_WriteDataReader` writes through the public `IDataReader` bridge, `Dataplat_WriteFromReader` uses the dbatools reader bridge, and the trusted text lane shows the faster path available when the caller already owns culture-aware formatting and schema validation.

| Rows | OfficeIMO IDataReader | dbatools reader | Sylvan reader | SEP projected | OfficeIMO trusted text |
| ---: | ---: | ---: | ---: | ---: | ---: |
| 1000 | 1.46 ms | 1.63 ms | 1.25 ms | 1.05 ms | 0.27 ms |
| 10000 | 16.84 ms | 18.17 ms | 13.17 ms | 11.98 ms | 3.68 ms |
| 25000 | 42.44 ms | 55.73 ms | 32.87 ms | 28.18 ms | 10.26 ms |

## Current Wide Read Snapshot

Fresh local short-job run on 2026-07-07:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*CsvWideBenchmarks*Read*FieldSpan*" --job short --warmupCount 1 --iterationCount 3
```

The table shows the fastest raw field-span read method per wide row-count lane. These lanes touch every field and compare OfficeIMO.CSV against SEP and Sylvan without PowerShell object overhead.

| Shape | Rows | Fastest method | Mean | SEP span read | Sylvan span read |
| --- | ---: | --- | ---: | ---: | ---: |
| Wide | 1000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.07 ms | 0.10 ms | 0.14 ms |
| Wide | 10000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 0.77 ms | 1.10 ms | 1.45 ms |
| Wide | 25000 | OfficeIMO_ReadTextFieldSpanVisitorSkipHeader | 1.91 ms | 2.87 ms | 3.71 ms |
