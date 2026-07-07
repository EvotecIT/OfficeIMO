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

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV trusted text-row writing, OfficeIMO.CSV reusable reads, OfficeIMO.CSV field-span reads, OfficeIMO.CSV string and inferred-schema DataTable materialization, OfficeIMO.CSV direct DbDataReader consumption and DbDataReader-to-DataTable loading, CsvHelper typed/projected writes, CsvHelper raw/typed reads, Sylvan raw/string/span field reads and DataTable loading, Dataplat.Dbatools.Csv reader/writer/DataTable paths, and Sep strict reader/writer paths.

Read lanes intentionally touch each field value and return a small checksum based on field count and text length. DataTable lanes materialize the table and then traverse the cells for the same checksum. Direct DbDataReader lanes traverse the public reader contract without first materializing a DataTable, while DbDataReader-to-DataTable lanes keep the ADO.NET table-loading path visible. This keeps the comparison honest: a lane cannot win by only counting rows or skipping the field payload.

For a SQL-shaped DataTable materialization pass:

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter "*ReadDataTable*" --job short --warmupCount 1 --iterationCount 3
```

CsvHelper, Sylvan.Data.Csv, Dataplat.Dbatools.Csv, and Sep are benchmark-only dependencies in this project. They should not be added to `OfficeIMO.CSV` unless a future design decision intentionally changes the runtime dependency model.

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
