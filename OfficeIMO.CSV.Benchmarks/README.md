# OfficeIMO CSV Benchmarks

This project compares raw .NET CSV paths without PowerShell object overhead. Use it beside the PSWriteOffice benchmark scoreboard, not as a replacement for it.

## Run

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter *CsvBenchmarks*
```

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV streaming/in-memory reads, CsvHelper typed record writing, CsvHelper projected-row writing, CsvHelper raw field reads, and CsvHelper typed record reads.

CsvHelper is a benchmark-only dependency in this project. It should not be added to `OfficeIMO.CSV` unless a future design decision intentionally changes the runtime dependency model.
