# OfficeIMO CSV Benchmarks

This project compares raw .NET CSV paths without PowerShell object overhead. Use it beside the PSWriteOffice benchmark scoreboard, not as a replacement for it.

## Run

```powershell
dotnet run --project .\OfficeIMO.CSV.Benchmarks\OfficeIMO.CSV.Benchmarks.csproj -c Release -f net8.0 -- --filter *Csv*Benchmarks*
```

The suite compares OfficeIMO.CSV object writing, OfficeIMO.CSV projected-row writing, OfficeIMO.CSV trusted text-row writing, OfficeIMO.CSV reusable reads, OfficeIMO.CSV field-span reads, CsvHelper typed/projected writes, CsvHelper raw/typed reads, Sylvan raw/string/span field reads and data-reader writes, Dataplat.Dbatools.Csv reader/writer paths, and Sep strict reader/writer paths.

CsvHelper, Sylvan.Data.Csv, Dataplat.Dbatools.Csv, and Sep are benchmark-only dependencies in this project. They should not be added to `OfficeIMO.CSV` unless a future design decision intentionally changes the runtime dependency model.
