# OfficeIMO.Tabular benchmark suite

This suite reproduces the public 65K-record CSV, XLSX, and XLSB comparisons with
one pinned fixture revision and the same typed getters or object binding on both
libraries. It benchmarks the supported `TabularReader` API, not internal CSV or
Excel backends.

The runner downloads the three fixtures from commit
`5e1113a1195bed985c10788a6b89caf551663bb1` of MarkPflug/Benchmarks into the
system temporary folder and verifies their SHA-256 hashes before validation or
measurement. Set `OFFICEIMO_TABULAR_BENCHMARK_DATA` to use an existing fixture
folder.

Validate every lane before measuring:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Tabular.Benchmarks -- --validate
```

Run a short local diagnostic:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Tabular.Benchmarks -- --quick --artifacts .\artifacts\tabular-quick
```

Run publication-grade BenchmarkDotNet jobs:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Tabular.Benchmarks -- --artifacts .\artifacts\tabular-full
```

Windows, Linux, and macOS results are independent evidence lanes. They are
never averaged together. A missing platform stays visible as missing rather
than borrowing another platform's result.
