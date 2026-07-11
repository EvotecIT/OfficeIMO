# OfficeIMO.OpenDocument benchmarks

This project tracks the performance and allocations of three contracts that are easy to regress: opening and enumerating a 2,000-paragraph ODT, writing an ODS cell at an extreme sparse coordinate, and evaluating a 1,000-cell OpenFormula range.

Run the full .NET 8 benchmark set from the repository root:

```powershell
dotnet run --project OfficeIMO.OpenDocument.Benchmarks/OfficeIMO.OpenDocument.Benchmarks.csproj -c Release -f net8.0
```

BenchmarkDotNet results are machine-specific engineering evidence, not universal throughput guarantees. Keep generated `BenchmarkDotNet.Artifacts` out of source control.
