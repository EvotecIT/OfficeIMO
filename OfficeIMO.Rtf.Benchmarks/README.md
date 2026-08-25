# OfficeIMO.Rtf benchmarks

This project owns OfficeIMO-only performance coverage for native RTF parsing,
lossless round trips, semantic rewrites, and the HTML, Markdown, PDF, Word, and
Reader adapters. BenchmarkDotNet records time and managed allocations. The
isolated budget runner also records peak working set and output bytes.

Run a dry execution check before measuring:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks -- --filter '*RtfCoreBenchmarks*' --job Dry --noOverwrite
```

Run the normal BenchmarkDotNet suite for environment-qualified measurements:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks -- --filter '*RtfCoreBenchmarks*' '*RtfAdapterBenchmarks*' --noOverwrite
```

Verify the current regression ceilings with the isolated workflow runner:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks -- --verify-budgets
```

The ceilings in `rtf-benchmark-budgets.json` catch large regressions and hangs;
they are not throughput claims. Rebaseline them only from repeatable Release
runs on representative systems.

Equivalent RTF-to-HTML comparisons with RtfPipe live in the opt-in
`OfficeIMO.Rtf.Benchmarks.Comparisons` project outside the normal solution.
