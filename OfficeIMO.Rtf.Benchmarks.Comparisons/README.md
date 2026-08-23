# OfficeIMO RTF comparisons

This opt-in BenchmarkDotNet project compares complete RTF-to-HTML conversions
through OfficeIMO and RtfPipe. Both lanes receive the same RTF text, include
parsing and HTML generation, and must preserve the expected text and table
structure before timing begins.

RtfPipe remains a benchmark-only dependency. This project is intentionally not
part of `OfficeIMO.sln`, so normal restore, build, test, and package operations
do not acquire it.

## Validate output and size

Run the semantic preflight without collecting timings:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\rtf\validation.json
```

The report records UTF-8 input and output bytes, visible record count, tables,
cells, and images. Output byte counts describe each library's generated HTML;
smaller output is useful evidence only after the semantic checks pass.

## Measure time and allocations

Start with a dry execution check:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- --filter '*RtfToHtmlComparisonBenchmarks*' --job Dry --noOverwrite
```

Use the short job while changing the workload, then the default job for a
recorded comparison:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- --filter '*RtfToHtmlComparisonBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Rtf.Benchmarks.Comparisons -- --filter '*RtfToHtmlComparisonBenchmarks*' --noOverwrite
```

BenchmarkDotNet artifacts and the validation JSON are machine-specific. Keep
them under `.benchmark-artifacts` or another ignored output root and publish
only environment-qualified results.

The shared runner adds source provenance and normalized evidence:

```powershell
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload rtfhtml -RunMode full -Framework net8.0
```
