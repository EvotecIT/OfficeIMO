# OfficeIMO OpenDocument comparisons

This opt-in BenchmarkDotNet project compares complete ODS create/save and
open/read workflows through OfficeIMO and OpenStandardLibrary. Both creation
lanes build the same dense string-cell corpus and return an in-memory ODS
package. Both read lanes receive the same OfficeIMO-generated package and
enumerate every populated cell.

OpenStandardLibrary remains a benchmark-only dependency. This project is
intentionally not part of `OfficeIMO.sln`, so normal restore, build, test, and
package operations do not acquire it. The comparison is limited to ODS because
OpenStandardLibrary does not expose equivalent ODT or ODP document models.

## Validate output and size

Run the semantic and package-size preflight without collecting timings:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\opendocument\comparison-validation.json
```

The validator reopens both generated packages through OfficeIMO and verifies
sheet count, expanded row and cell counts, content length, and boundary markers.
Output byte counts are meaningful only after those equivalent-content checks
pass.

## Measure time and allocations

Start with a dry execution check:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- --filter '*Ods*ComparisonBenchmarks*' --job Dry --noOverwrite
```

Use the short job while changing a workload, then the default job for recorded
evidence:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- --filter '*Ods*ComparisonBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net8.0 --project .\OfficeIMO.OpenDocument.Benchmarks.Comparisons -- --filter '*Ods*ComparisonBenchmarks*' --noOverwrite
```

BenchmarkDotNet artifacts and validation JSON are environment-specific. Keep
them under `.benchmark-artifacts` or another ignored output root.
