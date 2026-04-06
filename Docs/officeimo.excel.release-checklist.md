# OfficeIMO Excel Release Checklist

Use this checklist when publishing a new `OfficeIMO.Excel` package baseline.

## Goal

Ship an Excel package baseline that:

- matches the current public docs and compatibility notes
- has green Excel-focused validation on the supported targets
- has a benchmark harness available for quick sanity checks
- is clear about current strengths, partial areas, and roadmap gaps

## Before Bumping Versions

- Confirm the Excel-specific release-prep PRs are merged:
  - lifecycle trust fixes
  - read-path correctness fixes
  - benchmark harness
  - compatibility/parity documentation
- Confirm `OfficeIMO.Excel` and `OfficeIMO.Excel.Benchmarks` still build cleanly from `master`.
- Confirm no Excel-specific PRs remain open except intentionally deferred parity work.

## Validation Baseline

Run at minimum:

```powershell
dotnet build .\OfficeIMO.Excel\OfficeIMO.Excel.csproj -c Release
dotnet build .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -c Release
dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Excel"
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- --list flat
```

Recommended additional validation:

- run one focused benchmark class to confirm the harness still executes end-to-end
- regenerate the lightweight snapshot baseline for release notes or website updates
- regenerate the write-stage profile if the report export path changed materially
- check the snapshot/profile medians and raw samples before publishing performance claims from a single machine
- recheck `OfficeIMO.Excel\README.md` and `OfficeIMO.Excel\COMPATIBILITY.md`
- recheck [Docs/reviews/officeimo.excel-epplus-review-2026-04-04.md](reviews/officeimo.excel-epplus-review-2026-04-04.md) for any still-open competitive gaps
- verify package description/README metadata render correctly in the `.csproj`
- run `dotnet pack` and confirm the expected `.nupkg` artifact is produced

## Publish Sequence

1. Update `VersionPrefix` in `OfficeIMO.Excel\OfficeIMO.Excel.csproj`.
2. Update any release notes/changelog entries that mention the new Excel baseline.
3. Publish `OfficeIMO.Excel`.
4. Verify the package is visible on NuGet with the expected README/description metadata.

## Non-Goals For This Release

Do not wait for:

- full EPPlus feature parity across every workbook feature
- complete chart/pivot/encryption breadth
- every planned corpus or benchmark scenario

The goal is a stable, honest, and measurable Excel baseline that we can keep improving deliberately.
