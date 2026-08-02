# Optional ShapeCrawler comparison

This project provides an opt-in comparison for the two PowerPoint operations
that OfficeIMO and ShapeCrawler can perform with equivalent editable output:

- create and save a presentation;
- open, edit, and save an existing presentation.

It is deliberately excluded from `OfficeIMO.sln`. Running it restores
ShapeCrawler and its transitive dependencies, but normal OfficeIMO restore,
build, test, package, and benchmark paths remain unchanged.

The comparison uses the same deterministic `Small`, `Normal`, and `Large`
slide counts, slide dimensions, background/style pattern, editable text,
vector panels, tables, two-series clustered bar charts, and every-tenth-slide
edit cadence. It validates every produced package by reopening it, checking
slide and shape counts, and running the Open XML validator after timing and
peak-working-set capture. Probes are intentionally cold. Image and PDF export are not compared because the libraries do not
expose equivalent rendering contracts.

Run it independently from the repository root:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler -c Release -- --scale Normal --json artifacts/shapecrawler-normal.json
```

Run the OfficeIMO baseline on the same machine and runtime configuration:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks -c Release -- --scale Normal --json artifacts/officeimo-normal.json
```

Compare elapsed time, managed allocation, peak working set, package size, and
the validated slide/shape counts. These are baselines, not regression budgets.
Collect several runs on a quiet machine before defining thresholds, and compare
like-for-like operations and scales rather than the rendering-only OfficeIMO
lanes.
