# Optional ShapeCrawler comparison

This project provides an opt-in comparison for the two PowerPoint operations
that OfficeIMO and ShapeCrawler can perform with equivalent editable output:

- create and save a presentation;
- open, edit, and save an existing presentation.

It is deliberately excluded from `OfficeIMO.sln`. Running it restores
ShapeCrawler and its transitive dependencies, but normal OfficeIMO restore,
build, test, package, and benchmark paths remain unchanged.

The comparison uses the exact same prebuilt `Small`, `Normal`, and `Large`
PPTX packages for both open/edit/save lanes. Create/save retains equivalent
slide counts, dimensions, background/style patterns, editable text, vector
panels, tables, two-series clustered bar charts, and edit cadence because each
library must author its own package. It validates every produced package by reopening it, checking
slide and shape counts, and running the Open XML validator after timing and
peak-working-set capture. Probes are intentionally cold. Image and PDF export are not compared because the libraries do not
expose equivalent rendering contracts.

First run the OfficeIMO baseline from the repository root to create the shared
input corpus and measure OfficeIMO against those exact files:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks -c Release -- --scale Normal --corpus-dir artifacts/powerpoint-shared-corpus --json artifacts/officeimo-normal.json
```

Then run ShapeCrawler against the same corpus directory:

```powershell
dotnet run --project OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler -c Release -- --scale Normal --corpus-dir artifacts/powerpoint-shared-corpus --json artifacts/shapecrawler-normal.json
```

Use `--operation CreateSave|OpenEditSave` to isolate one contract and
`--repeat 3` (or more) to collect repeated cold probes over the same shared
corpus.

Compare elapsed time, managed allocation, peak working set, package size, and
the validated slide/shape counts. These are baselines, not regression budgets.
Collect several runs on a quiet machine before defining thresholds, and compare
like-for-like operations and scales rather than the rendering-only OfficeIMO
lanes.
