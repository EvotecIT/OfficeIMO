# Optional PowerPoint library comparison

This opt-in project compares OfficeIMO.PowerPoint with
[ShapeCrawler 0.82.0](https://www.nuget.org/packages/ShapeCrawler/0.82.0)
(MIT) and the MIT-licensed Open XML SDK 3.5.1. NPOI 2.8.1 is not included: its
.NET packages do not ship the XSLF PowerPoint user model.

## Warm in-process comparison

BenchmarkDotNet classes measure common deck workloads at 20 and 200 slides:

| Workload | Measured operation | Correctness proof |
| --- | --- | --- |
| Create | Create a deck where every slide has a bold title, a three-paragraph body, a 4x3 table, and a PNG picture, then serialize it | Every title, body paragraph, cell, and picture byte payload, plus Open XML validation |
| Read | Load the same deck and read every text box and table cell | Order-independent count, character count, and checksum of all text items |
| Edit | Load the deck, change every title and one table cell per slide, and save | Edited values, untouched content, and Open XML validation |
| Copy slides | Load the deck, append a copy of every slide, and save | Doubled slide count with copied content and picture payloads, plus Open XML validation |

Read, edit, and copy start from the same deck authored by the Open XML SDK
lane, so no high-level library reads its own output. The Open XML SDK create
lane starts from a zero-slide OfficeIMO package that supplies the slide master,
layouts, and theme, because writing that scaffolding by hand is not a
representative SDK workload. Each library uses its default image handling; the
SDK lane writes one image part per slide. Validation runs in global setup,
outside the measured operations.

```powershell
dotnet run -c Release --project .\OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler -- validate
dotnet run -c Release --project .\OfficeIMO.PowerPoint.Benchmarks.ShapeCrawler -- bench --filter '*PowerPoint*ComparisonBenchmarks*'
```

Charts are compared only by the cold-probe lane below; a hand-written Open XML
SDK chart with an embedded workbook is not an equivalent workload.

## Cold-probe comparison

The cold-probe runner covers the two PowerPoint operations
that OfficeIMO and ShapeCrawler can perform with equivalent editable output:

- create and save a presentation;
- open, edit, and save an existing presentation.

It is deliberately excluded from `OfficeIMO.slnx`. Running it restores
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
