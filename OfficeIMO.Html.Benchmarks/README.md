# OfficeIMO.Html.Benchmarks

`OfficeIMO.Html.Benchmarks` measures the shared HTML renderer and its first-party Drawing and PDF projections. It is a non-packable developer project and adds no dependency to shipped OfficeIMO packages. BenchmarkDotNet is already used by the repository's benchmark projects.

## Run

Run the complete suite from the repository root:

```powershell
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0
```

Run the stage or output lanes separately:

```powershell
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlRenderingStageBenchmarks*
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlRenderingOutputBenchmarks*
```

For a quick harness and allocation smoke, use BenchmarkDotNet's dry job:

```powershell
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --job Dry
```

## Coverage

The deterministic corpus measures parsing, computed styles, layout from prepared styles, combined parse/style/layout, Drawing projection, PNG, SVG, and rendered searchable PDF. Output benchmarks cover both ordinary WinAnsi report text and multilingual Unicode text so managed-font fallback costs remain visible.

Results are comparative evidence for regressions, not universal machine-independent pass/fail thresholds. Correctness stays protected by the end-to-end rendering corpus and focused contracts.

## Review budgets

Use these allocation ceilings as regression-review budgets for the deterministic corpus. They deliberately leave headroom above the July 2026 net8 reference run; timing should be compared against the same machine's previous healthy commit and reviewed when a lane exceeds 2x its baseline mean.

| Document class | Parse | Styles | Prepared layout | Parse/style/layout |
| --- | ---: | ---: | ---: | ---: |
| Small report, 10 rows | 0.5 MB | 3 MB | 6 MB | 10 MB |
| Standard report, 100 rows | 2 MB | 15 MB | 40 MB | 60 MB |

| Standard 40-row output | Allocation ceiling |
| --- | ---: |
| Drawing projection | 2 MB |
| SVG | 4 MB |
| PNG at scale 1 | 64 MB |
| Searchable PDF, WinAnsi text | 32 MB |
| Searchable PDF, multilingual Unicode text | 256 MB |

These are review triggers, not flaky unit-test assertions. A change may intentionally exceed one when the corpus or fidelity contract grows, but the new baseline and reason should be recorded in the change.
