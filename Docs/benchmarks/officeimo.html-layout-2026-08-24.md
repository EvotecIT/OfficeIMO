# OfficeIMO.Html non-PDF layout evidence — 2026-08-24

## Scope

This run measures complete in-memory HTML layout without PDF serialization or
raster image work. Every isolated child retains the complete
`HtmlRenderDocument` and validates required text markers, page counts for the
long-document workloads, and the strict no-loss contract for the static
standards workload.

No contender ratio is claimed. A general-purpose comparison library has not
been identified that produces the same paged, vector, searchable,
semantics-bearing render model. The evidence therefore compares the optimized
implementation with its clean pre-optimization commit and adds absolute
regression budgets.

## Environment and method

- Windows 11, x64, .NET 10
- Three fresh child processes per workload
- Baseline commit: `f0df26bb8`
- Optimized commit and budget manifest: `cd8a45c07`
- Both commits were clean and used the same workload corpus, validation, and
  machine. The optimized evidence runner stops its memory sampler before text
  validation; the baseline sampler included validation. Elapsed time and
  allocation retain identical render-only boundaries, while the managed-peak
  before/after values are directional rather than a pure product-code delta.
- Elapsed time and allocation cover rendering only. Retained and peak-memory
  measurements use a second render in the same isolated child; validation runs
  after the sampler stops so text extraction does not inflate the layout peak.

## Results

| Workload | Pages | Elapsed before → after | Allocation before → after | Observed managed peak before → after | Retained after | Process peak before → after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Report100 | 1 | 76.06 → 40.10 ms (-47.3%) | 37.68 → 21.02 MiB (-44.2%) | 39.56 → 21.10 MiB (-46.7%) | 0.58 MiB | 110.21 → 110.09 MiB |
| Purchase250 | 43 | 307.86 → 197.22 ms (-35.9%) | 132.71 → 66.68 MiB (-49.8%) | 51.99 → 44.97 MiB (-13.5%) | 1.66 MiB | 127.02 → 126.55 MiB |
| Purchase2500 | 418 | 1,451.56 → 906.29 ms (-37.6%) | 1,216.19 → 645.31 MiB (-46.9%) | 202.65 → 164.00 MiB (-19.1%) | 16.48 MiB | 348.89 → 306.88 MiB |
| Long100 | 100 | 150.71 → 97.37 ms (-35.4%) | 51.81 → 42.01 MiB (-18.9%) | 46.85 → 42.17 MiB (-10.0%) | 0.73 MiB | 116.41 → 112.12 MiB |
| Long1000 | 1,000 | 588.20 → 596.89 ms (+1.5%) | 482.98 → 393.02 MiB (-18.6%) | 82.83 → 82.22 MiB (-0.7%) | 7.27 MiB | 171.25 → 162.54 MiB |
| StaticStandards | 2 | 13.13 → 8.47 ms (-35.4%) | 3.62 → 2.43 MiB (-32.8%) | 3.69 → 2.47 MiB (-33.2%) | 0.02 MiB | 71.83 → 68.26 MiB |

The managed-peak baseline includes validation and is therefore not used to
attribute the full observed peak reduction to product code. The optimized
absolute peak and retained values are measured with validation outside the
sampler and are the regression-budget source.

The 1,000-page elapsed result is effectively unchanged within process and host
noise, while its deterministic allocation fell by 18.6%. Every other workload
improved materially in both elapsed time and allocation.

## Changes that moved the result

- CSS-backed resource discovery now computes the full cascade only when
  `url()`, `var()`, or `image-set()` resource resolution needs it.
- Computed-style construction owns already-created dictionaries and sets,
  avoids a second property dictionary when no custom-property substitution is
  needed, and no longer materializes inherited cascade objects for every child.
- Selector candidates are reused for pseudo-element evaluation, and pseudo
  style dictionaries are created only after a pseudo rule actually matches.
- Parsed rule declarations validate supported values once instead of repeating
  the same border, color, and length validation for every matching element.
- Semantic visual translation preserves known ordering without repeated LINQ
  sorts and intermediate copies.
- `OfficeShape.Clone()` reuses its immutable, read-only point and path-command
  snapshots instead of recopying vector geometry at every translation stage.

## Regression gates and remaining work

`html-layout-performance-budgets.json` covers all six workloads. The .NET 10
repeat-three run and a .NET 8 repeat-one run passed. Allocation, retained
managed growth, and managed peak are hard gates. Elapsed and absolute process
peak are looser gross-regression guards because they are host-sensitive.

The 2,500-row table still allocates approximately 645 MiB. That is much better
than 1.2 GiB, but it is not the desired end state. Table border/style parsing,
large render-box objects, and repeated visual translation remain the next
profiling targets. The 1,000-page lane's approximately 393 MiB allocation is
also an explicit optimization target. These budgets preserve the improvement;
they do not redefine the current numbers as optimal.

Output-file size is not applicable to this lane because it returns an in-memory
render model. PDF and image artifact size belong to their separate owners and
benchmark suites.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Html.Benchmarks -- --layout-verify-budgets --repeat 3 --json .benchmark-artifacts\html\layout-evidence.json
```
