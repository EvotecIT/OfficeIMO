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
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlPagedPurchaseTableBenchmarks*
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlLongDocumentBenchmarks*
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --filter *HtmlStaticStandardsBenchmarks*
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 -- --filter *HtmlProviderParsingBenchmarks*
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 -- --filter *HtmlProviderCssBenchmarks*
```

For a quick harness and allocation smoke, use BenchmarkDotNet's dry job:

```powershell
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net8.0 -- --job Dry
```

## H0 qualification baseline

The H0 baseline-report bundle is an independently authored static report with external CSS, a redistributable font, an SVG image, responsive screen profiles, and paged print rules. Capture the current engine evidence in an empty output directory:

```powershell
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 -- --qualification-baseline --output .benchmark-artifacts/html/qualification/h0-baseline-report
```

The runner verifies every input byte against `manifest.json`, then records the owned DOM and query results, logical and semantic projections, computed-style probes with cascade priority, resource resolution, layout pages, diagnostics, cancellation, provider versions, and source/environment identity. It writes PNG and SVG for each page plus a searchable PDF for the print profile. A missing marker, changed page count, undeclared or unused resource, style mismatch, diagnosed rendering loss, failed PDF readback, or pre-cancellation failure makes the command fail.

Stage timing and allocation values in `baseline.json` are labeled `single-run-observation`. They locate work within one qualification run; use the process-isolated layout evidence and BenchmarkDotNet lanes for repeatable performance comparisons and budgets.

## Coverage

The deterministic corpus measures parsing, computed styles, layout from prepared styles, combined parse/style/layout, Drawing projection, PNG, SVG, and rendered searchable PDF. Output benchmarks cover both ordinary WinAnsi report text and multilingual Unicode text so managed-font fallback costs remain visible. The paged-purchase-table lane adds 250-row and 2,500-row documents with wrapped descriptions, repeated table headers, CSS page furniture, a one-time totals block, and forward-only PDF object serialization to a non-retained destination. The long-document lane uses deterministic 100-page and 1,000-page legal-style packets with page counters and forced article boundaries; each measured layout and PDF operation rejects output whose exact page count differs from the requested corpus. The static-standards lane measures strict two-page layout and tagged PDF output with running elements, row subgrid, clip paths, SVG, page counters, bookmarks, and PDF semantic roles; setup reopens the generated PDF and requires its page and searchable-text contract before timing begins.

Results are comparative evidence for regressions, not universal machine-independent pass/fail thresholds. Correctness stays protected by the end-to-end rendering corpus and focused contracts.

BenchmarkDotNet's allocation totals measure work performed, not maximum live heap. The paged-table lane therefore exposes scaling regressions while the PDF serialization report remains the authoritative evidence for completed page/object payload limits. Whole-document HTML layout is intentionally reported separately and is not described as forward-only.

Capture process-isolated whole-document layout memory without PDF work:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Html.Benchmarks -- --layout-evidence --repeat 3 --json .benchmark-artifacts\html\layout-evidence.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Html.Benchmarks -- --layout-verify-budgets --repeat 1
```

Each child validates the expected page and text-marker contract, then records
allocation, retained managed heap, sampled managed-heap peak, process peak,
input bytes, page count, and rendered text characters. The checked-in
`html-layout-performance-budgets.json` file gates allocations and memory growth
for all six workloads. Elapsed and absolute process-peak limits are deliberately
looser gross-regression guards; they are not portable throughput claims.

Capture every page's output fingerprint separately from timing when changing layout allocation:

```powershell
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Html.Benchmarks -- --layout-fingerprint Purchase2500 > layout-fingerprint.json
```

The report records page dimensions and SVG hashes, the complete text hash, and diagnostics.
Compare it with the same workload on the previous revision so lower allocation cannot hide a
missing page or changed content. Fingerprint collection is outside the timed measurement.

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

## Provider decision evidence

The provider lanes separate native AngleSharp parsing, the OfficeIMO-owned document
projection, the lazy conversion-document path, raw AngleSharp.Css syntax parsing,
OfficeIMO's lossless CSS syntax result, and the complete OfficeIMO cascade. Raw syntax
parsers and the complete cascade perform different work; compare each lane with itself
across revisions rather than treating their elapsed-time ratio as parser overhead.

Capture process-isolated elapsed, allocation, retained managed heap per result,
provider assembly identity, and platform information:

```powershell
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Html.Benchmarks -- --provider-evidence --repeat 5 --json .benchmark-artifacts/html/provider-evidence.json
```

The 100-row HTML scenarios all validate the same 420-element recovered tree and
must produce one exact provider-neutral structural fingerprint across every repeat
and all four HTML lanes. The 100-rule CSS scenarios validate the provider stylesheet
and the owned 305-element
cascade separately. Retained-heap values use eight simultaneously retained results
inside a fresh child process and are diagnostic observations rather than hard budgets.
Run the same command on each platform and compare medians only for equivalent source,
runtime, architecture, and commit.

## Owned document and CSS syntax budgets

The owned API gate measures parse, query, edit, serialization, conversion after owned
document access, lossless CSS syntax, selected typed property grammar, contextual length
math, owned selector lists, structural and logical pseudo-classes, stylesheet namespaces,
nested qualified rules and interleaved nested declarations,
the normal managed CSS cascade, and the same cascade with opt-in traces. General operations
run at 10, 100, and 1,000 rows; the advanced selector lane runs at 100, 1,000, and 6,000
siblings so its largest case crosses the former default-budget failure boundary. Every
operation validates its structure or exact source before evidence is accepted.
Query and serialization clear the short provider projection lease before timing, so they
include reconstruction after memory pressure. The parse and conversion lanes retain their
result through a compacting collection, which makes duplicate-graph regressions visible.

```powershell
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Html.Benchmarks -- --owned-document-evidence --repeat 3 --json .benchmark-artifacts/html/owned-document.json
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Html.Benchmarks -- --owned-document-verify-budgets --repeat 3
```

`html-owned-document-performance-budgets.json` supplies cross-platform regression ceilings
for median elapsed time, median allocation, median retained heap, maximum sampled managed
heap, process peak, and output length. The cancellation lane starts 10,000-, 25,000-, and 100,000-row parses with a live token, signals from inside the selected provider's parse boundary, then cancels from a synchronized worker and requires
exit within its declared latency budget. These ceilings guard the declared workloads and do not
claim universal throughput or full CSS conformance. Compare `CssCascade` with
`CssCascadeTrace` at the same scale to review the cost of retaining the selected trace slice;
the default computation explicitly verifies that no trace graph was retained. The advanced
selector and nested-rule lanes additionally enforce Normal-to-Large elapsed and allocation
ratios for six times as many siblings or rules. The CSS corpus also verifies typed contextual dimensions, constant
numeric calculations, functional colors, attribute matching with an ASCII-insensitive
modifier, grouped selectors, namespace matching, and structural and logical pseudo-classes.
