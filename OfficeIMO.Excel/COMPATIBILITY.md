# OfficeIMO.Excel Compatibility And Competitive Matrix

This document tracks where `OfficeIMO.Excel` is already strong, where it is only partially competitive, and where it still needs explicit parity work before it can credibly compete with mature Excel libraries such as EPPlus.

It is intentionally honest. "Partial" means usable, not "done".

## Current Matrix

| Area | Status | Notes |
| --- | --- | --- |
| Workbook create/load/save | Supported | File and stream workflows are available, including normalized recovery for some malformed content-type packages. |
| Typed range reads | Supported | `RowsAs<T>()`, `ReadObjects<T>()`, `ReadRangeAs<T>()`, and dictionary/DataTable reads are available, with friendly header-to-property matching, explicit aliases via `DisplayName`, `DataMember(Name=...)`, and `ExcelColumn`, plus diagnostics for ambiguous typed mappings. |
| Editable row workflows | Supported | `RowsObjects()` and read bridges support practical read-modify-save flows. |
| Header handling | Supported | Duplicate, normalized-colliding, blank, and generated-fallback-vs-explicit headers now disambiguate deterministically, and header-map lookups stay fresh after in-memory edits. |
| Number/date import fidelity | Partial | Common formats work well; custom formats are better after the token-aware classifier, but the corpus still needs to grow. |
| Tables and named ranges | Supported | Table creation and naming safeguards are in place, with worksheet/global named range helpers. |
| Auto-fit and report ergonomics | Supported | Auto-fit, object insertion, and table/report helpers are a current strength. |
| Charts | Partial | Common chart authoring is present, but breadth and round-trip parity are still behind top-tier competitors. |
| Pivot tables | Partial | Basic support exists, but this is not yet a broad parity surface. |
| Formula/recalculation story | Partial | Formula authoring exists, but the package does not yet present a first-class recalculation/value-engine story comparable to EPPlus expectations. |
| Worksheet/workbook protection | Partial | Protection helpers exist, but broader permission fidelity and compatibility proof are still needed. |
| Encryption/password support | Roadmap | This remains a notable gap for enterprise-style workbook scenarios. |
| Streaming for very large workbooks | Partial | In-memory/file/stream workflows are strong, but the package still needs clearer large-workbook guidance and published benchmarks. |
| Import fidelity for ugly real-world workbooks | Partial | Correctness has improved, but corpus depth is still lighter than it should be for competitive claims. |
| Public benchmark evidence | Partial | Committed benchmark snapshots and write-stage profiles now exist. Recent optimization work materially improved report-style write performance, but OfficeIMO still trails ClosedXML on that workload. |

## Current Strengths

- ergonomic write APIs for reports, tables, and object insertion
- deterministic save/repair behavior compared with raw Open XML usage
- typed and editable read surfaces that are pleasant for application code
- practical stream support for service and automation scenarios

## Highest-Priority Gaps

1. Keep reducing report-export overhead, with `InsertObjects()` now the largest staged cost and the remaining variance concentrated in occasional outlier samples rather than the steady-state path.
2. Publish and refresh benchmark result sets on stable hardware instead of relying on a single developer machine snapshot.
3. Expand the corpus with messy, externally-authored workbooks and round-trip assertions.
4. Formalize formula/recalculation expectations so users know what is computed versus preserved.
5. Decide whether encryption/password support is a roadmap item or a deliberate non-goal.
6. Keep chart/pivot expansion driven by parity tests rather than isolated feature requests.

## Latest Snapshot Highlights

- Updated 5-sample end-to-end snapshot on 2026-04-05: `OfficeIMO.Excel` now averages `258.0 ms` for the 2,500-row report scenario, while `ClosedXML` averages `273.3 ms`.
- Read scenarios are still in the same general band as the comparison run: `ReadObjects()` averaged `118.3 ms`, `ReadRangeAsDataTable()` averaged `139.1 ms`, and the comparable `ClosedXML` row iteration averaged `99.9 ms`.
- Load/edit/save remains a strength in the refreshed snapshot: `OfficeIMO.Excel` averaged `123.6 ms` versus `ClosedXML` at `149.9 ms`.
- The latest write profile dated 2026-04-05 shows the current staged cost shape after the conservative auto-fit fast path and array-backed `InsertObjects()` buffers: `InsertObjects()` is about `179.3 ms`, `AddTable()` is about `28.3 ms`, `AutoFitColumns()` is about `88.2 ms`, and the OfficeIMO staged write total is about `318.6 ms`.
- The benchmark artifacts now include raw samples and medians so outliers are visible instead of hidden behind a single average.

## Evidence In Repo

- Benchmark harness: [`OfficeIMO.Excel.Benchmarks`](../OfficeIMO.Excel.Benchmarks/)
- Benchmark artifacts: [`../Docs/benchmarks/README.md`](../Docs/benchmarks/README.md)
- Compatibility corpus tests: [`../OfficeIMO.Tests/Excel.CompatibilityCorpus.cs`](../OfficeIMO.Tests/Excel.CompatibilityCorpus.cs)
- EPPlus-focused review: [`Docs/reviews/officeimo.excel-epplus-review-2026-04-04.md`](../Docs/reviews/officeimo.excel-epplus-review-2026-04-04.md)
- Excel package README: [`README.md`](README.md)

## Suggested Release-Prep Checks

1. Build `OfficeIMO.Excel` and the benchmark harness on the target SDKs.
2. Run the Excel-focused test slice before any release candidate.
3. Run at least one benchmark class in `Release` to catch accidental performance cliffs.
4. Update this matrix whenever a major Excel feature is added or a parity gap is closed.
