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
| Public benchmark evidence | Partial | Committed benchmark snapshots plus write/read profiles now exist. Recent local snapshots show OfficeIMO ahead of ClosedXML on covered write/read/load-edit-save workloads, including the refreshed 25,000-row report-export profile, but broader row-count and hardware coverage is still needed before making large-workbook claims. |

## Current Strengths

- ergonomic write APIs for reports, tables, and object insertion
- deterministic save/repair behavior compared with raw Open XML usage
- typed and editable read surfaces that are pleasant for application code
- practical stream support for service and automation scenarios

## Highest-Priority Gaps

1. Keep reducing report-export overhead, with table creation, auto-fit, and save costs now as important as `InsertObjects()` on the 25,000-row profile.
2. Publish and refresh benchmark result sets on stable hardware instead of relying on a single developer machine snapshot.
3. Expand the corpus with messy, externally-authored workbooks and round-trip assertions.
4. Formalize formula/recalculation expectations so users know what is computed versus preserved.
5. Decide whether encryption/password support is a roadmap item or a deliberate non-goal.
6. Keep chart/pivot expansion driven by parity tests rather than isolated feature requests.

## Latest Snapshot Highlights

- Updated 5-sample end-to-end snapshot on 2026-05-12: `OfficeIMO.Excel` averages `114.9 ms` for the 2,500-row report scenario, while `ClosedXML` averages `271.8 ms`.
- The 2026-05-13 local Release read comparison records build configuration, medians, and raw samples against `ClosedXML`, current `EPPlus`, and isolated `EPPlus 4.5.3.3`. Current-library scenarios are now measured in rotated groups so fixed OfficeIMO/ClosedXML/EPPlus ordering does not decide the result. `OfficeIMO.Excel` is ahead of `ClosedXML` on every covered read scenario by average in that snapshot, but current and legacy EPPlus still lead the full dense range scans.
- The same 2026-05-13 Release read comparison shows the strongest OfficeIMO shape on bounded reads: first-100-row range read averaged `4.4 ms` versus `96.4 ms` for ClosedXML, `33.8 ms` for EPPlus, and `27.8 ms` for EPPlus 4.5.3.3; first-100-row streaming averaged `3.5 ms` versus `66.9 ms`, `29.3 ms`, and `28.0 ms`.
- Typed `ReadObjects<T>()` still leads ClosedXML in the refreshed comparison (`59.4 ms` versus `73.5 ms`), but current and legacy EPPlus lead this scenario at about `30.9 ms` and `31.9 ms`; this remains a priority parity target.
- Full 2,500-row dense reads are now honest parity targets rather than solved claims: `ReadRange()` averaged `64.6 ms` versus `113.3 ms` for ClosedXML, `44.8 ms` for EPPlus, and `34.8 ms` for EPPlus 4.5.3.3; `ReadRangeStream()` averaged `46.3 ms` versus `84.3 ms`, `37.5 ms`, and `34.4 ms`.
- Load/edit/save remains ahead in the refreshed snapshot: `OfficeIMO.Excel` averaged `108.4 ms` versus `ClosedXML` at `124.7 ms`.
- The latest 2,500-row write profile dated 2026-05-12 shows the current staged cost shape with report-export AutoFit saves deferred to the document boundary: `InsertObjects()` is about `16.8 ms`, `AddTable()` is about `27.8 ms`, `AutoFitColumns()` is about `25.4 ms`, and the OfficeIMO staged write total is about `99.4 ms` versus ClosedXML at about `250.4 ms`.
- A 25,000-row write profile on 2026-05-12 shows the larger-report shape after the row-major append, appended-cell style-cache, bulk shared-string registration, table range-scan, contiguous table-range verification, column-reference caching, auto-fit planning, auto-fit shared-string text/run caching, auto-fit style/number-format cache fast paths, and deferred AutoFit worksheet saves: `InsertObjects()` is about `211.4 ms`, `AddTable()` is about `153.1 ms`, `AutoFitColumns()` is about `169.0 ms`, and the OfficeIMO staged write total is about `720.7 ms` versus ClosedXML at about `970.4 ms`. The AutoFit profile records `BuildPlan`, `CalculateWidths`, and `ApplyWidths` sub-stages.
- The benchmark artifacts now include raw samples and medians so outliers are visible instead of hidden behind a single average.

## Evidence In Repo

- Benchmark harness: [`OfficeIMO.Excel.Benchmarks`](../OfficeIMO.Excel.Benchmarks/)
- Benchmark artifacts: [`../Docs/benchmarks/README.md`](../Docs/benchmarks/README.md)
- Compatibility corpus tests: [`../OfficeIMO.Tests/Excel.CompatibilityCorpus.cs`](../OfficeIMO.Tests/Excel.CompatibilityCorpus.cs)
- Excel package README: [`README.md`](README.md)

## Suggested Release-Prep Checks

1. Build `OfficeIMO.Excel` and the benchmark harness on the target SDKs.
2. Run the Excel-focused test slice before any release candidate.
3. Run at least one benchmark class in `Release` to catch accidental performance cliffs.
4. Update this matrix whenever a major Excel feature is added or a parity gap is closed.
