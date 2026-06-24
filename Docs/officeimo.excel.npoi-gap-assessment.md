# OfficeIMO.Excel Versus NPOI Assessment

Date: 2026-06-24

This checkpoint compares the current `OfficeIMO.Excel` direction with the active
[`nissl-lab/npoi`](https://github.com/nissl-lab/npoi) project, which is the
current continuation of Tony Qu's NPOI work.

## Short Version

`OfficeIMO.Excel` is going in the right direction if the goal is a modern,
OpenXMLSDK-only, high-level workbook automation library: report generation,
typed reads, tables, templates, charts/pivots as business APIs, preflight,
preservation, and fast package saves.

NPOI is still far ahead as a broad low-level Excel file-format toolkit. It has
true HSSF/XSSF/SXSSF workbook models, shared `NPOI.SS` interfaces, legacy `.xls`
read/write/edit through `HSSFWorkbook`, `.xlsx` read/write/edit through
`XSSFWorkbook`, streaming write through `SXSSFWorkbook`, event-style `.xlsx`
readers, formula evaluation, encryption/decryption infrastructure, tables,
conditional formatting, data validation, drawings, charts, and pivot-table
surfaces.

For legacy `.xls`, our current importer is intentionally much younger. We are
building a model-first BIFF reader and projection path; NPOI already has the
general-purpose HSSF object model. That means we should not claim parity there.
The right target for OfficeIMO is not "clone HSSF quickly"; it is "grow a clean
first-party BIFF model with honest diagnostics, projection, and preservation."

## Source Signals

- [NPOI README](https://github.com/nissl-lab/npoi) says it is the .NET version
  of Apache POI and supports reading and writing Office 2003/2007 files, with
  supported formats including `xls`, `xlsx`, and `docx`.
- [NPOI 2.8.0 release notes](https://github.com/nissl-lab/npoi/discussions/1751)
  list recent Excel work such as XDDF chart support, XLOOKUP formula sync, lazy
  loading, shared-string/style performance work, larger SXSSF I/O buffers, and
  benchmark additions.
- [NPOI 2.7.6 release notes](https://github.com/nissl-lab/npoi/discussions/1684)
  list encryption support for `xls`, `xlsx`, and `xlsm`.
- [Apache POI's HSSF/XSSF limitations](https://poi.apache.org/components/spreadsheet/limitations.html)
  remain relevant for NPOI's ancestry: chart support is limited, macros cannot
  be created, HSSF does not support reading/creating pivot tables, and XSSF pivot
  support is limited.
- [NPOI discussion history](https://github.com/nissl-lab/npoi/discussions/744)
  explicitly says VBA/form-control features are not supported as editable
  features. The source does include macro reader/extractor infrastructure, so
  this should be treated as "inspect/extract/preserve", not "author/edit VBA
  projects."
- Starting with NPOI 2.8.0, the README and release notes describe an additional
  binary EULA / Open Source Maintenance Fee requirement for users who generate
  revenue from NPOI binaries. Benchmark-only local comparison is still a valid
  opt-in use for this repository, but the dependency should not become part of
  normal solution restore/build or any OfficeIMO runtime path.

## Feature Comparison

| Area | OfficeIMO.Excel Current Shape | NPOI Current Shape | Direction |
| --- | --- | --- | --- |
| `.xlsx` create/save | Strong high-level workbook, table, report, template, chart, pivot, fast-package, and preservation-oriented APIs. | Strong low-level XSSF workbook API with broad POI-style surface. | Keep OfficeIMO high-level and ergonomic; do not chase raw XSSF one-for-one. |
| `.xlsx` read | Strong typed/range/DataTable/dictionary reads with friendly binding and streaming slices; feature inspection/preflight is a differentiator. | Broad row/cell object model plus event readers and DataTable/DataSet helpers. | Add NPOI to read benchmarks where equivalent; continue improving ugly-workbook corpus and diagnostics. |
| `.xls` read | In progress: BIFF8 cell/style/formula/comment/hyperlink/validation/conditional-format/filter/external-reference/chart/pivot/drawing diagnostics and projection slices. | Mature HSSF workbook model for binary `.xls`. | NPOI is clearly ahead. Continue our AST/model-first BIFF plan. |
| `.xls` write/edit | Not supported beyond projection from supported import paths into `.xlsx`. | HSSF can create/write/edit `.xls`. | This is the biggest gap. Do not rush it until the read model is stable. |
| Formula calculation | Report-focused evaluator and diagnostics are growing quickly, but not a full Excel-compatible engine. | Larger POI-derived formula evaluator, with many function classes and recent XLOOKUP sync. | NPOI is broader. OfficeIMO should keep report-workflow support plus diagnostics, not promise full Excel calculation yet. |
| Charts | OfficeIMO has broad authored `.xlsx` chart families and dashboard helpers; chart-sheet and Excel-authored mutation remain gaps. | XSSF/XDDF chart support is improving, but POI docs still describe chart support as limited. | OfficeIMO can be better for high-level chart authoring; both need deeper mutation/interop proof. |
| Pivot tables | OfficeIMO has a business-friendly pivot API plus readback/preflight/preservation around slicers/timelines/connections. | XSSF pivot creation/read surfaces exist; HSSF pivots are not broadly supported per POI limitations. | Continue with interop-driven OfficeIMO pivots. NPOI is not a reason to pause. |
| Tables, validation, conditional formatting, filters | OfficeIMO has high-level authoring and partial readback/mutation. | NPOI has broad POI-style low-level APIs in both HSSF/XSSF areas. | Add benchmark scenarios for common equivalent paths; keep OfficeIMO API friendlier. |
| Templates | OfficeIMO has a real template/report workflow with preservation behavior. | NPOI is a workbook toolkit; templates are caller-built patterns. | OfficeIMO advantage. Keep investing. |
| Preservation/preflight | OfficeIMO has explicit feature inspection, unsupported/preserved reports, and corpus baselines. | NPOI often preserves through package/record retention, but preflight is not the same product concept. | OfficeIMO advantage and worth keeping central. |
| Encryption | OfficeIMO supports encrypted OOXML open/save. | NPOI 2.7.6+ reports encryption support for `xls`, `xlsx`, and `xlsm`. | NPOI broader because of `.xls`; OfficeIMO should not expand crypto before core BIFF read/projection is reliable. |
| Macros/VBA/forms | OfficeIMO inspects/preserves macro/form-control package signals in `.xlsx`; legacy `.xls` now reports compound/VBA/OLE signals. | NPOI has macro reader/extractor infrastructure but does not support VBA/form-control editing. | Both are mostly inspect/preserve, not author/edit. OfficeIMO's explicit preflight story is valuable. |
| Rendering/PDF | OfficeIMO has an Excel-to-PDF path for report-shaped workbooks. | NPOI has scratchpad/converter history but not a comparable first-class report PDF story. | OfficeIMO advantage for managed report export, still partial. |

## Missing In Both

- Full-fidelity Excel calculation, especially dynamic arrays, volatile behavior,
  full external workbook calculation, and Excel-identical edge cases.
- Full chart mutation/rendering parity for arbitrary Excel-authored workbooks.
- Full pivot/dashboard/data-model support, including slicers, timelines, pivot
  charts, PowerPivot/data model, and complex cache/query refresh behavior.
- VBA/form-control authoring/editing.
- Exact Excel rendering without Excel.
- Broad, regularly refreshed real-world compatibility corpora.

## Benchmark Recommendation

Add NPOI to benchmarks, but only where the scenario is natural for NPOI:

- `write-cellvalue-*` scalar/string/date/formula workloads using `XSSFWorkbook`.
- `write-datatable-direct`, `write-datareader-plain`, and simple object row
  export using explicit row/cell loops.
- `read-range`, `read-first-column`, `read-datatable`, shared-string read, and
  formula-text read using `WorkbookFactory`/`XSSFWorkbook`.
- A separate `.xls` compatibility comparison lane: OfficeIMO legacy importer
  versus NPOI `HSSFWorkbook` on cell values, styles, formulas-as-text/cached
  values, comments, hyperlinks, validations, conditional formatting, filters,
  and preserved unsupported feature counts.

Do not force NPOI into:

- OfficeIMO-only fluent/template/report APIs.
- Package-copy/preservation scenarios where NPOI would need an artificial
  row-by-row rewrite rather than equivalent package semantics.
- PDF/rendering scenarios.
- Direct DataSet fast-package scenarios that exist specifically to measure
  OfficeIMO's own optimized save path.

Keep NPOI comparison in the opt-in `OfficeIMO.Excel.Benchmarks.NPOI` project so
normal repo restore/build does not pull NPOI unless explicitly requested. The
first lanes cover plain `.xlsx` row/cell write, plain `.xlsx` row/cell read, and
`.xls` row/cell read against the same HSSF-generated bytes. The `.xls`
compatibility lane now also covers formula text/cached value reads and a metadata
read bucket for comments, hyperlinks, merged ranges, and list data validations.
It also covers a conditional-formatting rule read bucket for HSSF-generated
cell-is and formula rules. Because NPOI's shared AutoFilter API is basic
range-setting rather than full criteria editing, the filter comparison lane is
named `xls-read-autofilter-range` and measures the hidden `_FilterDatabase`
range/drop-down signal only. Later lanes can add DataTable/DataSet-style
import/export, styles, richer conditional-formatting style reads, and preserved
unsupported feature counts, then merge the opt-in JSON output into the existing
comparison summary format.

One implementation detail matters for repeatable local evidence: HSSF
comment/drawing reads load SkiaSharp at runtime, while NPOI's package metadata
excludes those runtime assets from its transitive reference. Keep that explicit
reference inside the opt-in benchmark project only; it must not leak into normal
OfficeIMO runtime projects.

## Direction Call

Continue the current OfficeIMO XLS work. We are behind NPOI on breadth, but our
approach is sane: structured BIFF model, diagnostics first, projection where
safe, and explicit preserve-only nodes for everything we cannot yet edit.

For `.xlsx`, we should not try to become a lower-level NPOI clone. OfficeIMO's
better path is to stay friendlier and more workflow-oriented, while borrowing
coverage ideas from NPOI/POI: broader formula evaluator tests, chart/pivot
interop tests, HSSF/XSSF-like corpus categories, and honest benchmark lanes.
