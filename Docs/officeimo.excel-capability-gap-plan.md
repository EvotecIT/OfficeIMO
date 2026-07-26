# OfficeIMO.Excel Capability Gap And Closure Plan

Date: 2026-07-25

This assessment compares `OfficeIMO.Excel` with a mature managed spreadsheet API and with current public demand signals from spreadsheet-library issue and pull-request backlogs. The comparison is intentionally capability-first: it describes the workflows users need, not another product's API shape.

The source review used `OfficeIMO` commit `14fd17dbd`, the current public package documentation, the latest stable external release available on the review date, and the external development branch at commit `98a3d960`.

## Executive Conclusion

`OfficeIMO.Excel` is already ahead of the normal managed spreadsheet baseline in several valuable areas: broad chart authoring, pivot grouping and calculated fields, typed and streaming data paths, `.xls` and `.xlsb` support, template generation, threaded comments, workbook preflight, unsupported-part preservation, rendering, and security-oriented diagnostics.

The main parity gap is a fully general mutable worksheet model. Users can build rich workbooks, but they do not yet have one public structural-edit contract for inserting, deleting, moving, and copying rows, columns, cells, and ranges while every dependent workbook structure is updated safely.

That missing foundation affects several apparently separate gaps:

- Table resize and schema mutation.
- Formula, array-formula, shared-formula, and defined-name rewrites.
- Conditional-formatting and data-validation range rewrites.
- Drawing, chart, sparkline, comment, and hyperlink movement.
- Pivot source, print area, print title, AutoFilter, and merge updates.
- Reliable copy, move, transpose, and template mutation outside specialized template paths.

The recommended sequence is therefore structural integrity first, common worksheet workflows second, and advanced dashboard interoperability third. Adding isolated convenience methods before the mutation foundation would create several competing implementations of reference rewriting.

## Current Implementation Status

The first structural slice now provides public row insertion and deletion on
`ExcelSheet`. One workbook-wide path updates cell, validation, conditional
formatting, table, chart, sparkline, hyperlink, and defined-name formulas, then
remaps row-bound package structures. Shared formulas are materialized into
equivalent member formulas before cells move so cross-sheet boundaries remain
correct. Unsafe array-formula and PivotTable output boundaries, table
header/totals-row deletion, and dependent-reference overflow are rejected before
mutation. Calculation-chain metadata is invalidated and the workbook requests a
full recalculation.

This is the row foundation, not completion of Milestone 1. Column and cell-shift
operations should generalize the same reference and metadata transforms rather
than introduce parallel implementations.

## Status Definitions

| Status | Meaning |
| --- | --- |
| Strong | A first-class public workflow exists and is supported by meaningful tests or documentation. |
| Partial | Useful functionality exists, but the public model, breadth, readback, or interoperability is incomplete. |
| Preserve/inspect | The package identifies or round-trips the feature more reliably than it authors or edits it. |
| Missing | No coherent first-class public workflow was found during the source and test inventory. |

## Capabilities Already At Or Beyond The Target

These should remain regression-protected, documented, and benchmarked. They should not consume parity work merely to imitate a different object model.

| Capability | Current position | Recommended action |
| --- | --- | --- |
| Typed values and bulk data | Strong object, dictionary, `DataTable`, `DataSet`, CSV/JSON, typed-row, `DateOnly`, `TimeOnly`, stream, and fast package paths. | Keep the fast path distinct from the mutable workbook model and publish workload-specific benchmarks. |
| File-format breadth | Managed `.xlsx`, `.xls`, and `.xlsb` workflows plus macro and embedded-package handling. | Expand real-file corpus tests; do not narrow the package around OOXML-only assumptions. |
| Charts and reports | Broad classic chart creation and update, modern report recipes, dashboard helpers, images, sparklines, and report rendering. | Focus new chart work on native modern chart structures, deeper read/edit, and fidelity—not basic chart presence. |
| Pivot reporting | Creation/readback, date and number grouping, filters, show-values-as, calculated fields, styles, cache metadata, and source updates. | Deepen native dashboard interoperability and externally backed cache behavior. |
| Formulas | Normal, shared, and array-formula support; dependency inspection; dirty/cached-result policies; a sizeable report-oriented evaluator; custom function extension. | Consolidate formula parsing and reference rewriting before expanding evaluator breadth. |
| Formatting | Practical fonts, fills, gradients, borders, number formats, alignment, protection, row/column sizing, conditional formatting, and page setup. | Add composable style objects, named styles, and full read/edit coverage where needed. |
| Images and comments | Rich drawing anchors, range anchoring, crop/rotate/flip/alt text, header/footer images, legacy notes, and editable threaded conversations. | Add native in-cell images and route reusable image decoding/color work through `OfficeIMO.Drawing`. |
| Templates | Repeating rows/sheets, optional sections, typed binding, image binding, formula/name/table rewrites, and preservation of many relationship-backed structures. | Reuse the structural-edit engine so template mutation and general mutation have one owner. |
| Safety and diagnostics | Feature inspection, preflight, preservation reporting, package validation, encryption, signatures/macros inspection, and repair-oriented workflows. | Make dry-run mutation impact analysis the next differentiator. |
| Export | PDF, HTML, and image-oriented report paths exist beyond the normal workbook-only baseline. | Improve fidelity against a defined report corpus rather than broadening without proof. |

## Prioritized Gap List

### P0 — Foundational Gaps

| Capability | Current state | Missing public contract | Completion evidence |
| --- | --- | --- | --- |
| Structural row, column, and cell edits | Specialized template paths can shift or repeat content, but no general insert/delete engine was found. | Insert/delete rows, columns, and cells with shift direction; update every affected reference and relationship; return a mutation report. | A dependency-matrix test suite proves formulas, names, tables, filters, validations, conditional formatting, merges, hyperlinks, comments, drawings, sparklines, pivots, and print metadata remain valid after each edit. |
| Reference parsing and rewriting | Several formula, template, chart, and dependency paths already interpret references. | One reusable syntax tree and rewrite service for A1, R1C1, sheet-qualified, 3-D, external, structured, name, union, intersection, spill, shared-formula, array-formula, and print-area references. | All consumers use the shared service; round-trip and mutation fixtures cover quoted sheet names, escaping, absolute/mixed references, unions, and external references. |
| Copy, move, and transpose | Useful range builders and specialized copy behavior exist, but no complete general workflow was found. | Copy/move ranges with explicit policies for formulas, styles, sizes, merges, validation, conditional formatting, comments, hyperlinks, and drawings; transpose values and formulas. | Cross-sheet and cross-workbook tests demonstrate deterministic collision, naming, and relationship policies. |
| External workbook compatibility corpus | Strong feature inspection and targeted preservation tests exist. | A versioned corpus for Excel-, LibreOffice-, Google Sheets-, and generated workbooks, including malformed and unusual external references. | Corpus runs never crash silently: files are edited, preserved, or rejected with actionable findings; before/after package differences are recorded. |

### P1 — High-Value Worksheet Workflows

| Capability | Current state | Missing public contract | Completion evidence |
| --- | --- | --- | --- |
| AutoFilter evaluation and editing | Text, comparison, and between filters plus range sorting are available. | Date groups, dynamic periods, top/bottom items and percentages, font/fill/icon filters, blank policies, reapply, row visibility, and complete filter-state readback/editing. | Filters authored by OfficeIMO and Excel round-trip with the same visible rows and criteria. |
| Table mutation | Creation, totals, styles, sorting, column lookup, and append-oriented helpers exist. | Resize, replace data, append typed records through the table object, rename/reorder/add/remove fields, toggle header/totals/filter rows, and rewrite structured references. | Schema and resize tests cover formulas, names, validations, pivots, charts, and empty/single-row tables. |
| Search and range algebra | Basic first-match and replace workflows exist. | Find-all over values, displayed text, and formulas; predicate search; union, intersection, difference, surrounding/grow/shrink, relative ranges, and visible-cell selection. | Public examples cover audit, cleanup, selective formatting, and filtered-data workflows without raw Open XML. |
| Formula authoring and calculation semantics | A broad bounded evaluator, dependency graph, cached values, shared formulas, and array formulas exist. | R1C1 authoring/conversion, dynamic-array/spill metadata, data-table formulas, iterative/circular calculation policy, and selected missing function families. | A documented calculation contract distinguishes authored, evaluated, cached, deferred-to-Excel, and unsupported formulas. Fixtures compare cached results and recalculation flags. |
| Allowed edit ranges and ignored errors | Workbook/sheet protection and workbook write-reservation/read-only-recommendation workflows exist; lower-level package support covers protected ranges and ignored errors. | Format-neutral high-level allowed-edit-range and ignored-error-region APIs. | Excel-authored metadata can be read, modified, removed, and recreated without losing unrelated protection or write-reservation settings. |
| Typed themes and named styles | Raw theme XML management and rich direct cell formatting exist. | Typed theme colors/fonts/effects, named style creation/application/readback, style snapshots/equality, and efficient style reuse. | Theme and named-style fixtures preserve theme references rather than flattening every style to literal values. |
| Sparkline lifecycle | Sparkline creation and several visual options exist. | Read/edit/remove groups, date axes, hidden/empty-cell behavior, individual/group min/max axes, right-to-left, line weight, and copy/move semantics. | Excel-authored sparkline groups round-trip and respond correctly to structural edits. |
| Complete conditional-format styling | A broad set of rule types and management APIs exists. | Full differential styles, data-bar gradient/border/axis/negative-value settings, custom icons, and complete read/edit coverage. | Rule settings round-trip from Excel-authored fixtures and remain correct after range edits. |
| Workbook and sheet view state | Freeze panes, grid/headings/zero values, zoom, right-to-left, page setup, print areas/titles, breaks, and headers/footers exist. | Active cell and multi-selection, split panes and top-left state, typed workbook defaults, formula-based print areas, print-comments/errors, DPI, draft/black-and-white, and first-page-number workflows. | View and print fixtures reopen with the intended selection, pane, and print behavior. |

### P2 — Advanced Interoperability And Differentiation

| Capability | Current state | Missing public contract | Completion evidence |
| --- | --- | --- | --- |
| Native slicers and timelines | Detection, preservation, and OfficeIMO-owned metadata exist; native authoring is incomplete. | Native slicer/timeline caches, relationships, layout, selection state, pivot/table bindings, read/edit/remove, and cache sharing. | Excel opens generated dashboards without repair and interactions filter the intended pivots/tables. |
| Deeper pivot interoperability | Pivot creation and common reporting behavior are strong. | Calculated items/members, region-specific pivot styles, shared-cache lifecycle, refresh/materialization choices, and robust external/query-backed source mutation. | A multi-pivot corpus covers shared caches, grouping, refresh, source resize, calculated members, and query-backed preservation. |
| Native modern chart structures | Current modern-looking helpers use broadly compatible classic chart recipes. | Native histogram, Pareto, waterfall, funnel, treemap, sunburst, and box-and-whisker structures plus read/edit support. | Generated and Excel-authored charts round-trip without repair and preserve chart-specific settings. |
| Native in-cell images | Drawings can be anchored precisely to a cell or range. | The spreadsheet's native rich in-cell image value, distinct from a floating/two-cell drawing anchor. | Generated files use the native cell-image model and survive sorting, filtering, row sizing, and copy/paste. |
| Drawing and media fidelity | Strong image placement and manipulation exist. | Shared ICC/CMYK/EXIF handling, additional image metadata compatibility, richer shapes/connectors/grouping, and consistent render/export behavior. | Image handling lives in `OfficeIMO.Drawing`; workbook, document, PDF, and renderer consumers share the same tested codec/color behavior. |
| Memory-bounded large workbook editing | Fast generation and typed streaming are strong. | Bounded-memory load/edit/save for large existing formatted sheets, not only generated tabular output. | A repeatable 500,000-row by 180-column formatted workload records peak memory, elapsed time, output size, and preservation results with configurable budgets. |
| Phonetic and locale metadata | Some legacy/binary projection and preservation behavior exists. | First-class phonetic runs/properties, locale-aware text metadata, and high-level author/read/edit APIs. | East Asian language fixtures round-trip display and phonetic settings across supported formats. |

## External Backlog Signals And How To Get Ahead

The following public demand signals are useful because they expose real failure modes and long-lived workflow needs. They are not a request to copy another API.

| Signal | OfficeIMO position | Recommended response |
| --- | --- | --- |
| [Charts remain a long-running demand area](https://github.com/ClosedXML/ClosedXML/issues/50) | Already ahead. | Protect broad chart authoring with public examples and Excel-open validation; invest next in native modern charts and mutation. |
| [`DateOnly` support is still requested](https://github.com/ClosedXML/ClosedXML/issues/2227) | Already ahead. | Keep `DateOnly`/`TimeOnly` coverage across direct, object, table, binary, and typed-read paths. |
| [Grouped or externally sourced pivots can fail to load](https://github.com/ClosedXML/ClosedXML/issues/2130) | Strong grouping support; external/query-backed mutation remains deeper work. | Add hostile real-file fixtures and explicit preserve/edit/reject findings for each pivot source type. |
| [Large formatted sheets can exhaust memory](https://github.com/ClosedXML/ClosedXML/issues/2734) and [long-lived processes can retain memory](https://github.com/ClosedXML/ClosedXML/issues/1636) | Fast generation is a lead; large existing-workbook editing is not the same workload. | Establish memory budgets, disposable resource ownership, and a streamed edit path with telemetry. |
| [Structural deletion can corrupt array-formula references](https://github.com/ClosedXML/ClosedXML/issues/2847) | This is the most important uncovered foundation. | Build one transactional mutation engine with an impact plan and post-edit validation. |
| [External references can crash workbook loading](https://github.com/ClosedXML/ClosedXML/issues/2820) | Preservation and inspection are strong. | Make “never crash without a diagnostic” a corpus-backed compatibility contract. |
| [Native images inside cells are requested](https://github.com/ClosedXML/ClosedXML/issues/2327) | Range-anchored drawings are strong; native cell images are distinct. | Add native in-cell images after the structural engine can track them through sort and mutation. |
| [Color filtering is being added](https://github.com/ClosedXML/ClosedXML/pull/2817) | Missing from the first-class AutoFilter surface. | Implement font, fill, and icon filters together, including reapply and visible-row enumeration. |
| [Ignored-error regions need a public API](https://github.com/ClosedXML/ClosedXML/pull/1987) | Some lower-level/legacy handling exists. | Promote ignored errors into a format-neutral high-level model and support inspect/add/update/remove. |
| [Data bars need complete gradient metadata](https://github.com/ClosedXML/ClosedXML/pull/2296) | Broad conditional formatting exists. | Complete the entire data-bar object model rather than adding one flag in isolation. |
| [Formula-bearing print areas need exact round-trip behavior](https://github.com/ClosedXML/ClosedXML/pull/2420) | Print areas and titles exist. | Route print definitions through the shared reference syntax tree and preserve formulas verbatim when not rewritten. |
| [Wrapped-text row height remains difficult](https://github.com/ClosedXML/ClosedXML/issues/934) | AutoFit and rendering infrastructure provide a better starting point than most libraries. | Use one shared text-measurement contract for AutoFit and rendering, with explicit font fallback diagnostics. |
| [JPEG metadata can trigger compatibility fixes](https://github.com/ClosedXML/ClosedXML/pull/2818) | Image workflows are broad. | Solve codec and color-profile behavior once in `OfficeIMO.Drawing`, then consume it from Excel and other packages. |
| [Comments from other producers can lose content](https://github.com/ClosedXML/ClosedXML/issues/1920) | Legacy and threaded comment workflows are already broad. | Add cross-producer fixtures and retain unknown comment extensions during edits. |

## Closure Plan

### Milestone 0 — Freeze The Capability Contract

- [ ] Replace ad hoc feature lists with a machine-readable capability manifest: `Author`, `Read`, `Edit`, `Preserve`, `Inspect`, or `Reject`.
- [ ] Record the public API snapshot and link each gap to existing source owners and tests.
- [ ] Build the cross-producer workbook corpus and a package-diff harness that ignores known volatile metadata.
- [ ] Add repeatable correctness, file-size, elapsed-time, and peak-memory baselines for representative small, medium, and large workbooks.

Exit criterion: every roadmap item has an owner, fixture, expected package behavior, and measurable acceptance test.

### Milestone 1 — Build The Structural Mutation Core

- [ ] Create a reusable reference syntax tree and rewriter; migrate formula dependencies, templates, charts, names, and print definitions to it.
- [ ] Add an `ExcelMutationPlan`-style dry run that lists impacted formulas, names, tables, drawings, pivots, and preservation-sensitive parts before making changes.
- [ ] Implement transactional row, column, and cell insertion/deletion with deterministic collision policies. The first row insertion/deletion slice and workbook-wide row-reference remapping are implemented.
- [ ] Implement copy, move, and transpose on the same core.
- [ ] Validate the resulting package, recalculate dependency metadata, and expose actionable post-edit diagnostics.

Exit criterion: one engine owns structural address changes, specialized template code delegates to it, and the dependency-matrix suite passes across all supported formats where the operation is valid.

### Milestone 2 — Complete Everyday Worksheet Editing

- [ ] Add full AutoFilter criteria, reapply, state readback, and visible-row enumeration.
- [ ] Add table resize, replace, schema mutation, typed append, and structured-reference rewrites.
- [ ] Add find-all, formula-aware search, predicates, and range algebra.
- [ ] Add format-neutral allowed edit ranges and ignored-error-region management.
- [ ] Add typed themes, named styles, reusable style snapshots, and view/print state gaps.
- [ ] Complete sparkline and conditional-format read/edit/remove lifecycles.

Exit criterion: common cleanup, import-normalize-export, table refresh, and protected-template jobs can be implemented without raw Open XML.

### Milestone 3 — Unify Formula Semantics

- [ ] Finish A1/R1C1 conversion and route all reference-bearing metadata through the shared syntax tree.
- [ ] Define and implement authored, cached, evaluated, dirty, deferred, and unsupported formula states.
- [ ] Add dynamic-array/spill and data-table metadata before expanding evaluator breadth.
- [ ] Add high-value missing function clusters based on usage and compatibility fixtures.
- [ ] Keep custom functions as an extension point; do not imply that a bounded server-side evaluator is Excel.

Exit criterion: callers can predict whether a result is calculated now, cached, recalculated by Excel, or rejected, and structural edits preserve every supported formula form.

### Milestone 4 — Finish Dashboard Interoperability

- [ ] Implement native slicer and timeline caches, bindings, selection state, read/edit/remove, and shared-cache behavior.
- [ ] Add pivot calculated items/members, shared-cache lifecycle, region styles, and external/query-backed source policies.
- [ ] Complete sparkline group editing and native chart read/mutate workflows.
- [ ] Add native modern chart structures in small, independently validated slices.

Exit criterion: OfficeIMO can load, modify, generate, and reopen representative interactive dashboards without Excel repair prompts or silent loss of interaction metadata.

### Milestone 5 — Extend The Lead

- [ ] Add native in-cell images and track them through sorting, filtering, resizing, copying, and structural edits.
- [ ] Add a memory-bounded edit path for large existing workbooks with configurable resource budgets.
- [ ] Share image codec/color management through `OfficeIMO.Drawing`.
- [ ] Turn preflight plus mutation planning into a CI policy: fail, preserve, warn, or explicitly accept risk by feature.
- [ ] Publish task-oriented recipes for migration, cleanup, dashboard refresh, safe unknown-workbook editing, and large-file processing.

Exit criterion: the differentiators are documented as supported workflows with corpus, performance, and Excel-open evidence—not only as API surface.

## Formula Evaluator Delta

The current evaluator already contains useful modern and reporting functions that are absent from the reviewed baseline, including `XLOOKUP`, `XMATCH`, `TEXTBEFORE`, `TEXTAFTER`, `AVERAGEIFS`, `MINIFS`, `MAXIFS`, `FORECAST.LINEAR`, `WORKDAY.INTL`, and several statistical functions.

The source comparison also found 77 built-in names present in the reviewed evaluator but not in OfficeIMO's built-in list. This is a prioritization input, not a requirement to implement every alias:

- Trigonometric and hyperbolic: `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `SEC`, `SECH`, `SIN`, `SINH`, `TAN`, `TANH`.
- Math and combinatorics: `BASE`, `DECIMAL`, `COMBIN`, `COMBINA`, `EVEN`, `ODD`, `FACT`, `FACTDOUBLE`, `GCD`, `LCM`, `LOG`, `MULTINOMIAL`, `QUOTIENT`, `ROMAN`, `SERIESSUM`, `SQRTPI`.
- Matrix and array: `MDETERM`, `MINVERSE`, `MMULT`, `TRANSPOSE`.
- Text, locale, and linking: `ARABIC`, `ASC`, `CHAR`, `CLEAN`, `CODE`, `DOLLAR`, `FIXED`, `HYPERLINK`, `NUMBERVALUE`, `REPLACE`, `T`.
- Information, logical, and volatile: `ERROR.TYPE`, `FALSE`, `ISEVEN`, `ISLOGICAL`, `ISNONTEXT`, `ISODD`, `ISREF`, `N`, `NA`, `RAND`, `RANDBETWEEN`, `TRUE`, `TYPE`.
- Financial, statistical, and compatibility aliases: `BINOM.DIST`, `BINOMDIST`, `FISHER`, `IPMT`, `STDEV`, `STDEVA`, `STDEVP`, `STDEVPA`, `VAR`, `VARA`, `VARP`, `VARPA`.

Recommended order:

1. Reference and matrix functions needed by structural edits and reporting.
2. Information and text functions used for workbook validation and cleanup.
3. Math/trigonometric families where one implementation can cover several functions consistently.
4. Financial/statistical functions supported by real OfficeIMO use cases.
5. Compatibility aliases as thin mappings after their canonical implementations exist.

## Ownership Rules

- Structural mutation, formula/reference syntax, tables, pivots, filters, protection metadata, and package preservation belong in `OfficeIMO.Excel`, not in examples or one-off template helpers.
- Image decoding, color profiles, orientation, and reusable image metadata belong in `OfficeIMO.Drawing`; Excel should own only spreadsheet placement and relationships.
- Rendering-specific layout belongs in the existing rendering/export packages and should consume the same measurement and style models as AutoFit.
- `.xls` and `.xlsb` surfaces should stay thin over format-neutral capability contracts. A feature that cannot be represented safely in a format should be rejected or reported explicitly, not approximated silently.
- Public methods should describe user jobs. Raw Open XML remains an escape hatch, not the primary implementation path.

## Scope Discipline

Parity should be measured by supported user workflows, not class count, overload count, or identical method names. The plan intentionally does not prioritize:

- Recreating another library's object hierarchy.
- Implementing every Excel calculation function before structural integrity.
- Adding format-specific public APIs when one format-neutral contract is possible.
- Hiding preservation risks or unsupported operations behind a successful `Save()`.
- Replacing the current fast generation paths with a heavyweight mutable model.
- Adding downstream workarounds when the reusable owner should be improved.

## Research Sources

Current public documentation and release material:

- [Workbook and worksheet API](https://docs.closedxml.io/en/latest/api/workbook.html)
- [Formula capabilities](https://docs.closedxml.io/en/latest/features/formulas.html)
- [Calculation and dirty tracking](https://docs.closedxml.io/en/latest/concepts/formula-calculation.html)
- [Function evaluation and dynamic arrays](https://docs.closedxml.io/en/latest/features/functions.html)
- [AutoFilter behavior](https://docs.closedxml.io/en/latest/features/autofilter.html)
- [Sorting behavior](https://docs.closedxml.io/en/latest/features/sort.html)
- [Tables](https://docs.closedxml.io/en/latest/features/tables.html)
- [Pivot tables](https://docs.closedxml.io/en/latest/features/pivot-tables.html)
- [Worksheet protection](https://docs.closedxml.io/en/latest/features/protect.html)
- [Themes](https://docs.closedxml.io/en/latest/features/themes.html)
- [Latest stable release at the review date](https://github.com/ClosedXML/ClosedXML/releases/tag/0.105.1)

OfficeIMO source-of-truth material:

- `OfficeIMO.Excel/README.md`
- `OfficeIMO.Excel/COMPATIBILITY.md`
- `Docs/officeimo.excel.roadmap.md`
- `Docs/officeimo.excel.large-workbook-guidance.md`
- `Docs/officeimo.excel.legacy-xls-roadmap.md`
- `OfficeIMO.Excel.Tests/Excel.Formula*.cs`
- `OfficeIMO.Excel.Tests/Excel.Pivot*.cs`
- `OfficeIMO.Excel.Tests/Excel.Template.cs`
- `OfficeIMO.Excel.Tests/Excel.CompatibilityCorpus*.cs`

This document should be refreshed when structural mutation lands, when a new compatibility corpus materially changes the preserve/edit boundary, or when the reference baseline changes its public calculation, dashboard, or worksheet-mutation contracts.
