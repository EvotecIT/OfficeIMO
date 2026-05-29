# OfficeIMO.Excel Roadmap

Date: 2026-05-21

This roadmap tracks where `OfficeIMO.Excel` should grow next while keeping the API pleasant, explicit, and easy to adopt. The theme is simple: make the common workbook/reporting path obvious, keep advanced escape hatches available, and be honest about what OfficeIMO can edit, preserve, calculate, or defer to Excel.

## Direction

`OfficeIMO.Excel` should stay focused on practical workbook automation:

- Fast data export and import for objects, tables, data sets, CSV, JSON, and typed rows.
- Clear report-building primitives for tables, charts, pivots, sparklines, styles, images, validation, and conditional formatting.
- Safe workbook handling that preserves unsupported workbook parts whenever possible.
- Formula diagnostics and calculation support for the reporting formulas users most often need in server-side workflows.
- Template workflows that let users design a workbook in Excel and bind data through code.

## Roadmap

### 1. Formula Calculation

Build the lightweight evaluator into a dependable reporting-calculation layer.

- Done initial slice: add `doc.Calculate()` and per-save formula options on `ExcelSaveOptions` for calculate-before-save flows.
- Done initial slice: add same-sheet dependency ordering so supported formulas can depend on other supported formulas, not only literal cells and ranges.
- Done initial slice: add numeric cross-sheet cell/range references for supported formulas.
- Done initial slice: add workbook-global and sheet-local named range references for supported numeric formulas.
- Done initial slice: add simple table structured references for supported numeric formulas.
- Done initial slice: add text and lookup helpers for `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `TEXTBEFORE`, `TEXTAFTER`, `FORMULATEXT`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `SUBSTITUTE`, `FIND`, `SEARCH`, `VALUE`, `EXACT`, `REPT`, and exact-match lookup results that return text.
- Done initial slice: add `XLOOKUP` if-not-found fallback support plus forward/reverse exact search for lightweight report calculations.
- Done initial slice: add `XMATCH` plus bounded next-smaller/next-larger numeric matching for `MATCH`, `XMATCH`, and `XLOOKUP`.
- Done initial slice: add clearer unsupported-formula diagnostics for unsupported functions, unsupported argument shapes, semicolon-separated formulas, text concatenation, and array constants.
- Done initial slice: add direct formula dependency reporting plus dependency issue guards for self-references and formula dependencies without cached results.
- Done initial slice: add workbook-level formula dependency graphs with direct dependencies, formula dependents, dependency issue rollups, lookup helpers, and Markdown export.
- Done initial slice: add `MINIFS` and `MAXIFS` to multi-criteria report aggregation support.
- Done initial slice: add `COUNTBLANK` and bounded `SUBTOTAL` support for common report summary and audit formulas.
- Done initial slice: add bounded financial report formulas for `PMT`, `PV`, `FV`, `NPER`, and `NPV`.
- Done initial slice: add bounded statistical report formulas for `SUMSQ`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `QUARTILE.INC`, `QUARTILE.EXC`, `PERCENTRANK.INC`, `PERCENTRANK.EXC`, `RANK.EQ`, `RANK.AVG`, `CORREL`, `SLOPE`, `INTERCEPT`, `RSQ`, and `FORECAST.LINEAR`.
- Done initial slice: add bounded descriptive-statistics report formulas for `MEDIAN`, `LARGE`, `SMALL`, `MODE.SNGL`, `MODE`, `GEOMEAN`, and `HARMEAN`.
- Done initial slice: add deviation and paired-array statistical report formulas for `AVEDEV`, `DEVSQ`, `SUMXMY2`, `SUMX2MY2`, and `SUMX2PY2`.
- Done initial slice: add covariance report formulas for `COVARIANCE.P`, `COVARIANCE.S`, and `COVAR`.
- Done initial slice: add bounded rounding-to-significance report formulas for `MROUND`, `CEILING.MATH`, and `FLOOR.MATH`.
- Done initial slice: add standard rounding helpers for `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `INT`, `CEILING`, and `FLOOR`, including negative digit positions for report-scale rounding.
- Done initial slice: add value-returning `CHOOSE` selectors for numeric and text report formulas.
- Done initial slice: add `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISERROR`, `ISERR`, `ISNA`, and `ISFORMULA` report guards for conditional formulas.
- Done initial slice: add `IFNA` fallbacks and strengthen explicit error-value handling for `IFERROR`.
- Done initial slice: add `UPPER`, `LOWER`, and `PROPER` text case helpers for report labels.
- Done initial slice: add `ROW`, `COLUMN`, `ROWS`, and `COLUMNS` reference-shape helpers for report diagnostics.
- Done initial slice: add `DATEVALUE` and `TIMEVALUE` helpers for deterministic report date/time label parsing.
- Done initial slice: add A-value aggregate report formulas for `AVERAGEA`, `MINA`, and `MAXA`, including literal text/logical values and mixed cell ranges.
- Done initial slice: add week and 360-day calendar report formulas for `WEEKNUM`, `ISOWEEKNUM`, and `DAYS360`.

### 2. Charts, Pivots, And Dashboards

Continue turning existing workbook features into polished report-building APIs.

- Done initial slice: add chart presets for KPI scorecards, contribution charts, and waterfall-style variance bridges.
- Done initial slice: add radar chart authoring, fluent range/table helper coverage, data update support, and OpenXML validation coverage.
- Done initial slice: add stock chart authoring for high-low-close and open-high-low-close ranges with update/readback validation.
- Done initial slice: add 3-D surface chart authoring, update support, series styling, and OpenXML validation coverage.
- Done initial slice: add wireframe surface, contour surface, and wireframe contour surface chart variants with update/readback support and OpenXML validation coverage.
- Done initial slice: add 3-D pie chart authoring, update/readback support, data labels, series styling, and OpenXML validation coverage.
- Done initial slice: add pie-of-pie and bar-of-pie chart authoring, update/readback support, labels, styling, and OpenXML validation coverage.
- Done initial slice: add 3-D clustered/stacked column and bar chart authoring, update/readback support, labels, styling, and OpenXML validation coverage.
- Done initial slice: add 100% stacked column/bar chart variants, including 3-D 100% stacked column/bar, update/readback support, styling, and OpenXML validation coverage.
- Done initial slice: add stacked and 100% stacked line chart variants, update/readback support, labels, markers, styling, and OpenXML validation coverage.
- Done initial slice: add 3-D line chart authoring, update/readback support, labels, markers, styling, and OpenXML validation coverage.
- Done initial slice: add 3-D standard/stacked area chart authoring, update/readback support, labels, styling, and OpenXML validation coverage.
- Done initial slice: add stacked and 100% stacked area chart variants, including 3-D 100% stacked area, update/readback support, styling, and OpenXML validation coverage.
- Expand chart type coverage in practical chunks: waterfall, funnel, histogram, treemap, sunburst, and box-and-whisker.
- Done initial slice: add pivot date and number grouping metadata with fluent helpers, typed cache shared items, and readback metadata.
- Done initial slice: add pivot show-values-as options for data fields, including a fluent percent-of-total helper.
- Done initial slice: add pivot label and value filters with fluent helpers and readback metadata.
- Done initial slice: add pivot top/bottom count, percent, and sum filters with readback metadata.
- Done initial slice: add formula-backed pivot calculated fields that can be used as data fields.
- Done initial slice: add broader label/value pivot filter helper variants, including negated and not-between filters.
- Done initial slice: add fixed date pivot filter helpers for date comparisons and date ranges.
- Done initial slice: add dynamic date pivot filter helpers for relative periods, months, and quarters.
- Done initial slice: add generated multi-level pivot date hierarchy fields for year/quarter/month-style row grouping, with cache metadata and readback coverage.
- Done initial slice: add explicit pivot grouping item metadata for grouped cache fields, including generated date hierarchy readback.
- Done initial slice: add base/parent field relationship metadata for generated pivot date hierarchy fields.
- Done initial slice: add fluent pivot helpers for item visibility filters and selected page/filter items.
- Done initial slice: add selected page/filter item readback on pivot field inspection.
- Done initial slice: add visible item readback for pivot field item filters.
- Done initial slice: add composable fluent pivot field helpers for sort and subtotal placement.
- Done initial slice: add fluent pivot field helpers for layout, display flags, breaks, and subtotal captions.
- Done initial slice: add fluent pivot field number-format helpers that compose with other field options.
- Done initial slice: add pivot field, data field, and calculated field number-format code readback.
- Done initial slice: add built-in Excel number-format code readback for pivot fields and data fields.
- Done initial slice: add preserve-only feature inspection coverage for slicer and timeline package metadata so dashboard workbooks can be preflighted before mutation.
- Continue pivot grouping work with deeper Excel interoperability checks against real Excel-authored grouped pivot files.
- Add calculated item/member scenarios.
- Add table and pivot slicers once the metadata model is solid.

### 3. Preservation And Feature Inspection

Make it easy to understand what a workbook contains and what OfficeIMO will safely edit or preserve.

- Done initial slice: expand `InspectFeatures()` findings with detail entries for preservation-sensitive package features, including workbook links/external relationships, query/connectors, slicers, timelines, VBA projects, embedded packages, custom XML, signatures, form controls, and OLE markers; slicer/timeline, digital signature origin/app-property metadata, and OLE/form-control package signals now have focused test coverage.
- Done initial slice: add round-trip preservation proof for external hyperlink relationships and custom XML package metadata.
- Done initial slice: add broader round-trip preservation proof for macro project parts and embedded package parts that OfficeIMO does not fully author yet.
- Done initial slice: add round-trip preservation proof for slicer and timeline package metadata that OfficeIMO detects as preserve-only.
- Done initial slice: add round-trip preservation proof for OLE object and form-control markers that OfficeIMO detects as preserve-only.
- Done initial slice: add threaded-comment feature inspection details for sheet, cell, and author plus round-trip preservation proof for threaded comments and workbook person metadata.
- Add broader round-trip preservation tests for additional package features OfficeIMO does not fully author yet.
- Add a workbook corpus covering Excel-authored, LibreOffice-authored, Google Sheets-authored, and generated files.
- Done initial slice: add targeted feature-report guards and examples showing how to fail fast when a workbook contains features a workflow does not want to touch.

### 4. Template Workflows

Turn workbook templates into a first-class report-generation path.

- Done initial slice: add single-row template repetition that inserts rows and binds each supplied row model.
- Done initial slice: add repeating worksheet templates that generate named sheets from per-sheet models and preserve cells, formulas, styles, merges, columns, data validations, conditional formatting, page setup, sheet-local defined names, workbook-scoped defined names that reference the template sheet, print areas/titles, table definition relationships with fresh table ids and workbook-unique table names, external hyperlink relationships with fresh relationship ids, internal hyperlink locations, legacy comments with copied VML note shapes, static image drawing parts with fresh image relationships, chart style/color parts, chart embedded package parts, chart user-shape drawing parts with nested images, diagram drawing parts with nested images, and generated chart/table/cell/validation/rule/name formulas rewritten to the generated sheet name.
- Done initial slice: add template missing-data policies so optional markers can be preserved, cleared, or rejected.
- Done initial slice: add optional row sections that can be kept and bound or removed while shifting following rows.
- Done initial slice: add image binding for whole-cell template markers from byte arrays, streams, and URLs.
- Done initial slice: add stronger formatter hooks for currency, percentages, dates, durations, and custom user formats.
- Done initial slice: preserve named ranges and table references when repeating or removing template rows.
- Preserve relationship-backed drawing structures, workbook-level structures, charts, and additional template binding scenarios.

### 5. Comments, Notes, And Review Metadata

Improve collaboration metadata without making it heavy.

- Done initial slice: add rich-text legacy comment authoring and update support using the existing `ExcelRichTextRun` model.
- Done initial slice: add threaded comment inspection and preservation checks.
- Done initial slice: add APIs for finding, updating, and removing legacy comments by author, range, and text.
- Document which comment features are editable and which are preservation-only.

### 6. Rendering Feasibility

Keep rendering scoped until the implementation path is proven.

- Run a feasibility spike for report-grade PDF, HTML, and image output using existing OfficeIMO primitives where possible.
- Start with sheets produced by OfficeIMO report APIs: tables, simple styles, merged cells, images, headers/footers, page setup, and chart placeholders or generated chart images.
- Add export APIs only after the scoped renderer is reliable enough for report workflows.

## Product Principles

- The simple path should read like the user’s intent.
- Advanced APIs should remain available without forcing every user into the Open XML object model.
- Unsupported features should be preserved where possible and reported clearly.
- Roadmap items should become small, tested slices before they become broad surface area.
- Documentation should be organized around jobs: export objects, read a spreadsheet, build a dashboard, create an invoice, inspect workbook features, and generate a report.

## Source Notes

- Current package documentation: `OfficeIMO.Excel/README.md`
- Current feature status: `OfficeIMO.Excel/COMPATIBILITY.md`
- Large workbook guidance: `Docs/officeimo.excel.large-workbook-guidance.md`
- Unified Word/Excel capability assessment: `Docs/officeimo.word-excel-capability-assessment.md`
- Release readiness checklist: `Docs/officeimo.excel.release-checklist.md`
