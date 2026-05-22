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
- Done initial slice: add text and lookup helpers for `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, and exact-match lookup results that return text.
- Done initial slice: add clearer unsupported-formula diagnostics for unsupported functions, unsupported argument shapes, semicolon-separated formulas, text concatenation, and array constants.

### 2. Charts, Pivots, And Dashboards

Continue turning existing workbook features into polished report-building APIs.

- Done initial slice: add chart presets for KPI scorecards, contribution charts, and waterfall-style variance bridges.
- Expand chart type coverage in practical chunks: waterfall, funnel, histogram, treemap, sunburst, box-and-whisker, stock, radar, and surface.
- Add pivot grouping for dates and numbers.
- Done initial slice: add pivot show-values-as options for data fields, including a fluent percent-of-total helper.
- Add pivot value filters, label filters, top/bottom filters, and calculated fields.
- Add table and pivot slicers once the metadata model is solid.

### 3. Preservation And Feature Inspection

Make it easy to understand what a workbook contains and what OfficeIMO will safely edit or preserve.

- Done initial slice: expand `InspectFeatures()` findings with detail entries for preservation-sensitive package features, including workbook links/external relationships, query/connectors, slicers, timelines, VBA projects, embedded packages, custom XML, signatures, form controls, and OLE markers.
- Done initial slice: add round-trip preservation proof for external hyperlink relationships and custom XML package metadata.
- Done initial slice: add broader round-trip preservation proof for macro project parts and embedded package parts that OfficeIMO does not fully author yet.
- Add broader round-trip preservation tests for additional package features OfficeIMO does not fully author yet.
- Add a workbook corpus covering Excel-authored, LibreOffice-authored, Google Sheets-authored, and generated files.
- Done initial slice: add targeted feature-report guards and examples showing how to fail fast when a workbook contains features a workflow does not want to touch.

### 4. Template Workflows

Turn workbook templates into a first-class report-generation path.

- Done initial slice: add single-row template repetition that inserts rows and binds each supplied row model.
- Add repeating sheets.
- Done initial slice: add template missing-data policies so optional markers can be preserved, cleared, or rejected.
- Done initial slice: add optional row sections that can be kept and bound or removed while shifting following rows.
- Done initial slice: add image binding for whole-cell template markers from byte arrays, streams, and URLs.
- Done initial slice: add stronger formatter hooks for currency, percentages, dates, durations, and custom user formats.
- Preserve surrounding formulas, named ranges, tables, styles, charts, and page setup during binding.

### 5. Comments, Notes, And Review Metadata

Improve collaboration metadata without making it heavy.

- Add rich-text comment authoring.
- Add threaded comment inspection and preservation checks.
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
- Release readiness checklist: `Docs/officeimo.excel.release-checklist.md`
