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

- Add a public calculation facade such as `doc.Calculate()` and save options for calculate-before-save flows.
- Add dependency ordering so formulas can depend on other formulas, not only literal cells and ranges.
- Add cross-sheet references, named ranges, and simple table references.
- Add text and lookup helpers such as `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, and text-returning lookup results.
- Add clear diagnostics for formulas that are preserved but not calculated by OfficeIMO.

### 2. Charts, Pivots, And Dashboards

Continue turning existing workbook features into polished report-building APIs.

- Add chart presets for variance waterfall, KPI scorecard, and contribution charts.
- Expand chart type coverage in practical chunks: waterfall, funnel, histogram, treemap, sunburst, box-and-whisker, stock, radar, and surface.
- Add pivot grouping for dates and numbers.
- Add pivot value filters, label filters, top/bottom filters, calculated fields, and show-values-as options.
- Add table and pivot slicers once the metadata model is solid.

### 3. Preservation And Feature Inspection

Make it easy to understand what a workbook contains and what OfficeIMO will safely edit or preserve.

- Expand `InspectFeatures()` with richer detail for workbook links, query tables, slicers, timelines, VBA projects, embedded objects, custom XML, signatures, and form controls.
- Add round-trip preservation tests for features OfficeIMO does not fully author yet.
- Add a workbook corpus covering Excel-authored, LibreOffice-authored, Google Sheets-authored, and generated files.
- Add small examples showing how to fail fast when a workbook contains features a workflow does not want to touch.

### 4. Template Workflows

Turn workbook templates into a first-class report-generation path.

- Add repeating rows and repeating sheets.
- Add optional sections and missing-data policies.
- Add image binding from byte arrays, streams, and URLs.
- Add stronger formatter hooks for currency, percentages, dates, durations, and custom user formats.
- Preserve surrounding formulas, named ranges, tables, styles, charts, and page setup during binding.

### 5. Comments, Notes, And Review Metadata

Improve collaboration metadata without making it heavy.

- Add rich-text comment authoring.
- Add threaded comment inspection and preservation checks.
- Add APIs for finding, updating, and removing comments by author, range, and text.
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
