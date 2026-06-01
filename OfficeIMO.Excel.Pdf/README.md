# OfficeIMO.Excel.Pdf

First-party Excel workbook to PDF export using `OfficeIMO.Pdf`.

The initial exporter maps everyday worksheet used ranges into reusable PDF tables. It is intentionally thin: workbook reading stays in `OfficeIMO.Excel`, layout and PDF writing stay in `OfficeIMO.Pdf`, and this package only translates worksheet data into PDF document primitives.

Current scope:

- All workbook sheets, or a selected sheet list.
- Worksheet used range detection through the existing Excel reader bridge.
- Worksheet print areas when configured.
- Worksheet orientation and margins when configured, with explicit PDF options still available for overrides.
- Hidden workbook worksheets omitted from default all-sheets exports, while explicitly selected hidden sheets can still be exported.
- Hidden worksheet rows and columns omitted by default.
- Repeated print-title rows mapped to PDF table header rows, including repeat-on-page behavior for long tables.
- Manual worksheet row and column page breaks mapped to explicit PDF page breaks between exported table chunks, preserving repeated header/title rows, with `ExcelPdfSaveOptions.UseWorksheetPageBreaks` available to disable that behavior.
- Simple worksheet header/footer text zones, first-page and even-page text variants, and supported header/footer images, including page number, total page count, sheet-name, date, time, workbook file-name, and workbook path tokens. Simple line-level header/footer font family/style uses the shared `OfficeIMO.Pdf` standard-font mapper for common office aliases such as Aptos, Calibri, Arial, Times New Roman, and Consolas, with font size and RGB text color mapped when the styled text can be represented by one first-party PDF header/footer line style.
- Sheet names as PDF headings.
- Cell display values rendered as PDF table cells, with common number formats, basic cell font emphasis, font color, fill color, two-color conditional color-scale fills, conditional data bars as proportional in-cell PDF table overlays, conditional icon-set indicators as first-party table-cell vector icons, horizontal/vertical alignment, simple cell borders including dashed, dotted, dash-dot, double, and diagonal strokes, external cell hyperlinks, internal workbook links mapped only when their target cell is exported as an exact PDF named destination, explicit worksheet column widths, explicit worksheet row heights, manual worksheet print scale, fit-to-width table sizing against effective page margins, and worksheet merged cells mapped through first-party rich table cells, per-cell table fills, per-cell table data bars, per-cell table icons, per-cell table alignment/border/padding overrides, relative table column widths, table max-width caps, row minimum heights, visible-row/column filtering, cell-owned URI and named-destination annotations, and table row/column spans.
- Supported worksheet drawing images anchored into exported PDF table cells when the anchor cell is exported and otherwise emitted as first-party PDF flow images in worksheet anchor order.
- Supported worksheet column, bar, line, area, scatter, radar, pie, and doughnut chart families exported as first-party vector drawing snapshots when the chart data can be read from the workbook.
- `ExcelPdfSaveOptions.Warnings` reports workbook features that are skipped or simplified during export, including mixed or rich per-run header/footer formatting, unsupported header/footer fields, unsupported or unreadable worksheet/header/footer images, unsupported or unreadable chart snapshots, and row truncation when `MaxRowsPerSheet` is used.
- First row styling through the reusable PDF table header model.
- Page size and margin options through first-party `OfficeIMO.Pdf` geometry types.
- Poppler visual baseline coverage for a daily two-sheet workbook with worksheet header/footer text and images, orientation/margins, merged title cells, fills/borders, number formats, explicit row/column sizing, hidden row/column filtering, worksheet images anchored into exported table cells, chart snapshots, and internal/external links.

Planned scope includes richer worksheet header/footer formatting beyond the current line-level style mapping, richer fit-to-height and automatic multi-page pagination/scaling, richer merged-cell edge cases, richer worksheet image placement fidelity beyond exported table-cell anchors, richer chart fidelity beyond the initial column/bar/line/area/scatter/radar/pie/doughnut snapshots, richer cell style fidelity such as additional conditional formats and locale-specific formats, and broader diagnostics for workbook features that still cannot be mapped faithfully.
