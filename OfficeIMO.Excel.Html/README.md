# OfficeIMO.Excel.Html

First-party HTML adapter for OfficeIMO.Excel. It exports semantic worksheet tables or visual review HTML using the shared OfficeIMO.Html profile contracts and the existing Excel SVG image exporter.

## Semantic round trips

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Excel.Html;

using ExcelDocument workbook = ExcelDocument.Load("report.xlsx", readOnly: true);
string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
    HeaderMode = ExcelHtmlHeaderMode.FirstRow
});

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
HtmlToExcelResult result = source.ToExcelDocumentResult();
using ExcelDocument imported = result.RequireValue();
imported.Save("report-roundtrip.xlsx");
```

Semantic output carries a versioned OfficeIMO envelope and preserves worksheet names and visibility, used-range coordinates, typed text/number/boolean/date-time values, formulas, comments, merged ranges, embedded image inventory, supported chart inventory, and inert pivot-definition review metadata. HTML `rowspan` and `colspan` values become native Excel merged ranges. Pivot refresh, drill, caches, slicers, and timelines remain native workbook behavior and are not executed in HTML.

`HeaderMode` makes the first-row assumption explicit. `FirstRow` is the compatibility default and emits a real `thead` with column headers. Use `None` when every row is data.

`ToExcelDocument()` is the convenience API. It throws `HtmlConversionException` when no semantic `section.officeimo-sheet` envelope exists. Use `ToExcelDocumentResult()` to receive the workbook plus structured diagnostics and loss classification. Export callers can use `ToHtmlResult()` for the same evidence shape; pivot simplification, truncation, unavailable chart or image content, and visual-renderer fallbacks are operation-scoped diagnostics rather than HTML-only prose.

Ordinary HTML tables are available explicitly through the shared generic projector:

```csharp
HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
    .ToExcelDocumentResult(new HtmlToExcelOptions {
        Mode = HtmlImportMode.Auto
    });
```

`Semantic` remains the default for strict round-trip compatibility. `Auto` selects a supported semantic envelope when present and otherwise maps ordinary tables to worksheets; `Generic` always uses the ordinary HTML path. `HtmlToExcelOptions.Limits` bounds worksheets, tables, cells, images, chart dimensions, metadata, and geometry before native allocations. `MaxTableCells` remains as a forwarding compatibility property.

On the ordinary HTML path, bounded positioned, floating, flex, and grid regions become editable merged-cell regions plus absolute DrawingML picture anchors. Solid backgrounds become cell fills, and foreground pictures retain supported native opacity. CSS background-image layers are omitted so they cannot cover editable cell text and produce a stable diagnostic. Excel has no editable cell-shadow equivalent, so shadows and unsupported effects are diagnosed while content and geometry remain editable. Set `ImportEditableLayoutRegions = false` to retain semantic flow only.

`SaveAsHtml` and `SaveAsHtmlAsync` write UTF-8 without a byte-order mark to paths or caller-owned streams. For import I/O, use `HtmlConversionDocument.Load(...)` or `LoadAsync(...)`, then call `ToExcelDocument()` or `ToExcelDocumentResult()` on the prepared document. Stream overloads leave caller-owned streams open.

## Typed values in ordinary report tables

Set `ImportTypedCellValues = true` with `Mode = HtmlImportMode.Generic` (or `Auto` when no semantic envelope is present) to import explicitly declared scalar values:

```html
<td data-officeimo-value-kind="number" data-officeimo-value="12.5"><strong>12.50</strong></td>
<td data-officeimo-value-kind="boolean" data-officeimo-value="true">Approved</td>
<td data-officeimo-value-kind="date-time" data-officeimo-value="2026-09-08T00:00:00">8 Sep 2026</td>
<td>00127</td>
```

The first three cells become a number, boolean, and date in Excel. The reference remains text, including its leading zeros. Supported kinds are `text`, `number`, `boolean`, and `date-time`; use invariant numeric values and ISO date/time values. Invalid or oversized metadata falls back to bounded visible text with a diagnostic. The option defaults to `false`, so existing generic imports keep their text behavior.

Scalar values survive bold, italic, and color formatting. Excel applies the first visible run's style to the whole scalar cell; it cannot store different rich-text styles inside a numeric cell. Excel displays native values using its number formats, so labels such as `Approved` become `TRUE` and decimal padding may differ. Generic imports do not execute formula metadata or honor semantic cell-coordinate overrides. Full workbook restoration remains a separate semantic-envelope operation with its existing trust controls.

The [multi-format report example](../OfficeIMO.Examples/Converters/Html/HtmlMultiFormatReport.cs) exports one HTML source to HTML, PDF, Word, and Excel.

## Visual review

Use `ExcelHtmlSaveOptions.CreateVisualReviewProfile()` or set `ExportProfile = ExcelHtmlExportProfile.VisualReview` to emit review HTML through OfficeIMO's dependency-free SVG renderer. `SharedProfile` exposes the corresponding generic engine lane. `DocumentOutput` controls full-document versus fragment output, title, language, theme, default styles, and newlines. Visual-review HTML is presentation evidence; use semantic tables when the HTML must be imported back into Excel.

## Targets

`netstandard2.0`, `net8.0`, and `net10.0`; `net472` is included when building on Windows.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Excel`, `OfficeIMO.Html`, and `OfficeIMO.Core` own the workbook, HTML source, mapping, visual review, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
