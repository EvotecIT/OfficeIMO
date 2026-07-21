# OfficeIMO.Excel.Html

First-party HTML adapter for OfficeIMO.Excel. It exports semantic worksheet tables or visual review HTML using the shared OfficeIMO.Html profile contracts and the existing Excel SVG image exporter.

## Semantic round trips

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;

using ExcelDocument workbook = ExcelDocument.Load("report.xlsx", readOnly: true);
string html = workbook.ToHtml(new ExcelHtmlSaveOptions {
    HeaderMode = ExcelHtmlHeaderMode.FirstRow
});

HtmlToExcelResult result = html.ToExcelDocumentResult();
using ExcelDocument imported = result.GetArtifactOrThrow();
imported.Save("report-roundtrip.xlsx");
```

Semantic output carries a versioned OfficeIMO envelope and preserves worksheet names and visibility, used-range coordinates, typed text/number/boolean/date-time values, formulas, comments, merged ranges, embedded image inventory, and supported chart inventory. HTML `rowspan` and `colspan` values become native Excel merged ranges.

`HeaderMode` makes the first-row assumption explicit. `FirstRow` is the compatibility default and emits a real `thead` with column headers. Use `None` when every row is data.

`ToExcelDocument()` is the convenience API. It throws `HtmlConversionException` when no semantic `section.officeimo-sheet` envelope exists. Use `ToExcelDocumentResult()` to receive the workbook plus structured diagnostics and loss classification. Export callers can use `ToHtmlResult()` for the same evidence shape.

Ordinary HTML tables are available explicitly through the shared generic projector:

```csharp
HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
    .ToExcelDocumentResult(new HtmlToExcelOptions {
        Mode = HtmlImportMode.Auto
    });
```

`Semantic` remains the default for strict round-trip compatibility. `Auto` selects a supported semantic envelope when present and otherwise maps ordinary tables to worksheets; `Generic` always uses the ordinary HTML path. `HtmlToExcelOptions.Limits` bounds worksheets, tables, cells, images, chart dimensions, metadata, and geometry before native allocations. `MaxTableCells` remains as a forwarding compatibility property.

Path, stream, and async save/import methods use UTF-8 without a byte-order mark. Stream overloads leave caller-owned streams open.

## Visual review

Set `Profile = OfficeHtmlConversionProfile.ExcelVisualReview` to emit review HTML through OfficeIMO's dependency-free SVG renderer. Visual-review HTML is presentation evidence; use semantic tables when the HTML must be imported back into Excel.

## Targets

`netstandard2.0`, `net8.0`, and `net10.0`; `net472` is included when building on Windows.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Excel`, `OfficeIMO.Html`, and `OfficeIMO.Drawing` own the workbook, HTML source, mapping, visual review, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
