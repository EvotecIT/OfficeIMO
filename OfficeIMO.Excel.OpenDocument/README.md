# OfficeIMO.Excel.OpenDocument

`OfficeIMO.Excel.OpenDocument` explicitly converts between `OfficeIMO.Excel` workbooks and native `OfficeIMO.OpenDocument` spreadsheets. It does not invoke Excel or LibreOffice; the adapter depends on the two OfficeIMO object-model packages it connects.

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;

using ExcelDocument workbook = ExcelDocument.Load("report.xlsx");
OdfConversionResult<OdsDocument> conversion = workbook.ToOpenDocumentResult();
conversion.Value.Save("report.ods");

foreach (var mapping in conversion.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter maps worksheets, typed cell values, formulas, hyperlinks, comments and ODF annotations, merges, row/column layout, named ranges, number-format categories, and a basic style subset. List, whole-number, decimal, text-length, constant-date (1904–9956), and whole-second time validations round-trip with input/error messages and error severity. Excel date and time bounds may use numeric serials or simple `DATE(year,month,day)` and `TIME(hour,minute,second)` formulas with literal, in-range arguments. Custom validation formulas round-trip when one contiguous range uses a local cell comparison against a number, text literal, or another local cell, or `ISBLANK`, `ISNUMBER`, or `ISTEXT` on one local cell. These predicates can be combined with `AND`, `OR`, and unary `NOT`. The ODF rule records the range's first cell as its formula base. Unsupported ODS validation expressions remain editable as ODF conditions and are reported during conversion. Unsupported Excel rules, including dynamic date formulas, subsecond times, other custom-formula functions, multiple target ranges, and dates outside the guaranteed OpenFormula range, are omitted with an explicit loss report.

The ODS-to-Excel route maps up to 16 ordered numeric `cell-content()` comparisons or `cell-content-is-between` / `cell-content-is-not-between` conditions with literal numeric bounds on one cell style to Excel cell-value rules. Rule priority and stop-if-true preserve the first matching style. Each applied common table-cell style needs a direct background fill or supported font property. The Excel differential style receives direct text color, explicit bold/italic/underline/strikethrough on or off, one font family, and absolute font sizes from 1 to 409 points. Font-only rules do not need a fill. Range bounds must be in ascending order. Double or patterned line-through styles are not mapped.

The route covers up to 4,096 styled cells per source style on each sheet when basic styles are included. The report marks mapped rules as approximate because other ODF applied-style properties, font-family fallback lists, relative font sizes, and evaluation details are not transferred. A style with more than 16 maps or any unsupported map remains wholly unsupported. Formula bounds, other conditions, maps without a supported direct style property, and rules above the cell limit remain unsupported.

The Excel-to-ODS route maps up to 16 numeric cell-value rules with solid differential fills on one shared contiguous range per sheet. It preserves priority and first-match behavior when every rule except the last uses stop-if-true, and applies the maps to at most 4,096 cells that have no basic cell style. The report marks these mappings as approximate because Excel differential formatting and ODF applied styles have different evaluation models. Styled target cells, overlapping merges, nonnumeric bounds, mixed differential styles, multiple ranges, other conditions, and rules beyond the limits remain explicit conversion loss; unsupported chains are not partially applied.

ODF permits one annotation per spreadsheet cell. An Excel threaded discussion is therefore flattened into one readable annotation transcript per cell, retaining available author, timestamp, identity, parent, resolved-state, and body metadata while reporting the thread mapping as an approximation.

Formula and address conversion uses typed Excel A1/OpenFormula syntax. Quoted worksheet names, absolute references, ranges, arrays, unions, intersections, strings, and separator changes are handled structurally; unsupported structured or external references fail closed and retain cached ODS values where available.

`ExcelOpenDocumentConversionOptions` bounds rows, columns, converted cells, and merge or validation-range materialization in both directions. Content omitted by those limits or disabled style options is returned as a `Skipped` mapping rather than silently disappearing. Set `LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss` when a workflow must reject any approximation, skipped feature, or unsupported mapping.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Excel` and `OfficeIMO.OpenDocument`; the adapter owns bounded feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Excel.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
