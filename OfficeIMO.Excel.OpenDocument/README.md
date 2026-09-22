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

The adapter maps worksheets, typed cell values, formulas, hyperlinks, comments and ODF annotations, merges, row/column layout, named ranges, number-format categories, and a basic style subset. List, whole-number, decimal, text-length, constant-date, and whole-second time validations round-trip with input/error messages and error severity. Date and time rules with dynamic formulas or unsupported precision, custom-formula rules, and implementation-specific validation expressions are preserved by ODF editing but reported as unsupported when an exact Excel mapping is unavailable.

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
