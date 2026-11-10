# OfficeIMO.Data.Conversions.Excel

Worksheets in OfficeIMO structured-data conversions: worksheets and workbooks to JSON, JSON Lines and XML, and the rows of JSON, YAML and XML data to workbooks through the streaming writer.

## Install

```powershell
dotnet add package OfficeIMO.Data.Conversions.Excel
```

## Examples

```csharp
using System.Text.Json.Nodes;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Excel;

using ExcelDocument workbook = ExcelDocument.Load("Sales.xlsx");
JsonNode? q3 = workbook["Q3"].ToJson();                                  // an array of row objects
workbook.SaveAsJson("sales.json");                                        // one key per sheet
workbook["Q3"].SaveAsJson("q3.json", new DataConversionOptions { Range = "A1:F200" });

JsonNode.Parse(File.ReadAllText("export.json"))!.SaveAsExcel("export.xlsx");
```

A worksheet is read with first-row headers; path headers such as `address.city` rebuild nested objects. Formulas contribute their cached values. Writing a workbook reports what a cell cannot hold: numbers beyond 15 significant digits become text, date-times lose their offset, and fractional seconds finer than a millisecond are rounded.

To insert JSON rows into an existing sheet, open them with `StructuredDataReader` and call `ExcelSheet.InsertDataReader`.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 7 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Data.Conversions.Excel` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
