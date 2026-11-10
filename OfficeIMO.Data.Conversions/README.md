# OfficeIMO.Data.Conversions

Streaming conversions between JSON, JSON Lines, YAML, XML and CSV. Data flows from each format's own reader to the target's writer as events, so a large file converts with bounded memory. Numbers keep their exact text (`1.10`, 23-digit integers), and everything a target cannot carry is reported in `OfficeConversionReport`.

Worksheets are in [OfficeIMO.Data.Conversions.Excel](../OfficeIMO.Data.Conversions.Excel/README.md). The `OfficeIMO` package includes both, so `OfficeConverter.Convert("a.yaml", "a.json")` works with one reference.

## Install

```powershell
dotnet add package OfficeIMO.Data.Conversions
```

## Convert documents

```csharp
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Yaml;

YamlDocument values = YamlDocument.Load("values.yaml");
JsonNode? json = values.ToJson();
OfficeConversionReport report = values.SaveAsJson("values.json");

XDocument.Load("orders.xml").SaveAsJson("orders.json");
CsvDocument.Load("people.csv").SaveAsJson("people.jsonl");     // JSON Lines from the extension

JsonNode response = JsonNode.Parse(File.ReadAllText("response.json"))!;
response.SaveAsCsv("items.csv", new DataConversionOptions { Select = "data.items" });
```

Every `To{Target}` method has an overload with `out OfficeConversionReport report`. Set `DataConversionOptions.FailOnLoss` to throw `OfficeConversionException` instead of writing output when anything is lost.

## Read rows as a data reader

`StructuredDataReader` exposes the rows of JSON, JSON Lines, YAML or XML as a `DbDataReader`, for `SqlBulkCopy`, `OfficeIMO.Data.Arrow`, `DataTable.Load` or `ExcelSheet.InsertDataReader`.

```csharp
using System.Data;
using OfficeIMO.Data.Conversions;

using StructuredDataReader rows = StructuredDataReader.Open("orders.json", new DataConversionOptions { Select = "data.items" });
var table = new DataTable();
table.Load(rows);   // columns such as id, customer.name and lines[0].sku, typed from the data
```

The input is read twice: once to find the columns and their types, then to stream the rows.

## Mapping

- **Rows** come from an array: the root, each document of a YAML stream or JSON Lines file, or the path in `Select`. A root object whose members are all arrays is several tables (`NextResult`, or one sheet each in Excel).
- **Columns** are paths: `address.city`, `tags[0]`. `ArrayColumns = Join` puts scalar arrays in one cell instead. Path headers rebuild nested data on the way back.
- **CSV types** are inferred per column from canonical spellings only, so `007`, `+1`, `1,000` and `yes` stay text.
- **XML** follows one convention: the root element is the single top-level key, attributes are `@name`, text beside attributes or elements is `#text`, repeated elements are arrays and `xsi:nil="true"` is null. DTDs are prohibited.
- **YAML** follows the document's schema: YAML 1.1 timestamps and merge keys apply only under `%YAML 1.1` or a forced schema. Aliases are expanded within `MaxAliasExpansion`.

The [structured-data conversions design](../Docs/officeimo.data-conversions-design.md) lists every mapping rule and loss code.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 16 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Data.Conversions` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
