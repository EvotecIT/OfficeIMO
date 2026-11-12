---
title: "Data Conversions"
description: "Convert between JSON, JSON Lines, YAML, TOML, XML, CSV, TSV and Excel worksheets in .NET with streaming, exact numbers and a report of everything the target format cannot hold."
order: 58
meta.seo_title: "Convert JSON, YAML, TOML, XML, CSV and Excel in C# with OfficeIMO"
---

`OfficeIMO.Data.Conversions` converts between JSON, JSON Lines, YAML, TOML, XML, CSV and TSV. `OfficeIMO.Data.Conversions.Excel` adds worksheets and workbooks. Data flows from each format's own reader to the target's writer as events, so a large file converts with bounded memory, and numbers keep their exact text: `1.10` stays `1.10`, and a 23-digit integer is not rounded through `double`.

```shell
dotnet add package OfficeIMO.Data.Conversions
dotnet add package OfficeIMO.Data.Conversions.Excel
```

The `OfficeIMO` package includes both.

## Convert files

Every route uses the same verbs: `To{Target}()` returns the target's document model, and `SaveAs{Target}(path)` writes a file or stream. The path extension picks JSON Lines (`.jsonl`, `.ndjson`) and TSV (`.tsv`, `.tab`).

```csharp
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.CSV;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Toml;
using OfficeIMO.Yaml;

YamlDocument.Load("values.yaml").SaveAsJson("values.json");
XDocument.Load("orders.xml").SaveAsJson("orders.json");
CsvDocument.Load("people.csv").SaveAsJson("people.jsonl");
TomlDocument.Load("Cargo.toml").SaveAsYaml("cargo.yaml");   // comments travel with their keys

JsonNode response = JsonNode.Parse(File.ReadAllText("response.json"))!;
response.SaveAsCsv("items.csv", new DataConversionOptions { Select = "data.items" });
```

With the `OfficeIMO` package, `OfficeConverter` runs the same routes by file extension, which suits code that receives paths:

```csharp
using OfficeIMO;

OfficeConverter.Convert("app.yaml", "app.json");
```

## Worksheets

A worksheet reads with first-row headers, and path headers such as `address.city` rebuild nested objects. Formulas contribute their cached values.

```csharp
using OfficeIMO.Data.Conversions;
using OfficeIMO.Excel;

using ExcelDocument workbook = ExcelDocument.Load("Sales.xlsx");
workbook["Q3"].SaveAsJson("q3.json", new DataConversionOptions { Range = "A1:F200" });
workbook.SaveAsJson("sales.json");   // one key per sheet

JsonNode.Parse(File.ReadAllText("export.json"))!.SaveAsExcel("export.xlsx");
```

## See what the target could not hold

Every format has limits. TOML has no null, CSV has no nesting, XML attributes have no JSON equivalent, and an Excel cell holds 15 significant digits. Each `To{Target}` method has an overload that returns an `OfficeConversionReport`, and each `SaveAs{Target}` returns one:

```csharp
OfficeConversionReport report = JsonNode.Parse(File.ReadAllText("service.json"))!.SaveAsToml("service.toml");

foreach (OfficeDiagnostic diagnostic in report.Diagnostics)
{
    Console.WriteLine($"{diagnostic.Code} at {diagnostic.Location}: {diagnostic.Message}");
}
// null-dropped at service.owner: TOML has no null, so null values were omitted.
```

Codes are stable, so a pipeline can allow some and reject others. Set `DataConversionOptions.FailOnLoss` to throw `OfficeConversionException` before any output is written.

## Rows as a data reader

`StructuredDataReader` exposes the rows of JSON, JSON Lines, YAML, TOML or XML as a `DbDataReader`, ready for `SqlBulkCopy`, `DataTable.Load` or `ExcelSheet.InsertDataReader`:

```csharp
using System.Data;
using OfficeIMO.Data.Conversions;

using StructuredDataReader rows = StructuredDataReader.Open("orders.json", new DataConversionOptions { Select = "data.items" });
var table = new DataTable();
table.Load(rows);   // columns such as id, customer.name and lines[0].sku, typed from the data
```

The input is read twice: once to find the columns and their types, then to stream the rows.

## How data maps between formats

- **Rows** come from an array: the root, each document of a YAML stream or JSON Lines file, or the path in `Select`. A root object whose members are all arrays is several tables, one sheet each in Excel.
- **Columns** are paths such as `address.city` and `tags[0]`. `ArrayColumns = Join` puts scalar arrays in one cell instead.
- **CSV types** are inferred per column from canonical spellings only, so `007`, `+1`, `1,000` and `yes` stay text.
- **TOML** keeps integers, floats, booleans and all four date-time kinds typed. Comments above a key or after a value travel to and from YAML and XML.
- **XML** uses one convention: the root element is the top-level key, attributes are `@name`, mixed text is `#text`, repeated elements are arrays, and `xsi:nil="true"` is null. DTDs are prohibited.

The [design document](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.data-conversions-design.md) lists every mapping rule and loss code.

## From PowerShell and the command line

PSWriteOffice runs the same routes through `Convert-OfficeDocument`, with `-FailOnLoss` and the same report. See [configuration files and data conversion](/docs/pswriteoffice/config-data-formats/). The `officeimo` command line runs them through `officeimo workflow batch --target`, with `--require-no-loss`. See [batch conversion and saved jobs](/docs/cli/batch-and-jobs/#structured-data). The format engines have their own guides: [YAML](/docs/yaml/), [TOML](/docs/toml/) and [CSV](/docs/csv/).
