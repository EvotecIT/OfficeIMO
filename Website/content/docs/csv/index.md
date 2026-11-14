---
title: CSV Documents
description: "Create, load, validate and save CSV files in .NET with the OfficeIMO.CSV document model, with links to the CSV reading, writing and conversion guides."
order: 50
verified: 2026-11-14
---

The `OfficeIMO.CSV` package provides a fluent, strongly typed CSV document model for .NET. It supports reading, writing, validation, schema enforcement, streaming, and object mapping with zero external dependencies.

## Key Classes

| Class | Description |
|-------|-------------|
| `CsvDocument` | Root class for creating, loading, and saving CSV data. |
| `CsvRow` | Represents a single data row with typed column access. |
| `CsvSchema` | Defines column names, types, and validation rules. |
| `CsvValidator` | Validates rows against a schema. |
| `CsvRowWriter` | Writes objects, projected rows, and data readers without materializing a document. |
| `OfficeIMO.Data.RowMapper<T>` | Defines neutral explicit typed assignments for trimming and NativeAOT-sensitive code. |

## Creating a CSV Document

```csharp
using OfficeIMO.CSV;

var csv = CsvDocument.Create()
    .WithDelimiter(',')
    .WithHeader("Name", "Age", "City")
    .AddRow("Alice", "30", "New York")
    .AddRow("Bob", "25", "London")
    .AddRow("Carol", "35", "Tokyo");

csv.Save("people.csv");
```

## Creating from Objects

Generate a CSV document from any collection of objects. Column names are inferred from property names or dictionary keys:

```csharp
using OfficeIMO.CSV;

var employees = new[] {
    new { Name = "Alice", Department = "Engineering", Salary = 95000 },
    new { Name = "Bob", Department = "Design", Salary = 85000 },
    new { Name = "Carol", Department = "Marketing", Salary = 90000 },
};

var csv = CsvDocument.FromObjects(employees);
csv.Save("employees.csv");
```

You can customize the delimiter, culture, and encoding:

```csharp
using System.Globalization;
using OfficeIMO.CSV;

var employees = new[] { new { Name = "Alice", Salary = 95000.50m } };
var csv = CsvDocument.FromObjects(
    employees,
    delimiter: ';',
    culture: new CultureInfo("de-DE")
);
```

## Loading a CSV File

```csharp
using OfficeIMO.CSV;

var csv = CsvDocument.Load("data.csv");

foreach (CsvRow row in csv.AsEnumerable()) {
    Console.WriteLine($"{row.AsString("Name")}: {row.AsInt32("Age")}");
}
```

### Reader options

```csharp
using System.Text;
using OfficeIMO.CSV;

var csv = CsvDocument.Load("data.csv", new CsvLoadOptions {
    Delimiter = '\t',
    HasHeaderRow = true,
    Encoding = Encoding.UTF8
});
```

### Forward-only reading

`CsvLoadOptions` supplies parsing settings to `Load` and `Parse`.
`CsvReaderOptions` supplies parsing and projection settings to `OpenReader`.
Shared parsing concepts use the same property names and defaults on both types. For large files, open a forward-only `CsvDataReader` instead of
loading a document:

```csharp
using OfficeIMO.CSV;

using CsvDataReader reader = CsvDocument.OpenReader("large.csv");

while (reader.Read()) {
    Console.WriteLine(reader.GetString(reader.GetOrdinal("Name")));
}
```

## Schema and Validation

Define a schema to enforce column types and constraints:

```csharp
using System.Text.RegularExpressions;
using OfficeIMO.CSV;

var csv = CsvDocument.Load("people.csv");
var validated = csv
    .EnsureSchema(schema => schema
        .Column("Name").AsString().Required()
        .Column("Age").AsInt32().Required().Validate(v => (int)v! >= 0 && (int)v! <= 150, "Age must be between 0 and 150.")
        .Column("Email").AsString().Optional().Validate(
            v => v is null || Regex.IsMatch((string)v, @"^[\w.-]+@[\w.-]+\.\w+$"),
            "Email must be a valid address."))
    .Validate(out var errors);

foreach (var error in errors) {
    Console.WriteLine($"Row {error.RowIndex}, Column '{error.ColumnName}': {error.Message}");
}
```

## Object Mapping

Map CSV rows to strongly-typed objects:

```csharp
using OfficeIMO.CSV;

var csv = CsvDocument.Load("people.csv");
var people = csv.RowsAs<Person>().ToList();

foreach (var person in people) {
    Console.WriteLine($"{person.Name} ({person.Age}) lives in {person.City}");
}

public class Person {
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}
```

## Save Options

```csharp
using System.Text;
using OfficeIMO.CSV;

var csv = CsvDocument.Load("people.csv");
csv.Save("output.csv", new CsvSaveOptions {
    Delimiter = ',',
    IncludeHeader = true,
    Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)
});
```

## Writing to a Stream

```csharp
using OfficeIMO.CSV;

var csv = CsvDocument.Load("people.csv");
using var stream = new MemoryStream();
csv.Save(stream);
```

## Custom Delimiters

```csharp
using OfficeIMO.CSV;

// Tab-separated
var tsv = CsvDocument.Create().WithDelimiter('\t');

// Semicolon-separated (common in European locales)
var csv = CsvDocument.Create().WithDelimiter(';');

// Pipe-separated
var psv = CsvDocument.Create().WithDelimiter('|');
```

## Culture-Aware Formatting

```csharp
using System.Globalization;
using OfficeIMO.CSV;

var csv = CsvDocument.Create()
    .WithCulture(new CultureInfo("fr-FR"))
    .WithHeader("Produit", "Prix")
    .AddRow("Widget A", "9,99")   // French decimal separator
    .AddRow("Widget B", "14,99");
```

## Task guides

- [Read CSV into typed objects](/docs/csv/read-typed/)
- [Write CSV from objects and data readers](/docs/csv/write-data/)
- [CSV dialects and encodings](/docs/csv/dialects-encodings/)
- [Convert CSV to Excel](/docs/csv/convert-to-excel/)

## Convert to other formats

To turn CSV or TSV into JSON, JSON Lines, YAML, TOML, XML or an Excel worksheet, or the other way round, see [data conversions](/docs/data-conversions/).
