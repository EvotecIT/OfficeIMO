---
title: CSV Documents
description: Overview of the OfficeIMO.CSV package for strongly-typed CSV document workflows.
order: 50
---

# CSV Documents

The `OfficeIMO.CSV` package provides a fluent, strongly-typed CSV document model for .NET. It supports reading, writing, validation, schema enforcement, streaming, and object mapping -- all with zero external dependencies.

## Key Classes

| Class | Description |
|-------|-------------|
| `CsvDocument` | Root class for creating, loading, and saving CSV data. |
| `CsvRow` | Represents a single data row with typed column access. |
| `CsvSchema` | Defines column names, types, and validation rules. |
| `CsvValidator` | Validates rows against a schema. |
| `CsvWriter` | Low-level writer for streaming CSV output. |
| `CsvParser` | Low-level parser for streaming CSV input. |
| `CsvMapper` | Maps CSV rows to/from strongly-typed objects. |
| `CsvStreamingSource` | Lazy streaming source for large files. |

## Creating a CSV Document

```csharp
using OfficeIMO.CSV;

var csv = new CsvDocument()
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

var csv = CsvDocument.FromObjects(
    employees,
    delimiter: ';',
    culture: new CultureInfo("de-DE")
);
```

## Loading a CSV File

```csharp
var csv = CsvDocument.Load("data.csv");

foreach (var row in csv.Rows) {
    Console.WriteLine($"{row["Name"]}: {row["Age"]}");
}
```

### Load Options

```csharp
using System.Text;

var csv = CsvDocument.Load("data.csv", new CsvLoadOptions {
    Delimiter = '\t',
    HasHeaderRow = true,
    Encoding = Encoding.UTF8,
    Mode = CsvLoadMode.InMemory
});
```

### Streaming Mode

For large files, use streaming mode to avoid loading everything into memory:

```csharp
var csv = CsvDocument.Load("large.csv", new CsvLoadOptions {
    Mode = CsvLoadMode.Stream
});

foreach (var row in csv.Rows) {
    // Rows are read one at a time from disk
    Console.WriteLine(row.Get<string>("Name"));
}
```

## Schema and Validation

Define a schema to enforce column types and constraints:

```csharp
using System.Text.RegularExpressions;

var validated = csv
    .EnsureSchema(schema => schema
        .Column("Name").AsString().Required()
        .Column("Age").AsInt32().Required().Validate(v => (int)v! >= 0 && (int)v! <= 150, "Age must be between 0 and 150.")
        .Column("Email").AsString().Optional().Validate(
            v => v is null || Regex.IsMatch((string)v, @"^[\w.-]+@[\w.-]+\.\w+$"),
            "Email must be a valid address."))
    .Validate(out var errors);

foreach (var error in errors) {
    Console.WriteLine($"Row {error.RowIndex}, Column '{error.Column}': {error.Message}");
}
```

## Object Mapping

Map CSV rows to strongly-typed objects:

```csharp
public class Person {
    public string Name { get; set; }
    public int Age { get; set; }
    public string City { get; set; }
}

var csv = CsvDocument.Load("people.csv");
var people = csv.Map<Person>(map => map
    .FromColumn<string>("Name", (person, value) => person.Name = value)
    .FromColumn<int>("Age", (person, value) => person.Age = value)
    .FromColumn<string>("City", (person, value) => person.City = value)
).ToList();

foreach (var person in people) {
    Console.WriteLine($"{person.Name} ({person.Age}) lives in {person.City}");
}
```

## Save Options

```csharp
using System.Text;

csv.Save("output.csv", new CsvSaveOptions {
    Delimiter = ',',
    IncludeHeader = true,
    Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)
});
```

## Writing to a Stream

```csharp
using var stream = new MemoryStream();
csv.Save(stream);
```

## Custom Delimiters

```csharp
// Tab-separated
var tsv = new CsvDocument().WithDelimiter('\t');

// Semicolon-separated (common in European locales)
var csv = new CsvDocument().WithDelimiter(';');

// Pipe-separated
var psv = new CsvDocument().WithDelimiter('|');
```

## Culture-Aware Formatting

```csharp
using System.Globalization;

var csv = new CsvDocument()
    .WithCulture(new CultureInfo("fr-FR"))
    .WithHeader("Produit", "Prix")
    .AddRow("Widget A", "9,99")   // French decimal separator
    .AddRow("Widget B", "14,99");
```
