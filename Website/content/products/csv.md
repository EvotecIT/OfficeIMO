---
title: "OfficeIMO.CSV"
description: "Typed CSV workflows with schema validation and streaming. AOT-friendly, trimming-safe, and zero dependencies."
layout: product
product_color: "#0891b2"
install: "dotnet add package OfficeIMO.CSV"
nuget: "OfficeIMO.CSV"
docs_url: "/docs/csv/"
api_url: ""
---

## Why OfficeIMO.CSV?

OfficeIMO.CSV treats CSV files as first-class documents rather than raw text. Define a schema, map rows to typed objects without reflection, validate on read, and stream through files of any size. It is AOT-friendly and trimming-safe by design, with zero external dependencies.

## Features

- **Document-centric CSV model** -- headers, rows, and metadata wrapped in a structured document object
- **Schema definition & validation** -- declare column names, types, and constraints; reject invalid rows at parse time
- **Typed mapping without reflection** -- map rows to POCOs using compile-time delegates instead of runtime reflection
- **Streaming mode for large files** -- process millions of rows with constant memory using `IAsyncEnumerable<T>`
- **Sort, filter & transform** -- chain LINQ-style operations directly on the CSV document
- **AOT-friendly & trimming-safe** -- compatible with Native AOT and IL trimming in .NET 8+
- **Zero external dependencies** -- ships as a single assembly with no third-party references

## Quick start

```csharp
using OfficeIMO.CSV;

// Define a schema
var schema = new CsvSchema()
    .AddColumn("Name", CsvColumnType.String, required: true)
    .AddColumn("Department", CsvColumnType.String)
    .AddColumn("Salary", CsvColumnType.Decimal, required: true)
    .AddColumn("StartDate", CsvColumnType.Date);

// Read and validate a CSV file
using var document = CsvDocument.Open("employees.csv", schema);

Console.WriteLine($"Rows: {document.Rows.Count}");
Console.WriteLine($"Valid: {document.ValidationErrors.Count == 0}");

// Typed mapping without reflection
var employees = document.MapRows(row => new
{
    Name = row.GetString("Name"),
    Department = row.GetString("Department"),
    Salary = row.GetDecimal("Salary"),
    StartDate = row.GetDate("StartDate")
});

// Filter and transform
var highEarners = employees
    .Where(e => e.Salary > 100_000m)
    .OrderByDescending(e => e.Salary);

foreach (var emp in highEarners)
{
    Console.WriteLine($"{emp.Name} -- {emp.Department} -- {emp.Salary:C}");
}

// Stream large files with constant memory
await foreach (var row in CsvDocument.StreamAsync("large-dataset.csv", schema))
{
    // Process each row individually
    ProcessRow(row);
}
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.CSV runs on Windows, Linux, and macOS. It handles RFC 4180 compliant files as well as common real-world variations (quoted fields, embedded newlines, BOM markers).
