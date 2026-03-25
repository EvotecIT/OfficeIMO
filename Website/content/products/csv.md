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
- **Streaming mode for large files** -- process millions of rows with constant memory using forward-only enumeration
- **Sort, filter & transform** -- chain LINQ-style operations directly on the CSV document
- **AOT-friendly & trimming-safe** -- compatible with Native AOT and IL trimming in .NET 8+
- **Zero external dependencies** -- ships as a single assembly with no third-party references

## Quick start

```csharp
using OfficeIMO.CSV;
using System.Globalization;

var document = CsvDocument.Load("employees.csv")
    .EnsureSchema(schema => schema
        .Column("Name").AsString().Required()
        .Column("Department").AsString().Optional()
        .Column("Salary").AsType(typeof(decimal)).Required()
        .Column("StartDate").AsDateTime().Optional()
    )
    .ValidateOrThrow();

var employees = document
    .Map<Employee>(map => map
        .FromColumn<string>("Name", (employee, value) => employee with { Name = value })
        .FromColumn<string>("Department", (employee, value) => employee with { Department = value })
        .FromColumn<decimal>("Salary", (employee, value) => employee with { Salary = value })
        .FromColumn<DateTime>("StartDate", (employee, value) => employee with { StartDate = value })
    )
    .ToList();

// Filter and transform
var highEarners = employees
    .Where(e => e.Salary > 100_000m)
    .OrderByDescending(e => e.Salary);

foreach (var emp in highEarners)
{
    Console.WriteLine($"{emp.Name} -- {emp.Department} -- {emp.Salary:C}");
}

foreach (var row in CsvDocument.Load("large-dataset.csv", new CsvLoadOptions
{
    Mode = CsvLoadMode.Stream,
    HasHeaderRow = true,
    Culture = CultureInfo.InvariantCulture
}).AsEnumerable())
{
    Console.WriteLine(row.Get<string>("Name"));
}

public sealed record Employee
{
    public string Name { get; init; } = string.Empty;
    public string Department { get; init; } = string.Empty;
    public decimal Salary { get; init; }
    public DateTime StartDate { get; init; }
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

## Related guides

| Guide | Description |
|-------|-------------|
| [CSV documentation](/docs/csv/) | Start with the package overview and document model. |
| [AOT and trimming](/docs/advanced/aot-trimming/) | Keep CSV tooling lean for Native AOT and trimmed deployments. |
| [Reader and extraction](/docs/reader/) | Feed CSV and other document types into one ingestion workflow. |
| [Getting started](/docs/getting-started/) | Review install and package-selection guidance across the suite. |
