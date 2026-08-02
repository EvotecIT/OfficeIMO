---
title: "OfficeIMO.CSV"
description: "Typed CSV workflows with schema validation, forward-only readers, zero third-party dependencies, and an executed NativeAOT parse scenario."
layout: product
product_color: "#0891b2"
install: "dotnet add package OfficeIMO.CSV"
nuget: "OfficeIMO.CSV"
docs_url: "/docs/csv/"
api_url: "/api/csv/"
meta.software.name: "OfficeIMO.CSV"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.CSV"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Why OfficeIMO.CSV?

OfficeIMO.CSV treats CSV files as first-class documents rather than raw text. Define a schema, map rows to typed objects automatically or with AOT-friendly delegates, validate on read, and stream through files of any size. Its checked-in NativeAOT smoke publishes a real parser, executes it, and verifies the resulting header schema.

## Features

- **Document-centric CSV model** — headers, rows, and metadata wrapped in a structured document object
- **Schema definition & validation** — declare column names, types, and constraints; reject invalid rows at parse time
- **Typed mapping** — map headers automatically for ordinary DTOs or use explicit delegates for trimming and NativeAOT
- **Forward-only reads for large files** — process rows through the standard `DbDataReader` contract
- **Sort, filter & transform** — chain LINQ-style operations directly on the CSV document
- **NativeAOT scenario in CI** — publishes and executes CSV parsing with schema readback instead of inferring compatibility from dependencies
- **Zero external dependencies** — ships as a single assembly with no third-party references

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
    .RowsAs<Employee>(map => map
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
    Console.WriteLine($"{emp.Name} — {emp.Department} — {emp.Salary:C}");
}

using var reader = CsvDocument.OpenDataReader("large-dataset.csv", new CsvLoadOptions
{
    HasHeaderRow = true,
    Culture = CultureInfo.InvariantCulture
});
while (reader.Read())
{
    Console.WriteLine(reader.GetString(reader.GetOrdinal("Name")));
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
