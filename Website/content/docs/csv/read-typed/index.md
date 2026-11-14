---
title: Read CSV into Typed Objects
description: Map CSV rows to .NET objects, read typed values through a forward-only data reader, and infer column types with OfficeIMO.CSV, PSWriteOffice, and the officeimo CLI.
order: 51
packages:
  - OfficeIMO.CSV
cmdlets:
  - Import-OfficeCsv
limits: "Input-size, field-length, and decompression limits are on `CsvLoadOptions` and `CsvReaderOptions`."
---

Choose the read path by what happens next. Load a `CsvDocument` when you will edit, sort, filter, or save the rows. Open a `CsvDataReader` when you read once, from start to end: it does not keep the file in memory, and it is a `DbDataReader`, so it plugs into anything that accepts one.

## Map rows to objects

`RowsAs<T>` matches headers to writable properties or constructor parameters. Matching ignores case, spaces, and punctuation, so `Order Date` maps to `OrderDate`.

### C#

```csharp
using OfficeIMO.CSV;

List<Person> people = CsvDocument.Load("people.csv")
    .RowsAs<Person>()
    .ToList();

foreach (Person person in people) {
    Console.WriteLine($"{person.Name} ({person.Age}) lives in {person.City}");
}

public sealed class Person {
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\people.csv | Where-Object City -eq 'London'
Import-OfficeCsv -Path .\people.csv -As Hashtable | ForEach-Object { $_['Name'] }
```

## Stream a large file

The forward-only reader never loads the whole file. `RowsAs<T>` works on it too, through `OfficeIMO.Data`.

### C#

```csharp
using OfficeIMO.CSV;
using OfficeIMO.Data;

using CsvDataReader reader = CsvDocument.OpenReader("events.csv");
int errors = 0;
foreach (LogEvent item in reader.RowsAs<LogEvent>()) {
    if (item.Level == "Error") errors++;
}
Console.WriteLine($"{errors} errors");

public sealed record LogEvent(DateTime Time, string Level, string Message);
```

On .NET 8 and later, `CsvDocument.OpenReaderAsync` and `RowsAsAsync<T>` read with asynchronous I/O and observe cancellation between rows.

## Infer column types

With `InferSchema`, the reader samples the first rows and returns typed values (`int`, `decimal`, `DateTime`, `bool`) instead of strings.

### C#

```csharp
using OfficeIMO.CSV;

using CsvDataReader reader = CsvDocument.OpenReader("orders.csv",
    new CsvReaderOptions { InferSchema = true, SchemaSampleRows = 1000 });

int amount = reader.GetOrdinal("Amount");
decimal total = 0;
while (reader.Read()) {
    total += reader.GetDecimal(amount);
}
Console.WriteLine(total);
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\orders.csv -InferSchema | Measure-Object Amount -Sum
```

### CLI

Print the inferred schema of a file:

```sh
officeimo tabular schema orders.csv
```
