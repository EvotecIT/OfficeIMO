# OfficeIMO.Reader.Json - JSON reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Json)](https://www.nuget.org/packages/OfficeIMO.Reader.Json)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Json?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Json)

`OfficeIMO.Reader.Json` provides a modular JSON ingestion adapter for `OfficeIMO.Reader.Core`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Json
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddJsonHandler()
    .Build();
```

## Examples

### Convert JSON paths into chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddJsonHandler(new JsonReadOptions {
        ChunkRows = 100,
        MaxDepth = 16,
        MaxNodes = 100_000,
        IncludeMarkdown = true
    })
    .Build();

foreach (var chunk in reader.Enumerate("appsettings.json", new ReaderOptions {
    MaxInputBytes = 5L * 1024L * 1024L
})) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a JSON stream

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddJsonHandler()
    .Build();

await using var stream = File.OpenRead("payload.json");
var chunks = reader.Enumerate(stream, "payload.json", new ReaderOptions {
    MaxChars = 3_000
}).ToList();
```

## What it emits

- AST traversal through `System.Text.Json`.
- Path/type/value rows.
- Chunked structured output with optional Markdown tables.
- Path and stream dispatch.
- Warning chunks for malformed JSON.

## Options

| Option | Default | Meaning |
| --- | --- | --- |
| `ChunkRows` | 200 | Rows per emitted chunk. |
| `MaxDepth` | 64 | Maximum nesting of objects and arrays, with the root container at depth one. |
| `MaxNodes` | 1,000,000 | Maximum JSON values: objects, arrays, strings, numbers, booleans and nulls. |
| `MaxScalarLength` | 1,048,576 | Maximum characters in one decoded string or property name, or in one number as written. |
| `IncludeMarkdown` | `true` | Adds a Markdown table projection to each chunk. |

## Limits and diagnostics

The adapter checks all three limits in one forward pass over the input, before it builds the document or emits a chunk. Exceeding a limit throws `ReaderResourceLimitException` whose `LimitName` is the option name and whose `Maximum` is the configured value; it never returns a shortened value or a parse warning. Values below 1 throw `ArgumentOutOfRangeException`. Raise the limits for trusted larger inputs. `ReaderOptions.MaxInputBytes` and aggregate `ResourceLimits` apply independently.

Comments and trailing commas are accepted. Malformed JSON produces a single `json-warning-0000` chunk instead of throwing.

## Boundaries

- Reader adapter configuration belongs here.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** `System.Text.Json`.
- **OfficeIMO:** `OfficeIMO.Reader.Core` owns chunking, limits, locations, warnings, and result projection.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
