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
        IncludeMarkdown = true
    })
    .Build();

foreach (var chunk in reader.Read("appsettings.json", new ReaderOptions {
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
var chunks = reader.Read(stream, "payload.json", new ReaderOptions {
    MaxChars = 3_000
}).ToList();
```

## What it emits

- AST traversal through `System.Text.Json`.
- Path/type/value rows.
- Chunked structured output with optional Markdown tables.
- Path and stream dispatch.
- Warning chunks for malformed JSON.

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
