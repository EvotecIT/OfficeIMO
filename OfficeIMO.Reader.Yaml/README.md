# OfficeIMO.Reader.Yaml - YAML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Yaml)](https://www.nuget.org/packages/OfficeIMO.Reader.Yaml)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Yaml?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Yaml)

`OfficeIMO.Reader.Yaml` provides a modular YAML ingestion adapter for `OfficeIMO.Reader.Core`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Yaml
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddYamlHandler()
    .Build();
```

## Examples

### Convert YAML paths into chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddYamlHandler(new YamlReadOptions {
        ChunkRows = 100,
        MaxDepth = 16,
        MaxNodes = 10_000,
        IncludeMarkdown = true
    })
    .Build();

foreach (var chunk in reader.Read("deployment.yaml", new ReaderOptions {
    MaxInputBytes = 5L * 1024L * 1024L
})) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a YAML stream

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddYamlHandler()
    .Build();

await using var stream = File.OpenRead("values.yaml");
var chunks = reader.Read(stream, "values.yaml", new ReaderOptions {
    MaxChars = 3_000
}).ToList();
```

## What it emits

- YAML representation traversal through `YamlDotNet`.
- Path/type/value rows.
- Multi-document YAML stream support.
- Chunked structured output with optional Markdown tables.
- Path and stream dispatch.
- Warning chunks for malformed YAML.

## Boundaries

- Reader adapter configuration belongs here.
- YAML object serialization/deserialization is intentionally out of scope.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** YamlDotNet for YAML representation parsing.
- **OfficeIMO:** `OfficeIMO.Reader.Core` owns traversal projection, chunking, limits, locations, and warnings.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
