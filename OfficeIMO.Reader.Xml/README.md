# OfficeIMO.Reader.Xml - XML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Xml)](https://www.nuget.org/packages/OfficeIMO.Reader.Xml)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Xml?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Xml)

`OfficeIMO.Reader.Xml` provides a modular XML ingestion adapter for `OfficeIMO.Reader.Core`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Xml
```

## Configure

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddXmlHandler()
    .Build();
```

## Examples

### Convert XML elements and attributes into chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddXmlHandler(new XmlReadOptions {
        ChunkRows = 150,
        IncludeMarkdown = true
    })
    .Build();

foreach (var chunk in reader.Read("configuration.xml")) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.Path}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read XML from a bounded stream

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddXmlHandler()
    .Build();

await using var stream = File.OpenRead("upload.xml");
var chunks = reader.Read(stream, "upload.xml", new ReaderOptions {
    MaxInputBytes = 10L * 1024L * 1024L,
    MaxChars = 4_000
}).ToList();
```

## What it emits

- XML tree traversal to element/attribute path rows.
- Chunked structured output with optional Markdown tables.
- Path and stream dispatch.
- Warning chunks for malformed XML.

## Boundaries

- Reader adapter configuration belongs here.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond platform XML APIs.
- **OfficeIMO:** `OfficeIMO.Reader.Core` owns traversal projection, chunking, limits, locations, and warnings.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
