# OfficeIMO.Reader.Xml - XML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Xml)](https://www.nuget.org/packages/OfficeIMO.Reader.Xml)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Xml?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Xml)

`OfficeIMO.Reader.Xml` registers a modular XML ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Xml
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(replaceExisting: true);
```

## Examples

### Convert XML elements and attributes into chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(new XmlReadOptions {
    ChunkRows = 150,
    IncludeMarkdown = true
}, replaceExisting: true);

foreach (var chunk in DocumentReader.Read("configuration.xml")) {
    Console.WriteLine($"{chunk.Id}: {chunk.Location.Path}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read XML from a bounded stream

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xml;

DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler();

await using var stream = File.OpenRead("upload.xml");
var chunks = DocumentReader.Read(stream, "upload.xml", new ReaderOptions {
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

- Reader adapter registration belongs here.
- Shared extraction contracts belong in `OfficeIMO.Reader`.
- `OfficeIMO.Reader.Text` exists only as a compatibility orchestrator for structured text adapters.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
