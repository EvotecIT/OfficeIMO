# OfficeIMO.Reader.Yaml - YAML reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Yaml)](https://www.nuget.org/packages/OfficeIMO.Reader.Yaml)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Yaml?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Yaml)

`OfficeIMO.Reader.Yaml` registers a modular YAML ingestion adapter for `OfficeIMO.Reader`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Yaml
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

DocumentReaderYamlRegistrationExtensions.RegisterYamlHandler(replaceExisting: true);
```

## Examples

### Convert YAML paths into chunks

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

DocumentReaderYamlRegistrationExtensions.RegisterYamlHandler(new YamlReadOptions {
    ChunkRows = 100,
    MaxDepth = 16,
    MaxNodes = 10_000,
    IncludeMarkdown = true
}, replaceExisting: true);

foreach (var chunk in DocumentReader.Read("deployment.yaml", new ReaderOptions {
    MaxInputBytes = 5L * 1024L * 1024L
})) {
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Read a YAML stream

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Yaml;

DocumentReaderYamlRegistrationExtensions.RegisterYamlHandler();

await using var stream = File.OpenRead("values.yaml");
var chunks = DocumentReader.Read(stream, "values.yaml", new ReaderOptions {
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

- Reader adapter registration belongs here.
- YAML object serialization/deserialization is intentionally out of scope.
- Shared extraction contracts belong in `OfficeIMO.Reader`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
