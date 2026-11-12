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
using OfficeIMO.Yaml;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddYamlHandler(new YamlReadOptions {
        ChunkRows = 100,
        IncludeMarkdown = true,
        LoadOptions = new YamlLoadOptions {
            MaxDepth = 16,
            MaxNodes = 10_000
        }
    })
    .Build();

foreach (var chunk in reader.Enumerate("deployment.yaml", new ReaderOptions {
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
var chunks = reader.Enumerate(stream, "values.yaml", new ReaderOptions {
    MaxChars = 3_000
}).ToList();
```

## What it emits

- Path/type/value rows for every scalar and empty collection, in document order. Aliases expand at the alias's path.
- Paths use exact key text (`$.metadata.name`, `$["app.kubernetes.io/name"]`, `$[1].spec` in multi-document streams), so `YamlDocument.Select(path)` returns the node a row came from.
- Types follow the document's schema: YAML 1.2 core by default, YAML 1.1 for `%YAML 1.1` documents, and explicit tags such as `!!str`.
- Chunks of `ChunkRows` rows with source `StartLine`/`EndLine`; each document of a multi-document stream starts a new chunk.
- Path and stream dispatch.
- A warning chunk with the engine's error code, message, line and column for malformed YAML or an exceeded limit, followed by the offending source line with a caret under the column. Findings of a successful load, such as each repeated key (the last one wins unless `YamlLoadOptions` is strict or sets `DuplicateKeys`) or YAML that only the default mode accepts, become warnings on the first chunk.

## Limits

`YamlReadOptions.LoadOptions` takes the engine's `YamlLoadOptions`: `MaxDepth`, `MaxNodes` (per document), `MaxAliasExpansion`, `MaxScalarLength`, `MaxInputBytes`, `Strict`, `DuplicateKeys` and `Schema`. Without it the engine defaults apply; use `YamlLoadOptions.UntrustedDefaults` for files from outside. `ReaderOptions.MaxInputBytes` is applied as well, and the lower value wins.

## Boundaries

- Reader adapter configuration belongs here.
- Object serialization, editing and merging belong to [`OfficeIMO.Yaml`](../OfficeIMO.Yaml/README.md).
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** none.
- **OfficeIMO:** `OfficeIMO.Yaml` parses and loads YAML and enforces its limits; `OfficeIMO.Reader.Core` owns chunk contracts, locations, and warnings.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
