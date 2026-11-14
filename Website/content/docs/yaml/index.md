---
title: YAML Documents
description: Read, edit, merge, validate and write YAML 1.2 files in .NET with OfficeIMO.Yaml, keeping comments and formatting, with matching PowerShell cmdlets.
order: 51
verified: 2026-11-14
commands:
  - officeimo read
packages:
  - OfficeIMO.Yaml
cmdlets:
  - Open-OfficeYaml
  - Get-OfficeYamlNode
  - Test-OfficeYaml
limits: YamlLoadOptions bound depth (256), scalar length (16 MiB), nodes (10,000,000) and alias expansion (100,000); UntrustedDefaults lowers them for files from outside.
search_terms:
  - yml
  - yaml parser
  - read yaml c#
  - validate yaml
  - yaml lint
---

`OfficeIMO.Yaml` is OfficeIMO's own YAML 1.2 engine. It depends only on `OfficeIMO.Core` and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

```sh
dotnet add package OfficeIMO.Yaml
```

It has four entry points:

| Type | Use it to |
|-------|-------------|
| `YamlDocument` | Load, edit and save a file losslessly. Unchanged input saves back byte for byte. |
| `YamlSerializer` | Convert .NET objects to and from YAML, with reflection or a source-generated `YamlSerializerContext`. |
| `YamlReader` | Read events forward-only, without allocating per event. |
| `YamlWriter` | Emit YAML forward-only with deterministic output. |

The task guides cover [typed reading](/docs/yaml/read-typed/), [editing and merging](/docs/yaml/edit-merge/) and [writing](/docs/yaml/write/).

## Load a YAML file and read values

`Load` reads a file or stream, `LoadAsync` does the same asynchronously and `Parse` takes a string. The encoding and byte-order mark are detected and kept.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");

string? name = document.Select("metadata.name")?.AsString();
int replicas = document.Select("spec.replicas")?.GetValue<int>() ?? 1;
string? image = document.Select("spec.template.spec.containers[0].image")?.AsString();
Console.WriteLine($"{name}: {replicas} x {image}");

YamlMapping labels = document.Select("metadata.labels")!.AsMapping();
foreach (YamlMappingEntry entry in labels) {
    Console.WriteLine($"{entry.Key.AsString()} = {entry.Value.AsString()}");
}
```

Typed reads go through the document's schema: `GetValue<T>()` and `TryGetValue<T>()` cover `string`, `bool`, `int`, `long`, `double`, `float` and `decimal`. Aliases and `<<` merge keys are followed by lookups and paths without being flattened into the tree. Dispose the document after its last use; it returns pooled storage.

### PowerShell

```powershell
$document = Open-OfficeYaml -Path ./deployment.yaml
(Get-OfficeYamlNode -Document $document -Select 'spec.template.spec.containers[0].image').Value
Close-OfficeYaml -Document $document
```

### CLI

Extract a YAML file as Markdown or Reader JSON, with one path and value per entry:

```sh
officeimo read deployment.yaml --format json
```

## Validate YAML files

Malformed input throws `YamlException` with a stable `Code`, line and column. `ToDiagnostic()` turns it into an `OfficeDiagnostic`, so failed files can be reported next to the warnings of files that load.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

foreach (string path in Directory.EnumerateFiles("deploy", "*.yaml", SearchOption.AllDirectories)) {
    IReadOnlyList<OfficeDiagnostic> diagnostics;
    try {
        using YamlDocument document = YamlDocument.Load(path, new YamlLoadOptions { Strict = true });
        diagnostics = document.Diagnostics;
    } catch (YamlException exception) {
        diagnostics = new[] { exception.ToDiagnostic() };
    }

    foreach (OfficeDiagnostic diagnostic in diagnostics) {
        Console.WriteLine($"{path}: {diagnostic.Code} {diagnostic.Message}");
    }
}
```

By default the engine reads YAML the way real files use it: flow collections and quoted continuations indented no deeper than their key, merge keys without a `%YAML 1.1` directive, and repeated keys, where the last one wins. Each is a warning in `Diagnostics` with its line and column. `Strict` applies the specification's exact rules instead.

### PowerShell

```powershell
Get-ChildItem ./deploy -Recurse -Filter *.yaml | Test-OfficeYaml -Detailed |
    Where-Object { -not $_.IsValid } |
    Select-Object -Property Path, Line, Column, Code, Message
```

## Choose a schema and limits

| Options | Used by | Settings |
|---|---|---|
| `YamlLoadOptions` | `Load`, `LoadAsync`, `Parse` | `Schema`, `Strict`, `DuplicateKeys`, limits; `UntrustedDefaults` |
| `YamlReaderOptions` | `YamlReader`, the serializer | `Schema`, `Strict`, `IncludeComments`, limits; `UntrustedDefaults` |
| `YamlSaveOptions` | `Save`, `SaveAsync` | `Reformat`, `WriterOptions` |
| `YamlWriterOptions` | `YamlWriter`, the serializer | `MappingIndent`, `SequenceIndent`, `SequenceDashOffset`, `NewLine` |
| `YamlMergeOptions` | `Merge` | `Sequences`, `NullRemovesKey` |

The default `Core` schema resolves `null`, `true`/`false`, integers and floats. `Yaml11` adds `yes`/`no`/`on`/`off`, `0`-prefixed octals and sexagesimal numbers, and a `%YAML 1.1` directive selects it for that document. `Json` and `Failsafe` are stricter.

### C#

```csharp
using OfficeIMO.Yaml;

YamlLoadOptions options = YamlLoadOptions.UntrustedDefaults;
options.Schema = YamlSchema.Yaml11;   // the file comes from a tool that assumes YAML 1.1

using YamlDocument document = YamlDocument.Load("legacy.yaml", options);
bool enabled = document.Select("feature.enabled")?.GetValue<bool>() ?? false;   // "yes" reads as true
Console.WriteLine(enabled);
```

### PowerShell

```powershell
$document = Open-OfficeYaml -LiteralPath ./uploads/values.yaml -Untrusted -Strict
Get-OfficeYamlNode -Document $document -Select 'image.tag'
Close-OfficeYaml -Document $document
```

Exceeding `MaxDepth`, `MaxNodes`, `MaxAliasExpansion`, `MaxScalarLength` or `MaxInputBytes` throws `YamlLimitExceededException` with the limit's name. The [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Yaml/README.md) covers anchors, tags, schemas and every supported serializer type.
