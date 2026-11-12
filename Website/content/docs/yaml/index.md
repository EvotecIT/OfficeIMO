---
title: YAML Documents
description: OfficeIMO.Yaml engine overview: lossless YAML editing, overlay merging, streaming and object serialization, with examples, options, limits and API links.
order: 51
---

`OfficeIMO.Yaml` is OfficeIMO's own YAML 1.2 engine. It reads, edits, merges and writes YAML without changing anything you didn't change, and converts .NET objects with generated or reflection metadata. It depends only on `OfficeIMO.Core`.

```powershell
dotnet add package OfficeIMO.Yaml
```

## Key classes

| Class | Description |
|-------|-------------|
| `YamlDocument` | A YAML file or stream you can edit and save losslessly. |
| `YamlMapping`, `YamlSequence`, `YamlScalar`, `YamlAlias` | The nodes of a document. |
| `YamlMergeReport` | The added, replaced and removed paths of an overlay merge, with diagnostics. |
| `YamlReader` | A forward-only pull reader that allocates nothing per event. |
| `YamlWriter` | A forward-only emitter with deterministic output. |
| `YamlSerializer` | Converts objects to and from YAML, following the `System.Text.Json` API. |
| `YamlSerializerContext` | Generated serializer metadata for trimmed and NativeAOT applications. |

## Loading and saving

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");
int replicas = document.Select("spec.replicas")!.GetValue<int>();
document.Save();   // unchanged: the same bytes as the input
```

`Load` reads a file or stream, `LoadAsync` does the same asynchronously and `Parse` takes a string. `Create(path)` starts an empty document for a file. `Save()` writes to the file the document was created or loaded with; `Save(path)` replaces a file atomically. The encoding and byte-order mark are detected and kept.

## Editing

```csharp
using YamlDocument document = YamlDocument.Load("deployment.yaml");
document.Set("spec.replicas", 3);
YamlMapping container = document.Select("spec.template.spec.containers[0]")!.AsMapping();
container.Add("imagePullPolicy", "Always").SetComment("pinned by CI");
container.Remove("command");
document.Save();
```

Saving copies every untouched part of the source and writes only what changed. A value equal to the current one keeps its spelling, new nodes follow the indentation and quoting of their neighbors, and removing an entry removes its own comments while its neighbors keep theirs. New strings are quoted wherever a YAML 1.1 reader such as PyYAML would read them as something else, so `yes`, `on` and `012` stay strings.

## Paths

| Step | Selects |
|---|---|
| `name`, `.name` | The value of the key whose text is `name`, whatever its quoting. |
| `["key"]` | Any key, with JSON escapes: `labels["app.kubernetes.io/name"]`. |
| `[n]` | The zero-based item of a sequence. |
| `$`, `$[n]` | The starting node; `$[n]` picks a document of a multi-document stream. |

`Select` returns a node or `null`, `Set` creates missing mappings on the way, `Add` appends to a sequence and `Remove` returns `false` when nothing is there. `node.GetPath()` writes the same form, and `OfficeIMO.Reader.Yaml` reports the same paths.

## Merging overlays

```csharp
using var values = YamlDocument.Load("values.yaml");
using var production = YamlDocument.Load("values.prod.yaml");
YamlMergeReport report = values.Merge(production, new YamlMergeOptions {
    NullRemovesKey = true,
    Sequences = YamlSequenceMerge.MergeByKey("name")
});
values.Save("values.merged.yaml");
```

Mappings merge key by key. Sequences are replaced by default, as in Helm, or appended or merged by key. With `NullRemovesKey`, an overlay `null` removes the base key. The merged file differs from the base only where the report says it does.

## Serializing objects

```csharp
[YamlSerializable(typeof(AppConfig))]
internal partial class AppYamlContext : YamlSerializerContext { }

var context = new AppYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase });
AppConfig? config = YamlSerializer.Deserialize(File.ReadAllText("app.yaml"), context.AppConfig);
string yaml = YamlSerializer.Serialize(config!, context.AppConfig);
```

The source generator ships in the package. Overloads that take `YamlSerializerOptions` alone use reflection metadata, which is off in `PublishAot` and `PublishTrimmed` applications. Polymorphism works only through `[YamlDerivedType]` registrations; a tag in the input never chooses a .NET type by itself.

## Options and diagnostics

| Options | Used by | Settings |
|---|---|---|
| `YamlLoadOptions` | `Load`, `LoadAsync`, `Parse` | `Schema`, `Strict`, `DuplicateKeys`, limits, `UntrustedDefaults` |
| `YamlReaderOptions` | `YamlReader`, the serializer | `Schema`, `Strict`, `IncludeComments`, limits, `UntrustedDefaults` |
| `YamlSaveOptions` | `Save`, `SaveAsync` | `Reformat`, `WriterOptions` |
| `YamlWriterOptions` | `YamlWriter`, the serializer | Mapping indent, sequence indent, dash offset, line break |
| `YamlMergeOptions` | `Merge` | `Sequences`, `NullRemovesKey` |

By default the engine reads YAML the way real files use it: flow collections and quoted continuations indented no deeper than their key, merge keys without a `%YAML 1.1` directive, and repeated keys, where the last one wins. Each is reported in `Diagnostics` with its line and column. `Strict` applies the specification's exact rules instead.

Malformed input throws `YamlException` with a stable `Code`, line and column; `ToDiagnostic()` returns it as an `OfficeDiagnostic`. Exceeding `MaxDepth`, `MaxNodes`, `MaxAliasExpansion`, `MaxScalarLength` or `MaxInputBytes` throws `YamlLimitExceededException`. Use `YamlLoadOptions.UntrustedDefaults` for files from outside.

## PowerShell

PSWriteOffice builds its YAML commands on the same API:

```powershell
Open-OfficeYaml -Path .\deployment.yaml {
    Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
}
```

`New-`, `Open-`, `Save-` and `Close-OfficeYaml` manage documents, `Get-`, `Set-` and `Remove-OfficeYamlNode` edit by path, `ConvertFrom-` and `ConvertTo-OfficeYaml` follow `ConvertFrom-Json` and `ConvertTo-Json`, `Import-` and `Export-OfficeYaml` read and write objects through files, `Test-OfficeYaml` validates, and `Merge-OfficeDocument` merges YAML overlays.

## Further reading

- [OfficeIMO.Yaml README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Yaml/README.md): the complete API guide, including comments, anchors, schemas and supported serializer types.
- [YAML engine design](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.yaml-engine-design.md): architecture, editing rules and qualification.
- [OfficeIMO.Yaml API reference](/api/yaml/)
