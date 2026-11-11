---
title: "OfficeIMO.Yaml"
description: "Read, edit, merge and write YAML 1.2 in .NET without losing comments or formatting, with a source-generated serializer and no third-party parser."
layout: product
meta.seo_title: "Lossless YAML editing and serialization for .NET"
product_color: "#ca8a04"
product_label: "YAML engine"
install: "dotnet add package OfficeIMO.Yaml"
nuget: "OfficeIMO.Yaml"
docs_url: "/docs/yaml/"
api_url: "/api/yaml/"
meta.software.name: "OfficeIMO.Yaml"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Yaml"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## Edit configuration files without rewriting them

`OfficeIMO.Yaml` loads a YAML file into a document you can edit and saves only what changed. Comments, blank lines, quoting, indentation, key order and spellings such as `TRUE` or `0x1F` stay exactly as they were everywhere else, so a pull request that changes one value shows one changed line.

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");
document.Set("spec.replicas", 3);
document.Set("spec.template.spec.containers[0].image", "nginx:1.27");
document.Save();   // only the two values change in the file
```

## Merge overlays the way Helm does

Layer per-environment values over a base file. The result is a minimal diff of the base, and a report lists every added, replaced and removed path.

```csharp
using var values = YamlDocument.Load("values.yaml");
using var production = YamlDocument.Load("values.prod.yaml");
YamlMergeReport report = values.Merge(production, new YamlMergeOptions { NullRemovesKey = true });
values.Save("values.merged.yaml");
```

## Serialize objects, including in NativeAOT

`YamlSerializer` follows the `System.Text.Json` API. A source generator ships in the package, so trimmed and NativeAOT applications get generated metadata instead of reflection.

```csharp
[YamlSerializable(typeof(AppConfig))]
internal partial class AppYamlContext : YamlSerializerContext { }

AppConfig? config = YamlSerializer.Deserialize(File.ReadAllText("app.yaml"), AppYamlContext.Default.AppConfig);
string yaml = YamlSerializer.Serialize(config!, AppYamlContext.Default.AppConfig);
```

## Features

- **Lossless document model**: unchanged files save byte for byte, including encoding and byte-order mark.
- **YAML 1.2.2**: checked against the yaml-test-suite, with Failsafe, JSON, Core and YAML 1.1 schemas, anchors, aliases, merge keys, tags and multi-document streams.
- **Real-world files**: reads the indentation and merge-key forms that CI, Compose and Helm files use, with a positioned warning for each; `Strict` applies the specification's exact rules.
- **Safe output**: new strings are quoted wherever a YAML 1.1 reader such as PyYAML would read them as something else, so `yes`, `on` and `012` stay strings.
- **Streaming reader and writer**: `YamlReader` allocates nothing per event, and `YamlWriter` output is deterministic.
- **Bounded input**: limits on depth, nodes, alias expansion, scalar size and input size, with `UntrustedDefaults` for files from outside.
- **No third-party parser**: depends only on `OfficeIMO.Core`.

## Compatibility

| Target framework | Supported |
|---|---|
| .NET 10.0 | Yes |
| .NET 8.0 | Yes |
| .NET Standard 2.0 | Yes |
| .NET Framework 4.7.2 | Yes |

## Related guides

| Guide | Description |
|---|---|
| [YAML documentation](/docs/yaml/) | Loading, editing, paths, merging, the serializer and options. |
| [Reader and extraction](/docs/reader/) | Turn YAML into path/type/value chunks with `OfficeIMO.Reader.Yaml`. |
| [Markdown](/docs/markdown/) | Front matter is read and written through the YAML engine. |
| [PSWriteOffice](/docs/pswriteoffice/) | `Open-OfficeYaml`, `Set-OfficeYamlNode`, `ConvertFrom-OfficeYaml` and the other YAML commands. |
