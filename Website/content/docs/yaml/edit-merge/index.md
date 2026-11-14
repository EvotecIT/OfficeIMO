---
title: Edit and Merge YAML Without Losing Comments
description: Change values in YAML files by path, add and remove entries, set comments, and merge Helm-style overlays while every untouched comment, blank line and spelling stays as it was.
order: 53
verified: 2026-11-14
packages:
  - OfficeIMO.Yaml
cmdlets:
  - Open-OfficeYaml
  - Get-OfficeYamlNode
  - Set-OfficeYamlNode
  - Remove-OfficeYamlNode
  - Merge-OfficeDocument
limits: Mappings that a YAML merge key brings in are never edited; setting a key they provide adds an entry of its own.
search_terms:
  - yml
  - update yaml
  - preserve comments
  - helm values overlay
  - merge yaml files
  - yaml path
---

Load the file as a `YamlDocument` when a person will read the result. Saving copies every untouched part of the source and writes only what changed, so comments, blank lines, quoting, indentation, key order and spellings such as `TRUE` or `0x1F` survive everywhere else. An unchanged document saves back byte for byte.

## Change values by path

`Select` returns a node or `null`. `Set` creates missing mappings on the way, `Add` appends to the sequence at a path, and `Remove` returns `false` when nothing is there.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");

int replicas = document.Select("spec.replicas")?.GetValue<int>() ?? 1;
document.Set("spec.replicas", replicas + 1);                       // only the number changes in the file
document.Set("metadata.labels[\"app.kubernetes.io/version\"]", "1.29.2");
document.Add("spec.template.spec.containers[0].args", "--verbose");
document.Remove("spec.template.spec.containers[0].command");       // its comment lines go with it

document.Save();
```

Paths use `.name` for keys, `["key"]` with JSON escapes for keys that contain `.` or `[`, and `[n]` for zero-based sequence items. A value equal to the current one under the document's schema keeps its source text: setting `true` on `TRUE` changes nothing. New strings are quoted when a YAML 1.1 reader such as PyYAML could read them as something else (`yes`, `on`, `012`), so they stay strings.

### PowerShell

```powershell
Open-OfficeYaml -Path ./deployment.yaml {
    Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
    Set-OfficeYamlNode -Select 'resources.limits' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' })
    Remove-OfficeYamlNode -Select 'spec.template.spec.containers[0].command'
}

$document = Open-OfficeYaml -Path ./deployment.yaml
Get-OfficeYamlNode -Document $document -Select 'spec.replicas', 'metadata.name' |
    Select-Object -Property Path, Value, Line, Column
Close-OfficeYaml -Document $document
```

## Add entries and comments

`YamlMapping` and `YamlSequence` edit a subtree directly. `Add` and `Insert` return the node they placed, so comment, anchor and tag setters chain. New nodes follow their neighbors' indentation, sequence style and quoting.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");

YamlMapping container = document.Select("spec.template.spec.containers[0]")!.AsMapping();
container.Add("imagePullPolicy", "Always").SetComment("pinned by CI");
container.RenameKey("cmd", "command");

YamlSequence ports = container["ports"].AsSequence();
ports.Add(new YamlMapping { { "containerPort", 8443 }, { "name", "https" } });

foreach (YamlNode node in ports) {
    Console.WriteLine($"{node.GetPath()} at line {node.Start.Line}");
}

document.Save();
```

`CommentBefore` returns the comment lines above an entry and `Comment` the comment at the end of its line. Removing an entry removes the comment lines directly above it; a comment block separated by a blank line, such as a commented-out example, stays.

## Merge an overlay into a base file

`Merge` applies an overlay the way Helm applies `values.prod.yaml` over `values.yaml`. Mappings merge key by key; scalars are assigned; sequences are replaced by default. The returned `YamlMergeReport` lists the added, replaced and removed paths.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument values = YamlDocument.Load("values.yaml");
using YamlDocument production = YamlDocument.Load("values.prod.yaml");

YamlMergeReport report = values.Merge(production, new YamlMergeOptions {
    NullRemovesKey = true,                                // an overlay null removes the base key
    Sequences = YamlSequenceMerge.MergeByKey("name")      // merge list items with the same name
});
values.Save("values.merged.yaml");

foreach (string path in report.Added) Console.WriteLine("added " + path);
foreach (string path in report.Replaced) Console.WriteLine("changed " + path);
foreach (string path in report.Removed) Console.WriteLine("removed " + path);
```

`YamlSequenceMerge.Replace` (the default) and `Append` are the other sequence modes. A mapping replacing a scalar lets the overlay win and is reported in `report.Diagnostics`. To merge into one subtree, call `YamlMapping.Merge`, for example `values.Select("ingress")!.AsMapping().Merge(overlay)`.

### PowerShell

```powershell
Merge-OfficeDocument -Path ./values.yaml, ./values.prod.yaml -OutputPath ./values.merged.yaml -Options @{ NullRemovesKey = $true } -PassThruReport
```

## Normalize the layout

When you want the opposite of a minimal edit, such as a consistent indentation across files, save with `Reformat`. The whole document is written anew through the emitter, keeping values, anchors, tags and comments.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("values.yaml");
document.Save("values.formatted.yaml", new YamlSaveOptions {
    Reformat = true,
    WriterOptions = new YamlWriterOptions { MappingIndent = 2, SequenceIndent = 4, SequenceDashOffset = 2 }
});
```

### PowerShell

```powershell
$document = Open-OfficeYaml -Path ./values.yaml
Save-OfficeYaml -Document $document -Path ./values.formatted.yaml -WriteOptions @{ Reformat = $true }
Close-OfficeYaml -Document $document
```
