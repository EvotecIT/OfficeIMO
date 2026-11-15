---
title: Write YAML Files from .NET
description: Write YAML from .NET objects, build a new YAML document by path, stream large output with YamlWriter, and convert JSON or TOML to YAML.
order: 54
verified: 2026-11-14
commands:
  - officeimo convert
packages:
  - OfficeIMO.Yaml
  - OfficeIMO.Data.Conversions
cmdlets:
  - ConvertTo-OfficeYaml
  - Export-OfficeYaml
  - New-OfficeYaml
  - Set-OfficeYamlNode
search_terms:
  - yml
  - serialize yaml
  - object to yaml
  - json to yaml
  - create yaml file
---

Pick the writer by where the data comes from. `YamlSerializer` writes objects you already have. A new `YamlDocument` builds a file by path, which suits small configuration files you assemble from settings. `YamlWriter` is a forward-only emitter for large or generated output that should not be held in memory. All three write deterministic output and quote strings that YAML 1.1 readers would otherwise read as booleans or numbers.

## Write objects as YAML

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

var config = new ServiceConfig {
    Name = "web",
    Replicas = 3,
    Ports = { 80, 443 },
    Labels = { ["app"] = "web", ["tier"] = "yes" }
};

var options = new YamlSerializerOptions {
    PropertyNamingPolicy = OfficeNamingPolicy.CamelCase,
    DefaultIgnoreCondition = OfficeIgnoreCondition.WhenWritingNull
};
File.WriteAllText("service.yaml", YamlSerializer.Serialize(config, options));

public sealed class ServiceConfig {
    public string Name { get; set; } = "";
    public int Replicas { get; set; }
    public List<int> Ports { get; set; } = new();
    public Dictionary<string, string> Labels { get; set; } = new();
    public string? Description { get; set; }
}
```

The output is:

```yaml
name: web
replicas: 3
ports:
- 80
- 443
labels:
  app: web
  tier: 'yes'
```

`[OfficePropertyName]` sets a key and wins over the naming policy, and `[OfficeIgnore]` leaves a property out. With `ReferenceHandling = YamlReferenceHandling.Preserve`, an object reached twice is written as an anchor and an alias instead of failing as a cycle.

### PowerShell

```powershell
[ordered]@{ name = 'web'; replicas = 3; ports = 80, 443 } | ConvertTo-OfficeYaml

[ordered]@{ logging = [ordered]@{ level = 'Information' }; retries = 3 } |
    Export-OfficeYaml -Path ./settings.yaml
```

## Build a new YAML file by path

`YamlDocument.Create(path)` starts an empty document for a file without writing it. `Set` creates the mappings on the way, and comments can be set as you go.

### C#

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Create("values.yaml");
document.Set("image.repository", "nginx");
document.Set("image.tag", "1.27.2");
document.Set("replicaCount", 2);
document.Add("ingress.hosts", "example.com");
document.Select("image.tag")!.SetComment("updated by the release pipeline");
document.Save();
```

### PowerShell

```powershell
New-OfficeYaml -Path ./values.yaml {
    Set-OfficeYamlNode -Select 'image.repository' -Value 'nginx'
    Set-OfficeYamlNode -Select 'image.tag' -Value '1.27.2'
    Set-OfficeYamlNode -Select 'replicaCount' -Value 2
}
```

## Stream YAML with YamlWriter

`YamlWriter` writes events in order to a stream or `TextWriter`. It allocates nothing per call beyond its pooled buffer, never folds long lines, and writes multi-line strings as literal block scalars.

### C#

```csharp
using OfficeIMO.Yaml;

using FileStream stream = File.Create("inventory.yaml");
using var writer = new YamlWriter(stream);   // UTF-8 without a byte-order mark

writer.WriteStartSequence();
foreach (string host in new[] { "web-01", "web-02", "db-01" }) {
    writer.WriteStartMapping();
    writer.WriteKey("host");
    writer.WriteScalar(host);
    writer.WriteKey("enabled");
    writer.WriteScalar("yes");                // written 'yes', so YAML 1.1 readers keep it a string
    writer.WriteKey("ports");
    writer.WriteStartSequence(YamlStyle.Flow);
    writer.WriteScalar(80);
    writer.WriteScalar(443);
    writer.WriteEndSequence();
    writer.WriteEndMapping();
}
writer.WriteEndSequence();
```

`YamlWriterOptions` sets the mapping indent, sequence indent, dash offset and line break: 2, 2 and 0 write `key:\n- item`, while 2, 4 and 2 write `key:\n  - item`.

## Convert JSON or TOML to YAML

Structured-data conversion streams events from one format's reader to another's writer, so numbers keep their exact text and comments travel between YAML and TOML.

### C#

```csharp
using System.Text.Json.Nodes;
using OfficeIMO;
using OfficeIMO.Data.Conversions;
using OfficeIMO.Toml;

JsonNode.Parse(File.ReadAllText("appsettings.json")).SaveAsYaml("appsettings.yaml");

using TomlDocument cargo = TomlDocument.Load("Cargo.toml");
OfficeConversionReport report = cargo.SaveAsYaml("cargo.yaml");   // comments travel with their keys
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) Console.WriteLine(diagnostic.Message);
```

### CLI

```sh
officeimo convert appsettings.json appsettings.yaml
officeimo convert Cargo.toml cargo.yaml --require-no-loss
```

See [data conversions](/docs/data-conversions/) for every source and target format and the loss report.
