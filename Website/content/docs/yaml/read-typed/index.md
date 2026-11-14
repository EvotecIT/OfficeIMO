---
title: Read YAML into Typed Objects
description: Deserialize YAML files into .NET classes and records with OfficeIMO.Yaml, use source-generated contexts for trimmed and NativeAOT apps, and read multi-document streams.
order: 52
verified: 2026-11-14
packages:
  - OfficeIMO.Yaml
cmdlets:
  - Import-OfficeYaml
  - ConvertFrom-OfficeYaml
limits: Polymorphic types need [OfficeDerivedType] registrations; a tag in the input never chooses a .NET type by itself.
search_terms:
  - yml
  - deserialize yaml
  - yaml to object
  - yaml config c#
  - kubernetes manifest
  - nativeaot yaml
---

`YamlSerializer` maps YAML to .NET objects with an API that follows `System.Text.Json`: `Deserialize<T>` reads, `Serialize` writes, and `YamlSerializerOptions` sets the naming policy and limits. Use it when you want the values of a configuration file as typed objects. When you also need to change the file and keep its comments, load it as a `YamlDocument` instead (see [Edit and Merge YAML](/docs/yaml/edit-merge/)); the two combine, as the last section shows.

## Read a YAML file into a class

Keys match properties through the naming policy, unknown keys are skipped, and scalars are read by the target type whatever their quoting, so `"8080"` reads into an `int`.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

var options = new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase };

ServiceConfig? config = YamlSerializer.Deserialize<ServiceConfig>(File.ReadAllText("service.yaml"), options);
Console.WriteLine($"{config!.Name}: {config.Replicas} replicas, ports {string.Join(", ", config.Ports)}");

public sealed class ServiceConfig {
    public string Name { get; set; } = "";
    public int Replicas { get; set; } = 1;
    public List<int> Ports { get; set; } = new();
    public Dictionary<string, string> Labels { get; set; } = new();
    public TimeSpan Timeout { get; set; }
}
```

Constructor parameters bind to properties of the same name, so records work too. C# `required` members and properties marked `[OfficeRequired]` must be present; a missing one throws `YamlException` with code `MissingRequiredMember` and the path of the mapping.

### PowerShell

`Import-OfficeYaml` returns `PSCustomObject` values (or hashtables with `-AsHashtable`):

```powershell
$settings = Import-OfficeYaml -Path ./service.yaml
$settings.ports | ForEach-Object { "Listening on $_" }

$config = Get-Content -Path ./service.yaml -Raw | ConvertFrom-OfficeYaml -AsHashtable
$config['labels']['app']
```

## Use a source-generated context

The reflection overloads are turned off in `PublishAot` and `PublishTrimmed` applications. Declare a context and pass its generated metadata instead. The source generator ships in the `OfficeIMO.Yaml` package; nothing else is installed.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

var context = new AppYamlContext(new YamlSerializerOptions {
    PropertyNamingPolicy = OfficeNamingPolicy.CamelCase
});

AppConfig? config = YamlSerializer.Deserialize(File.ReadAllText("app.yaml"), context.AppConfig);
Console.WriteLine($"{config!.Name} logs at {config.Logging.Level}");

public sealed class AppConfig {
    public string Name { get; set; } = "";
    public LoggingConfig Logging { get; set; } = new();
}

public sealed class LoggingConfig {
    public string Level { get; set; } = "Information";
}

[YamlSerializable(typeof(AppConfig))]
internal sealed partial class AppYamlContext : YamlSerializerContext {
}
```

The context covers the root type and every type it reaches. A shape the serializer cannot handle is a build error (`YAML001` to `YAML005`), not a runtime failure. `AppYamlContext.Default` uses default options.

## Read every document of a multi-document stream

Kubernetes manifests and other streams hold several documents separated by `---`. `Deserialize` on a `YamlReader` reads the next document's root, so a loop reads them all.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

var options = new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase };

using var reader = new YamlReader(File.ReadAllText("manifests.yaml"));
while (YamlSerializer.Deserialize<Manifest>(reader, options) is Manifest manifest) {
    Console.WriteLine($"{manifest.Kind}/{manifest.Metadata.Name}");
}

public sealed class Manifest {
    public string ApiVersion { get; set; } = "";
    public string Kind { get; set; } = "";
    public Metadata Metadata { get; set; } = new();
}

public sealed class Metadata {
    public string Name { get; set; } = "";
}
```

### PowerShell

```powershell
Get-Content ./manifests.yaml -Raw | ConvertFrom-OfficeYaml -AllDocuments -AsHashtable |
    ForEach-Object { '{0}/{1}' -f $_.kind, $_.metadata.name }
```

## Read one section of a loaded document

`Deserialize` also takes any node of a `YamlDocument`. Read a typed section, change it, and write it back with `SerializeToNode`: only that section is rewritten, and comments everywhere else stay.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Yaml;

var options = new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase };

using YamlDocument document = YamlDocument.Load("values.yaml");
Resources resources = YamlSerializer.Deserialize<Resources>(document.Select("resources")!, options) ?? new Resources();
resources.Limits["memory"] = "512Mi";
document.Set("resources", YamlSerializer.SerializeToNode(resources, options));
document.Save();

public sealed class Resources {
    public Dictionary<string, string> Limits { get; set; } = new();
    public Dictionary<string, string> Requests { get; set; } = new();
}
```

## Read files from outside safely

Pass `YamlReaderOptions.UntrustedDefaults` for input you do not control. It lowers the limits to 16 MiB of input, depth 64 and 10,000 alias-expanded nodes, so a "billion laughs" file fails quickly with `YamlLimitExceededException`.

### C#

```csharp
using OfficeIMO.Yaml;

var options = new YamlSerializerOptions { ReaderOptions = YamlReaderOptions.UntrustedDefaults };

try {
    Dictionary<string, object?>? values = YamlSerializer.Deserialize<Dictionary<string, object?>>(File.ReadAllText("upload.yaml"), options);
    Console.WriteLine($"{values?.Count ?? 0} top-level keys");
} catch (YamlException exception) {
    Console.WriteLine($"{exception.Code} at {exception.Line}:{exception.Column}: {exception.Message}");
}
```

### PowerShell

```powershell
Import-OfficeYaml -Path ./upload.yaml -Untrusted -Strict
```
