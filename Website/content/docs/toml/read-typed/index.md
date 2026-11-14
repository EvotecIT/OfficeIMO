---
title: Read TOML into Typed Objects
description: Deserialize TOML files such as Cargo.toml and pyproject.toml into .NET classes and records with OfficeIMO.Toml, with reflection or a source-generated context for NativeAOT.
order: 58
verified: 2026-11-14
packages:
  - OfficeIMO.Toml
cmdlets:
  - Import-OfficeToml
  - ConvertFrom-OfficeToml
related:
  - /docs/toml/
  - /docs/toml/edit-write/
  - /docs/yaml/read-typed/
  - /docs/data-conversions/
limits: TOML has no null, so a missing key reads back as the property's default; a null array item fails with UnrepresentableValue.
search_terms:
  - deserialize toml
  - toml to object
  - cargo.toml
  - pyproject.toml
  - toml config c#
---

`TomlSerializer` maps TOML tables to .NET objects. Its API follows `System.Text.Json` and `YamlSerializer`, and it reads the same `[OfficePropertyName]`, `[OfficeIgnore]`, `[OfficeRequired]` and `[OfficeConstructor]` attributes, so one model maps to both TOML and YAML. Use it when you want typed settings; load a `TomlDocument` instead when you will change the file and keep its comments.

## Read a TOML file into a class

TOML keys are often kebab-case, so set the naming policy to match. Tables become objects or dictionaries, arrays and arrays of tables become lists, and unknown keys are skipped.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

var options = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };

CargoManifest manifest = TomlSerializer.Deserialize<CargoManifest>(File.ReadAllText("Cargo.toml"), options);
Console.WriteLine($"{manifest.Package.Name} {manifest.Package.Version}");
foreach (BinaryTarget bin in manifest.Bin) Console.WriteLine($"binary {bin.Name} at {bin.Path}");

public sealed class CargoManifest {
    public PackageInfo Package { get; set; } = new();
    public Dictionary<string, object?> Dependencies { get; set; } = new();
    public List<BinaryTarget> Bin { get; set; } = new();
}

public sealed class PackageInfo {
    public string Name { get; set; } = "";
    public string Version { get; set; } = "";
    public string? RustVersion { get; set; }
}

public sealed class BinaryTarget {
    public string Name { get; set; } = "";
    public string Path { get; set; } = "";
}
```

Values that do not fit their type fail with `TomlException` code `TypeMismatch`, the line and column, and the path, such as `bin[1].path`. `DateTimeOffset`, `DateTime`, `DateOnly`, `TimeOnly` and `TomlDateTime` read TOML's four date-time kinds; `decimal` reads a float's exact digits without going through a `double`.

### PowerShell

```powershell
$ruff = (Import-OfficeToml -Path ./pyproject.toml).tool.ruff
$ruff.'line-length'

(Get-Content ./Cargo.lock -Raw | ConvertFrom-OfficeToml -AsHashtable).package |
    ForEach-Object { '{0} {1}' -f $_.name, $_.version }
```

## Use a source-generated context

In `PublishAot` and `PublishTrimmed` applications the reflection overloads are off. Declare a context; the source generator ships in the `OfficeIMO.Toml` package.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

var context = new SettingsTomlContext(new TomlSerializerOptions {
    PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower
});

ServiceSettings settings = TomlSerializer.Deserialize(File.ReadAllText("settings.toml"), context.ServiceSettings);
Console.WriteLine($"{settings.Name} listens on {settings.Server.Port}");

public sealed class ServiceSettings {
    public string Name { get; set; } = "";
    public ServerSettings Server { get; set; } = new();
}

public sealed class ServerSettings {
    public string Host { get; set; } = "localhost";
    public int Port { get; set; } = 8080;
    public TimeSpan RequestTimeout { get; set; } = TimeSpan.FromSeconds(30);
}

[TomlSerializable(typeof(ServiceSettings))]
internal sealed partial class SettingsTomlContext : TomlSerializerContext {
}
```

A shape the serializer cannot handle is a build error (`TOML001` to `TOML005`).

## Read one table of a loaded document

`Deserialize` takes any node of a `TomlDocument`, which suits files where only one table is yours, such as `[tool.ruff]` in `pyproject.toml`.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

var options = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };

using TomlDocument document = TomlDocument.Load("pyproject.toml");
if (document.Select("tool.ruff") is TomlNode node) {
    RuffSettings ruff = TomlSerializer.Deserialize<RuffSettings>(node, options);
    Console.WriteLine($"line length {ruff.LineLength}, target {ruff.TargetVersion}");
}

public sealed class RuffSettings {
    public int LineLength { get; set; } = 88;
    public string? TargetVersion { get; set; }
    public List<string> Exclude { get; set; } = new();
}
```

## Read polymorphic tables

TOML has no tags, so a discriminator key selects the derived type. By default it is `"$type"`; `[OfficePolymorphic]` names another key. An unknown discriminator fails with `UnknownDiscriminator`.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

const string toml = """
    [[probe]]
    kind = "http"
    path = "/health"

    [[probe]]
    kind = "tcp"
    port = 5432
    """;

HealthChecks checks = TomlSerializer.Deserialize<HealthChecks>(toml);
foreach (Probe probe in checks.Probe) Console.WriteLine(probe.GetType().Name);

public sealed class HealthChecks {
    public List<Probe> Probe { get; set; } = new();
}

[OfficePolymorphic(DiscriminatorPropertyName = "kind")]
[OfficeDerivedType(typeof(HttpProbe), "http")]
[OfficeDerivedType(typeof(TcpProbe), "tcp")]
public abstract class Probe { }

public sealed class HttpProbe : Probe {
    public string Path { get; set; } = "/";
}

public sealed class TcpProbe : Probe {
    public int Port { get; set; }
}
```
