---
title: Edit and Write TOML Files
description: "Change values, tables and arrays in TOML files by path while comments and formatting stay byte for byte, and write new TOML from .NET objects."
order: 59
verified: 2026-11-14
commands:
  - officeimo convert
packages:
  - OfficeIMO.Toml
  - OfficeIMO.Data.Conversions
cmdlets:
  - Open-OfficeToml
  - Set-OfficeTomlNode
  - Remove-OfficeTomlNode
  - New-OfficeToml
  - Export-OfficeToml
  - ConvertTo-OfficeToml
related:
  - /docs/toml/
  - /docs/toml/read-typed/
  - /docs/yaml/edit-merge/
  - /docs/data-conversions/
limits: TOML has no null; passing null to Add or Set throws, and the serializer leaves null properties out unless NullHandling is Error.
search_terms:
  - update toml
  - preserve comments
  - bump version pyproject
  - serialize toml
  - create toml file
---

`TomlDocument` edits touch only the text they change. Saving an unchanged document reproduces the input byte for byte, including comments, blank lines, quoting, key order, alignment, line endings and spellings such as `0xFF` or `1_000`. For new files, `TomlSerializer` writes objects and `TomlWriter` emits TOML forward-only.

## Change values by path

Paths use dots, `["key"]` for keys that contain dots or brackets, and `[n]` for array items and the tables of an array of tables (`package[3].name`). `Set` adds a missing table header when needed.

### C#

```csharp
using OfficeIMO.Toml;

using TomlDocument document = TomlDocument.Load("pyproject.toml");

string version = document.Select("project.version")?.AsString() ?? "0.0.0";
document.Set("project.version", "1.2.0");                          // only the version text changes
document.Select("project.dependencies")!.AsArray().Add("httpx");   // a new line in the array's own layout
document.Set("tool.ruff.line-length", 100);                        // adds [tool.ruff] if it is missing
document.Remove("tool.ruff.lint.ignore");                          // with the comment lines above it

document.Save();
Console.WriteLine($"{version} -> 1.2.0");
```

`SetValue` keeps the current style when it can hold the new value: hexadecimal stays hexadecimal, a literal string stays literal, and a date-time keeps its `T` or space. An equal value keeps its original spelling.

### PowerShell

```powershell
Open-OfficeToml -Path ./pyproject.toml {
    Set-OfficeTomlNode -Select 'project.version' -Value '1.2.0'
    Set-OfficeTomlNode -Select 'project.dependencies[0]' -Value 'httpx>=0.27'
    Remove-OfficeTomlNode -Select 'tool.ruff.lint.ignore'
}
```

## Add tables, arrays of tables and comments

Every editing method returns the node it creates or changes, so calls chain.

### C#

```csharp
using OfficeIMO.Toml;

using TomlDocument document = TomlDocument.Load("Cargo.toml");
TomlTable root = document.Root;

TomlTable dependencies = document.Select("dependencies")?.AsTable() ?? root.AddTable("dependencies");
TomlTable serde = dependencies.AddInlineTable("serde");               // serde = { ... }
serde.Add("version", "1");
serde.Add("features", new[] { "derive" });
dependencies.Add("anyhow", "1.0").SetComment("error handling");

TomlTable bin = root.AddTableArray("bin").AddTable();                   // a new [[bin]] section
bin.Add("name", "cli");
bin.Add("path", "src/main.rs");

document.Save();
```

`AddTable` and `AddTableArray` place the new `[header]` or `[[header]]` after the table's last section. `LeadingComments` reads the comment lines above a node and `TrailingComment` the comment after it.

### PowerShell

```powershell
Open-OfficeToml -Path ./Cargo.toml {
    Set-OfficeTomlNode -Select 'dependencies.serde' -Value ([ordered]@{ version = '1'; features = @('derive') })
}
```

## Write objects as TOML

Each table's key/value pairs come first and its sub-tables after them, as TOML requires. Collections of objects become arrays of tables, and `[TomlInlineTable]` writes a type or property inline, as Cargo dependencies are written.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

var settings = new ServiceSettings {
    Name = "reporting",
    Server = new ServerSettings { Host = "0.0.0.0", Port = 8443 },
    Mirrors = { new Mirror { Url = "https://a.example.com" }, new Mirror { Url = "https://b.example.com" } }
};

var options = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };
File.WriteAllText("settings.toml", TomlSerializer.Serialize(settings, options));

public sealed class ServiceSettings {
    public string Name { get; set; } = "";
    public ServerSettings Server { get; set; } = new();
    public List<Mirror> Mirrors { get; set; } = new();
}

public sealed class ServerSettings {
    public string Host { get; set; } = "localhost";
    public int Port { get; set; }
}

public sealed class Mirror {
    public string Url { get; set; } = "";
}
```

The output is:

```toml
name = "reporting"

[server]
host = "0.0.0.0"
port = 8443

[[mirrors]]
url = "https://a.example.com"

[[mirrors]]
url = "https://b.example.com"
```

### PowerShell

```powershell
[ordered]@{ name = 'web'; port = 8080; tls = [ordered]@{ enabled = $true } } | ConvertTo-OfficeToml

[ordered]@{ title = 'Report'; owner = [ordered]@{ name = 'Ops'; updated = Get-Date } } |
    Export-OfficeToml -Path ./settings.toml
```

## Build a new TOML file

### C#

```csharp
using OfficeIMO.Toml;

using TomlDocument document = TomlDocument.Create();
document.Set("project.name", "reporting");
document.Set("project.version", "0.1.0");
document.Set("tool.ruff.line-length", 100);
document.Save("pyproject.toml");
```

### PowerShell

```powershell
New-OfficeToml -Path ./pyproject.toml {
    Set-OfficeTomlNode -Select 'project.name' -Value 'reporting'
    Set-OfficeTomlNode -Select 'tool.ruff.line-length' -Value 100
}
```

## Stream TOML with TomlWriter

`TomlWriter` writes headers, keys, values, arrays and inline tables in order and rejects calls TOML does not allow at that point. Text it writes stays readable by TOML 1.0 parsers.

### C#

```csharp
using OfficeIMO.Toml;

using FileStream stream = File.Create("settings.toml");
using var writer = new TomlWriter(stream);

writer.WriteKey("name");
writer.WriteString("officeimo");
writer.WriteTableHeader("tool", "ruff");
writer.WriteKey("line-length");
writer.WriteInteger(120);
```

`TomlWriterOptions.NewLine` selects `\n` (the default) or `\r\n`.

## Convert JSON or YAML to TOML

### C#

```csharp
using System.Text.Json.Nodes;
using OfficeIMO;
using OfficeIMO.Data.Conversions;

OfficeConversionReport report = JsonNode.Parse(File.ReadAllText("service.json")).SaveAsToml("service.toml");
foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine(diagnostic.Message);   // for example, null values TOML cannot hold
}
```

### CLI

```sh
officeimo convert service.json service.toml
officeimo convert values.yaml values.toml --require-no-loss
```
