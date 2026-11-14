---
title: "TOML Files"
description: "Read, edit and write TOML 1.1 files in .NET with a lossless document model, a streaming pull reader and an object serializer from OfficeIMO.Toml."
order: 57
verified: 2026-11-14
commands:
  - officeimo read
meta.seo_title: "Read and edit TOML in C# with OfficeIMO.Toml"
packages:
  - OfficeIMO.Toml
cmdlets:
  - Open-OfficeToml
  - Get-OfficeTomlNode
  - Import-OfficeToml
related:
  - /docs/toml/read-typed/
  - /docs/toml/edit-write/
  - /docs/data-conversions/
  - /docs/yaml/
limits: TomlLoadOptions bound depth (256), nodes (1,000,000) and string length (16 MiB) by default; UntrustedDefaults lowers them to depth 64, 200,000 nodes, 1 MiB strings and 16 MiB of input.
search_terms:
  - toml parser
  - read toml c#
  - cargo.toml
  - pyproject.toml
  - cargo.lock
---

`OfficeIMO.Toml` is a TOML 1.1 engine with no runtime dependencies beyond `OfficeIMO.Core`. It targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

```sh
dotnet add package OfficeIMO.Toml
```

| Type | Use it to |
|---|---|
| `TomlDocument` | Load, edit and save a file losslessly. |
| `TomlSerializer` | Convert .NET objects to and from TOML, with reflection or a source-generated `TomlSerializerContext`. |
| `TomlReader` | Read tokens forward-only from a string, stream or `TextReader`. |
| `TomlWriter` | Emit TOML forward-only with deterministic output. |

The engine passes all 214 valid and 467 invalid TOML 1.1 cases of toml-test 2.2.0, and real-world files (`pyproject.toml`, `Cargo.toml`, `Cargo.lock`, `uv.lock`, Hugo, Netlify, Starship) save back byte for byte.

## Load a TOML file and read values

`Load` reads a file or stream, `LoadAsync` does the same asynchronously and `Parse` takes a string. `Root` is the root table; a table is one object even when the file spreads it over `[a]`, `[a.b]` and dotted keys.

### C#

```csharp
using OfficeIMO.Toml;

using TomlDocument document = TomlDocument.Load("Cargo.toml");

string name = document.Select("package.name")!.AsString();
string edition = document.Select("package.edition")?.AsString() ?? "2015";
Console.WriteLine($"{name} (edition {edition})");

if (document.Select("dependencies") is TomlNode node) {
    foreach (KeyValuePair<string, TomlNode> dependency in node.AsTable()) {
        Console.WriteLine($"  {dependency.Key}");
    }
}
```

Typed reads cover `string`, `bool`, `int`, `long`, `double`, `float`, `decimal`, `TomlDateTime`, `DateTime`, `DateTimeOffset` and `TimeSpan`, plus `DateOnly` and `TimeOnly` on .NET 8 and later. `TomlDateTime` keeps all four TOML date-time kinds with nanosecond precision.

### PowerShell

```powershell
$document = Open-OfficeToml -Path ./Cargo.toml
Get-OfficeTomlNode -Document $document -Select 'package.name', 'package.version' |
    Select-Object -Property Path, Value, Line, Column
Close-OfficeToml -Document $document
```

### CLI

```sh
officeimo read Cargo.toml --format json
```

## Stream a large lock file

`TomlReader` reads a stream without loading the whole input and allocates nothing per token; strings are created only when you call `GetString()`. It checks TOML 1.1 syntax completely. Whole-document rules such as duplicate keys are checked by `TomlDocument`.

### C#

```csharp
using OfficeIMO.Toml;

using FileStream stream = File.OpenRead("Cargo.lock");
using var reader = new TomlReader(stream);

int packages = 0;
while (reader.Read()) {
    if (reader.TokenType == TomlTokenType.ArrayOfTablesHeaderStart) packages++;
}
Console.WriteLine($"{packages} locked packages");
```

## Load TOML from outside safely

Invalid input throws `TomlException` with a stable `Code`, `Line` and `Column`. Exceeding a limit throws `TomlLimitExceededException` naming the limit.

### C#

```csharp
using OfficeIMO.Toml;

try {
    using TomlDocument config = TomlDocument.Load("upload.toml", TomlLoadOptions.UntrustedDefaults);
    Console.WriteLine($"{config.Root.Count} top-level keys");
} catch (TomlException exception) {
    Console.WriteLine($"{exception.Code} at {exception.Line}:{exception.Column}");
}
```

### PowerShell

```powershell
$document = Open-OfficeToml -LiteralPath ./uploads/Cargo.toml -Untrusted
Get-OfficeTomlNode -Document $document -Select 'package.name'
Close-OfficeToml -Document $document
```

To extract TOML into chunks for search or indexing, use `OfficeIMO.Reader.Toml`, which emits one path, type and value row per value with the same path grammar as `Select`. The [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Toml/README.md) lists every node type and editing method.
