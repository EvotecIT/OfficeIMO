---
title: "TOML Files"
description: "Read, edit and write TOML 1.1 in .NET with a lossless document model that saves unchanged input byte for byte, a pull reader and a forward-only writer."
order: 57
meta.seo_title: "Read and edit TOML in C# with OfficeIMO.Toml"
---

`OfficeIMO.Toml` is a TOML 1.1 engine with four entry points: `TomlDocument`, a lossless editable model; `TomlReader`, a forward-only pull reader; `TomlWriter`, a forward-only emitter; and `TomlSerializer`, which maps .NET objects to TOML. It depends only on `OfficeIMO.Core` and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

```shell
dotnet add package OfficeIMO.Toml
```

## Edit a file without reformatting it

```csharp
using OfficeIMO.Toml;

using var document = TomlDocument.Load("pyproject.toml");
string name = document.Select("project.name")!.AsString();
int lineLength = document.Select("tool.ruff.line-length")?.GetValue<int>() ?? 88;

document.Set("project.version", "1.2.0");                        // only the version text changes
document.Select("project.dependencies")!.AsArray().Add("httpx"); // follows the array's own layout
document.Set("tool.ruff.line-length", 100);                      // adds [tool.ruff] if it is missing
document.Save();
```

Saving an unchanged document reproduces the input byte for byte: comments, blank lines, quoting, key order, alignment, line endings, a UTF-8 byte-order mark and the original spelling of values such as `0xFF`, `1_000` or `1979-05-27 07:32:00Z`. Edits touch only the text they change:

- `SetValue` keeps the current style when it can hold the new value (hexadecimal stays hexadecimal, literal strings stay literal).
- `Add` and `Set` insert a `key = value` line after the table's last pair, using its indentation and `=` spacing.
- Array edits follow the array's layout: a new line in a multi-line array, `, item` on a single line.
- `Remove` deletes a key with the comment lines directly above it.
- `LeadingComments`, `TrailingComment` and `SetComment` read and change comments.

Paths use dots, `["key"]` for keys containing dots or brackets, and `[n]` for array items and array-of-tables entries (`package[3].name`). Typed reads cover `string`, `bool`, `int`, `long`, `double`, `float`, `decimal`, `TomlDateTime`, `DateTime`, `DateTimeOffset` and `TimeSpan`, plus `DateOnly` and `TimeOnly` on .NET 8 and later. `TomlDateTime` keeps all four TOML date-time kinds with nanosecond precision.

## Read tokens

```csharp
var reader = new TomlReader(File.ReadAllText("Cargo.toml"));
while (reader.Read()) {
    if (reader.TokenType == TomlTokenType.Key) Console.Write(reader.GetString() + " ");
    else if (reader.TokenType == TomlTokenType.String) Console.WriteLine("= " + reader.GetString());
}
```

The reader allocates nothing per token; strings are created only when you call `GetString()`, and `ValueSpan` exposes the source text of the current token. It validates TOML 1.1 syntax completely. Whole-document rules such as duplicate keys and table redefinition are checked by `TomlDocument`.

## Write

```csharp
using var stream = File.Create("settings.toml");
using var writer = new TomlWriter(stream);
writer.WriteKey("name");
writer.WriteString("officeimo");
writer.WriteTableHeader("tool", "ruff");
writer.WriteKey("line-length");
writer.WriteInteger(120);
```

`TomlWriter` rejects calls TOML does not allow at that point and produces deterministic output. `TomlWriterOptions.NewLine` selects `\n` (default) or `\r\n`. Text written by OfficeIMO stays readable by TOML 1.0 parsers: no `\e` or `\x` escapes, and times always include seconds.

## Serializing objects

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

[TomlSerializable(typeof(CargoManifest))]
internal partial class CargoContext : TomlSerializerContext { }

var context = new CargoContext(new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower });
CargoManifest manifest = TomlSerializer.Deserialize(File.ReadAllText("Cargo.toml"), context.CargoManifest);
string toml = TomlSerializer.Serialize(manifest, context.CargoManifest);
```

`TomlSerializer` reads the same `[OfficePropertyName]`, `[OfficeIgnore]`, `[OfficeRequired]`, `[OfficeConstructor]` and `[OfficeDerivedType]` attributes as `YamlSerializer`, so one model maps to both formats. Objects and dictionaries become `[table]` sections after the key/value pairs, collections of objects become `[[table]]` arrays of tables, and `[TomlInlineTable]` writes a type or property inline, as Cargo dependencies are written. TOML has no null: null values are left out, or fail when `NullHandling` is `Error`. The source generator ships in the package; overloads that take `TomlSerializerOptions` alone use reflection metadata, which is off in `PublishAot` and `PublishTrimmed` applications.

## Untrusted input

```csharp
try {
    using var config = TomlDocument.Load("upload.toml", TomlLoadOptions.UntrustedDefaults);
}
catch (TomlException exception) {
    Console.WriteLine($"{exception.Code} at {exception.Line}:{exception.Column}");
}
```

`TomlLoadOptions` bounds nesting depth (256), node count (1,000,000) and string length (16 MiB) by default. `UntrustedDefaults` tightens these to 16 MiB of input, depth 64, 200,000 nodes and 1 MiB per string. Exceeding a limit throws `TomlLimitExceededException` naming the limit.

## Conformance and limits

The engine passes all 214 valid and 467 invalid TOML 1.1 cases of toml-test 2.2.0, and valid cases and real-world files (`pyproject.toml`, `Cargo.toml`, `Cargo.lock`, `uv.lock`, Hugo, Netlify, Starship) save back byte for byte.

- TOML has no null; passing `null` to `Add` or `Set` throws.
- The engine reads and writes TOML only. To extract TOML into chunks for search or indexing, use `OfficeIMO.Reader.Toml`, which emits one path, type and value row per value with the same path grammar as `Select`.

To convert TOML to and from JSON, YAML, XML, CSV or Excel, see [data conversions](/docs/data-conversions/).

See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Toml/README.md) and the [Reader overview](/docs/reader/).
