# OfficeIMO.Toml

OfficeIMO's own TOML 1.1 engine. It reads, edits and writes TOML without changing anything you didn't change, has no runtime dependencies beyond `OfficeIMO.Core`, and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

The package provides `TomlDocument`, a lossless editable model, `TomlReader`, a forward-only pull reader, `TomlWriter`, a forward-only emitter, and `TomlSerializer`, which maps .NET objects to TOML with a NativeAOT source generator.

## Loading and editing a document

```csharp
using OfficeIMO.Toml;

using var document = TomlDocument.Load("pyproject.toml");
string name = document.Select("project.name")!.AsString();
int lineLength = document.Select("tool.ruff.line-length")?.GetValue<int>() ?? 88;

document.Set("project.version", "1.2.0");                        // only the version text changes
document.Select("project.dependencies")!.AsArray().Add("httpx"); // a new line in the array's own layout
document.Set("tool.ruff.line-length", 100);                      // adds [tool.ruff] if it is missing
document.Save();                                                 // back to pyproject.toml
```

`TomlDocument.Load` reads a file or stream, `LoadAsync` does the same asynchronously, and `Parse` takes a string. `Create()` starts an empty document. `Save()` writes to the document's `FilePath`, the file it was loaded from, created for or last saved to; `Save(path)` replaces a file atomically and makes it the `FilePath`. `Save(stream)`, `SaveAsync` and `ToString()` are also available. Saving an unchanged document reproduces the input byte for byte, including comments, blank lines, quoting, key order, alignment, line endings, a UTF-8 byte-order mark and the original spelling of every value (`0xFF`, `1_000`, `1979-05-27 07:32:00Z`).

`Root` is the root table. Nodes are:

- `TomlTable`: the root, a `[table]`, an inline table, a table made by dotted keys, or one table of an array of tables (`TableKind` says which). Keys are unique and keep source order.
- `TomlTableArray`: the tables of the `[[name]]` headers with one name, such as the `[[package]]` entries of a lock file.
- `TomlArray`: an array value in brackets.
- `TomlValue`: a string, integer, float, boolean or date-time. `Kind` tells which; `Value` is the decoded `string`, `long`, `double`, `bool` or `TomlDateTime`.

A table is one object even when the file spreads it out: `[a]`, `[a.b]` and `a.c = 1` lines all contribute to the same tree.

- **Typed reads**: `GetValue<T>()` and `TryGetValue<T>()` for `string`, `bool`, `int`, `long`, `double`, `float`, `decimal`, `TomlDateTime`, `DateTime`, `DateTimeOffset` and `TimeSpan`, plus `DateOnly` and `TimeOnly` on .NET 8 and later. `AsString()` returns a string's text, or the TOML spelling of any other value. `AsTable()`, `AsArray()`, `AsTableArray()` and `AsValue()` throw an exception naming the actual kind.
- **Paths**: `Select("tool.ruff.line-length")` returns a node or `null`. Use `["key"]` with JSON escapes for keys containing dots or brackets, and `[n]` for array items and the tables of an array of tables (`package[3].name`). `Set(path, value)` and `Remove(path)` edit by path. The grammar matches the paths `OfficeIMO.Reader.Toml` reports.
- **Positions**: `Start` and `End` locate a loaded node in the text, and `KeyStart` and `KeyEnd` the key segment that names it (for a table made by dotted keys or deeper headers, its segment of the first such key). Nodes added by editing have the default position, whose `Line` is 0.
- **Comments**: `LeadingComments` returns the comment lines directly above a node, and `TrailingComment` the comment after it. `SetComment(text)` sets, replaces or removes the comment after a key/value pair or table header.
- **Dates and times**: `TomlDateTime` keeps all four TOML kinds (offset date-time, local date-time, local date, local time) with nanosecond precision, and converts to `DateTimeOffset`, `DateTime` and `TimeSpan`.

### Editing

Every editing method returns the node it creates or changes, so calls chain: `table.Add("version", "1.0").SetComment("pinned")`.

| Method | Effect on save |
|---|---|
| `TomlValue.SetValue(value)` | Only the value's text changes. An equal value keeps the original spelling; otherwise the current style is kept when it can hold the new value: hexadecimal stays hexadecimal, a literal string stays literal, a date-time keeps its `T` or space. |
| `TomlTable.Add(key, value)` / `Set(key, value)` | A `key = value` line after the table's last pair, with its indentation and `=` spacing. Inside an inline table, an entry inside the braces. New strings follow the quoting of their siblings. |
| `TomlTable.AddTable(key)` / `AddTableArray(key)` | A `[header]` or `[[header]]` section after the table's last section, separated by a blank line when the file does that. |
| `TomlTable.AddArray(key)` / `AddInlineTable(key)` | `key = []` or `key = {}` to fill. |
| `TomlArray.Add` / `Insert` / `Set` / `RemoveAt` | In the array's own layout (`Set` at the array's count appends, as does a path `Set` such as `deps[3]` on a three-item array): a new line with the same indentation and comma style in a multi-line array, `, item` on one line otherwise. |
| `TomlTable.Remove(key)` / `TomlTableArray.RemoveAt(i)` | The key/value line with the comment lines directly above it, or a table's sections, wherever in the file they are. |
| `SetComment(text)` | Only the comment after the node changes. |

Values passed to `Add` and `Set` can be strings, booleans, integers, floating-point numbers, `TomlDateTime`, `DateTime`, `DateTimeOffset`, `DateOnly`, `TimeOnly`, sequences (arrays) and dictionaries with string keys (inline tables). TOML has no null; passing `null` throws.

New text uses the document's line endings. Text written by OfficeIMO stays readable by TOML 1.0 parsers: strings never use the 1.1-only `\e` and `\x` escapes, and times always have seconds.

## Reading tokens

```csharp
using OfficeIMO.Toml;

var reader = new TomlReader(File.ReadAllText("Cargo.toml"));
while (reader.Read()) {
    if (reader.TokenType == TomlTokenType.Key) Console.Write(reader.GetString() + " ");
    else if (reader.TokenType == TomlTokenType.String) Console.WriteLine("= " + reader.GetString());
}
```

`Read()` moves through header brackets, key segments, values, array and inline-table brackets and, with `TomlReaderOptions.IncludeComments`, comments. A dotted key `a.b.c = 1` reads as three `Key` tokens followed by the value. The reader allocates nothing per token; strings are created only when you call `GetString()`, and `ValueSpan` exposes the source text of the current token without copying. `GetInt64`, `GetDouble`, `GetBoolean` and `GetDateTime` decode values.

The reader checks TOML 1.1 syntax completely. Rules that need the whole document, such as duplicate keys and table redefinition, are checked by `TomlDocument`.

The reader also reads a `Stream` (UTF-8, with or without a byte-order mark) or a `TextReader` without loading the whole input, which suits very large lock files:

```csharp
using var stream = File.OpenRead("Cargo.lock");
using var reader = new TomlReader(stream);
int packages = 0;
while (reader.Read()) {
    if (reader.TokenType == TomlTokenType.ArrayOfTablesHeaderStart) packages++;
}
```

It keeps whole lines in a pooled buffer that grows only while one line or multi-line string is longer than it. Tokens, positions and errors are the same as for the same text in memory, and invalid UTF-8 fails with `TomlErrorCodes.InvalidEncoding` at the first bad byte. `ValueSpan` is valid until the next `Read()`. Dispose the reader to return its buffers; it never disposes your stream or text reader. `TomlReaderOptions.MaxInputBytes` bounds the input (bytes for a stream, characters otherwise). `TomlDocument.Load` keeps the whole text, because saving reproduces it.

## Writing

```csharp
using var writer = new TomlWriter(stream);
writer.WriteKey("name");
writer.WriteString("officeimo");
writer.WriteTableHeader("tool", "ruff");
writer.WriteKey("line-length");
writer.WriteInteger(120);
```

`TomlWriter` writes headers, keys (dotted with several arguments), values, arrays and inline tables in order and rejects calls TOML does not allow at that point. Output is deterministic. Comments written just before a header stay directly above it, with the blank line that separates sections above the comments. `TomlWriterOptions.NewLine` selects `\n` (the default) or `\r\n`.

## Serializing objects

`TomlSerializer` converts .NET objects to and from TOML documents. Its API follows `System.Text.Json` and `YamlSerializer`, and it reads the same mapping attributes, naming policy and ignore condition, which live in the `OfficeIMO` namespace, so one model maps to YAML and TOML alike:

```csharp
using OfficeIMO;
using OfficeIMO.Toml;

var options = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };
CargoManifest manifest = TomlSerializer.Deserialize<CargoManifest>(File.ReadAllText("Cargo.toml"), options);
string toml = TomlSerializer.Serialize(manifest, options);

using var document = TomlDocument.Load("pyproject.toml");
RuffSettings ruff = TomlSerializer.Deserialize<RuffSettings>(document.Select("tool.ruff")!, options);
```

These overloads create metadata by reflection. For trimmed and NativeAOT applications, declare a context and pass its generated metadata. The source generator ships in this package:

```csharp
[TomlSerializable(typeof(CargoManifest))]
internal partial class CargoContext : TomlSerializerContext { }

var context = new CargoContext(new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower });
CargoManifest manifest = TomlSerializer.Deserialize(text, context.CargoManifest);
string toml = TomlSerializer.Serialize(manifest, context.CargoManifest);
```

A shape the serializer cannot handle is a build error (`TOML001` to `TOML005`). In a `PublishAot` or `PublishTrimmed` application the reflection overloads are off; set the `TomlSerializerIsReflectionEnabledByDefault` MSBuild property to change this. Every method takes a string, a UTF-8 `Stream`, or a `TextReader`/`TextWriter`; `Deserialize` also takes any `TomlNode` of a loaded document.

### Layout

A TOML document is a table, so the root is an object or a dictionary. Writing puts each table's key/value pairs first and its tables after them, as TOML requires:

```toml
version = "1.2.0"

[package]
name = "demo"

[dependencies]
serde = { version = "1", features = ["derive"] }

[[bin]]
name = "cli"
path = "src/main.rs"

[profile.release]
lto = true
```

- Objects and dictionaries are `[table]` sections. A table with only sub-tables gets no header of its own (`[profile.release]` without `[profile]`).
- A non-empty collection of objects is an array of tables (`[[bin]]`); an empty one is `bin = []`.
- Inside an array or inline table everything is inline: `owner = { name = "Ada", address = { city = "London" } }`.
- `[TomlInlineTable]` on a class or struct writes its values as inline tables, as Cargo dependencies are written. On a property it writes that property's object or collection of objects inline, or, for a dictionary, keeps the dictionary's section and writes its entries inline.

Reading accepts every form for every type: an array of tables or an array of inline tables reads into the same list.

### Supported types

- **Objects:** classes, structs and records with public properties, as `YamlSerializer` maps them: `init` properties, the parameterless, only or `[OfficeConstructor]` constructor with parameters bound by name, C# `required` and `[OfficeRequired]` members, `[OfficePropertyName]` and `[OfficeIgnore]`. Unknown keys are skipped.
- **Collections:** arrays, `List<T>`, the list and set interfaces, `HashSet<T>`, and other collections with a parameterless constructor and `Add`.
- **Dictionaries:** tables. Keys are strings, or numbers, enums and other scalars written as their text (`[ports]` with `8080 = "web"`).
- **Scalars:**

  | Type | TOML |
  | --- | --- |
  | `string`, `char`, `Guid`, `Uri` | string |
  | `bool` | boolean |
  | integer types, `BigInteger` | integer; a whole float in range also reads. Values outside the 64-bit signed range fail with `UnrepresentableValue` |
  | `double`, `float` | float, including `inf`, `-inf` and `nan`; an integer also reads |
  | `decimal` | float with its exact digits (`0.1`), read from the text without going through a double |
  | `DateTimeOffset` | offset date-time `1979-05-27T07:32:00-07:00` |
  | `DateTime` | local date-time `1979-05-27T07:32:00`; a UTC or local kind is an offset date-time. An offset date-time reads as its UTC instant, a local date as midnight |
  | `DateOnly`, `TimeOnly` | local date `1979-05-27`, local time `07:32:00` |
  | `TimeSpan` | local time when it is a time of day, otherwise the string `"1.02:03:04"`; both read |
  | `TomlDateTime` | any of the four kinds, unchanged |
  | `byte[]` | Base64 string |
  | enums | member names; `[Flags]` combinations as `"Read, Write"`, and an array of names also reads |

- **Nullable value types** read and write the value. **TOML has no null:** a null property or dictionary value is left out, so it reads back as the default. Set `TomlSerializerOptions.NullHandling` to `Error` to fail instead. A null array item always fails with `UnrepresentableValue`, because leaving it out would shift the items after it.
- **`object`:** tables read as `Dictionary<string, object?>`, arrays and arrays of tables as `List<object?>`, and values as `string`, `long`, `double`, `bool` or `TomlDateTime`, which keeps the date-time kind.
- **Polymorphism** through `[OfficeDerivedType]` registrations on the base type. TOML has no tags, so the discriminator is the first key, `"$type" = "http"`, or the key `[OfficePolymorphic(DiscriminatorPropertyName = "kind")]` names. Reading finds it anywhere in the table; an unknown discriminator fails with `UnknownDiscriminator`.

Reading applies every TOML rule and the limits in `TomlSerializerOptions.LoadOptions`, as `TomlDocument` does. A value that does not fit its type fails with `TomlException` code `TypeMismatch`, its line and column, and its path, such as `tool.ruff.lint.select[1]`. A missing required key fails with `MissingRequiredMember`; an object that contains itself fails with `ReferenceCycle`.

## Safety and errors

Invalid input throws `TomlException` with a stable `Code` from `TomlErrorCodes`, plus `Line`, `Column` and `Position`; `ToDiagnostic()` returns the failure as an error `OfficeDiagnostic` (code `toml-load-duplicate-key` for `Toml.Load.DuplicateKey`, location `line:column`) to report next to other documents' findings. `TomlLoadOptions` bounds nesting depth (256), node count (1,000,000), string length (16 MiB) and, when set, input size; exceeding a limit throws `TomlLimitExceededException` naming the limit. Use `TomlLoadOptions.UntrustedDefaults` for files from outside: 16 MiB of input, depth 64, 200,000 nodes and 1 MiB per string. Loading checks cancellation at bounded intervals.

## Conformance

The engine passes all 214 valid and 467 invalid TOML 1.1 cases of [toml-test](https://github.com/toml-lang/toml-test) 2.2.0, and every valid case and real-world corpus file (`pyproject.toml`, `Cargo.toml`, `Cargo.lock`, `uv.lock`, Hugo, Netlify, Starship) saves back byte for byte. An opt-in run over about 5,400 pinned real-world files (crates.io and workspace manifests, lock files, `pyproject.toml`, Hugo sites) also checks stream reading, scripted edits and hostile mutations; the [TOML engine design](../Docs/officeimo.toml-engine-design.md#qualification) has the results and the benchmark comparison with Tomlyn.
