# OfficeIMO.Toml

OfficeIMO's own TOML 1.1 engine. It reads, edits and writes TOML without changing anything you didn't change, has no runtime dependencies beyond `OfficeIMO.Core`, and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

The package provides `TomlDocument`, a lossless editable model, `TomlReader`, a forward-only pull reader, and `TomlWriter`, a forward-only emitter.

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

`TomlDocument.Load` reads a file or stream, `LoadAsync` does the same asynchronously, and `Parse` takes a string. `Create()` starts an empty document. `Save()` writes to the file the document was loaded from or created for; `Save(path)`, `Save(stream)`, `SaveAsync` and `ToString()` are also available. Saving an unchanged document reproduces the input byte for byte, including comments, blank lines, quoting, key order, alignment, line endings, a UTF-8 byte-order mark and the original spelling of every value (`0xFF`, `1_000`, `1979-05-27 07:32:00Z`).

`Root` is the root table. Nodes are:

- `TomlTable`: the root, a `[table]`, an inline table, a table made by dotted keys, or one table of an array of tables (`TableKind` says which). Keys are unique and keep source order.
- `TomlTableArray`: the tables of the `[[name]]` headers with one name, such as the `[[package]]` entries of a lock file.
- `TomlArray`: an array value in brackets.
- `TomlValue`: a string, integer, float, boolean or date-time. `Kind` tells which; `Value` is the decoded `string`, `long`, `double`, `bool` or `TomlDateTime`.

A table is one object even when the file spreads it out: `[a]`, `[a.b]` and `a.c = 1` lines all contribute to the same tree.

- **Typed reads**: `GetValue<T>()` and `TryGetValue<T>()` for `string`, `bool`, `int`, `long`, `double`, `float`, `decimal`, `TomlDateTime`, `DateTime`, `DateTimeOffset` and `TimeSpan`, plus `DateOnly` and `TimeOnly` on .NET 8 and later. `AsString()` returns a string's text, or the TOML spelling of any other value. `AsTable()`, `AsArray()`, `AsTableArray()` and `AsValue()` throw an exception naming the actual kind.
- **Paths**: `Select("tool.ruff.line-length")` returns a node or `null`. Use `["key"]` with JSON escapes for keys containing dots or brackets, and `[n]` for array items and the tables of an array of tables (`package[3].name`). `Set(path, value)` and `Remove(path)` edit by path. The grammar matches the paths `OfficeIMO.Reader.Toml` reports.
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
| `TomlArray.Add` / `Insert` / `Set` / `RemoveAt` | In the array's own layout: a new line with the same indentation and comma style in a multi-line array, `, item` on one line otherwise. |
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

## Safety and errors

Invalid input throws `TomlException` with a stable `Code` from `TomlErrorCodes`, plus `Line`, `Column` and `Position`. `TomlLoadOptions` bounds nesting depth (256), node count (1,000,000), string length (16 MiB) and, when set, input size; exceeding a limit throws `TomlLimitExceededException` naming the limit. Use `TomlLoadOptions.UntrustedDefaults` for files from outside: 16 MiB of input, depth 64, 200,000 nodes and 1 MiB per string. Loading checks cancellation at bounded intervals.

## Conformance

The engine passes all 214 valid and 467 invalid TOML 1.1 cases of [toml-test](https://github.com/toml-lang/toml-test) 2.2.0, and every valid case and real-world corpus file (`pyproject.toml`, `Cargo.toml`, `Cargo.lock`, `uv.lock`, Hugo, Netlify, Starship) saves back byte for byte.
