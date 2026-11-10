# OfficeIMO.Yaml

OfficeIMO's own YAML 1.2 engine. It has no runtime dependencies beyond `OfficeIMO.Core` and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

The package provides `YamlDocument`, a lossless document model, and `YamlReader`, a forward-only pull reader over YAML 1.2 text.

## Loading a document

```csharp
using OfficeIMO.Yaml;

YamlDocument document = YamlDocument.Load("deployment.yaml");
YamlMapping spec = document.Root!.AsMapping()["spec"].AsMapping();
int replicas = spec["replicas"].GetValue<int>();
string? image = document.Select("spec.template.spec.containers[0].image")?.AsString();
document.Save("deployment.yaml");   // unchanged: the same bytes as the input
```

`YamlDocument.Load` reads a file or stream, `LoadAsync` does the same asynchronously, and `Parse` takes a string. The encoding (UTF-8, UTF-16 or UTF-32, either byte order) and byte-order mark are detected and kept, so `Save`, `SaveAsync` and `ToString()` reproduce an unchanged input byte for byte, including comments, blank lines, quoting, key order, line endings and the original spelling of every scalar.

`Documents` lists every document of a multi-document stream as a `YamlDocumentPart` with its root, `%YAML` version, `%TAG` directives and `---`/`...` markers; `Root` is the first document's root. Nodes are `YamlMapping` (ordered entries; string-key lookup), `YamlSequence`, `YamlScalar` and `YamlAlias`. They store offsets into the retained source rather than copies, and scalar values are decoded on first use.

- **Typed reads** go through the document's schema: `GetValue<T>()` and `TryGetValue<T>()` for `string`, `bool`, `int`, `long`, `double`, `float`, `decimal` and their nullable forms; `AsString()` for the decoded text; `IsNull`. `AsMapping()`, `AsSequence()` and `AsScalar()` throw an exception naming the actual node kind. Reading never changes the source text.
- **Aliases** stay `YamlAlias` nodes in the tree. Typed reads and `Target` look through them to the anchored node.
- **Merge keys** (`<<`) apply to key lookups and paths under the YAML 1.1 schema, without being flattened into the tree.
- **Paths**: `Select("spec.containers[0].image")` returns a node or `null`. Use `["key"]` with JSON escapes for keys containing dots or brackets, `[n]` for sequence items, and `$[n]` on a multi-document `YamlDocument` to pick a document. The grammar matches the paths `OfficeIMO.Reader.Yaml` reports.
- **Comments**: `LeadingComments` returns the comments on their own lines before a node, and `TrailingComment` the comment on the line where it ends. A comment belongs to the innermost node next to it; the XML documentation gives the exact rule.
- **Duplicate keys** fail the load by default, as the specification requires. `YamlLoadOptions.DuplicateKeys` set to `FirstWins` or `LastWins` loads such files, keeps every entry and reports each repeat in `Diagnostics`.
- **Custom tags** such as `!Ref` are kept verbatim in `Tag` and never fail a load.

## Reading events

```csharp
using OfficeIMO.Yaml;

using var reader = new YamlReader("""
    name: web
    replicas: 3
    ports: [80, 443]
    """);

while (reader.Read()) {
    if (reader.TokenType == YamlTokenType.Scalar) {
        if (reader.TryGetInt64(out long number)) Console.WriteLine($"{reader.Start}: int {number}");
        else Console.WriteLine($"{reader.Start}: {reader.GetString()}");
    }
}
```

`Read()` moves through stream, document, mapping, sequence, scalar and alias events. The reader allocates nothing per event; strings are created only when you call `GetString()`, `Tag` or `Anchor`. `ValueSpan` exposes the raw source text of the current value without copying.

Readers accept a `string`, `ReadOnlyMemory<char>`, `TextReader` or `Stream`. Stream input detects UTF-8, UTF-16 and UTF-32 (either byte order) from the byte-order mark or the first bytes, as the specification requires. Stream and `TextReader` input is read through a pooled sliding buffer. The reader never disposes the stream or text reader you pass in.

## Typed values and schemas

`TryGetBoolean`, `TryGetInt64`, `TryGetDouble`, `TryGetDecimal` and `IsNull` resolve untagged plain scalars through the document's schema. Quoted and block scalars are strings unless they carry an explicit tag such as `!!int`.

| Schema | Resolves |
|---|---|
| `Core` (default) | `null`, `~`, empty; `true`/`false` in three casings; decimal, `0o` and `0x` integers; floats, `.inf`, `.nan` |
| `Json` | `null`, `true`, `false` and JSON numbers only |
| `Failsafe` | nothing; every scalar is a string |
| `Yaml11` | `yes`/`no`/`on`/`off`/`y`/`n`, `0b` and `0`-octal integers, underscores, sexagesimal numbers |

A `%YAML 1.1` directive selects `Yaml11` for that document. `YamlReaderOptions.Schema` and `YamlLoadOptions.Schema` force one schema for every document, for tools that assume YAML 1.1 without declaring it.

## Comments

Set `YamlReaderOptions.IncludeComments` to receive `YamlTokenType.Comment` events in source order. Comments cost nothing when the option is off.

## Errors and limits

Malformed input throws `YamlException` with a stable `Code` (see `YamlErrorCodes`), line, column and offset. `YamlReaderOptions` and `YamlLoadOptions` bound nesting depth (`MaxDepth`, 256), scalar length (`MaxScalarLength`, 16 MiB of characters) and optionally the whole input (`MaxInputBytes`). `YamlLoadOptions` also bounds the nodes of each document (`MaxNodes`, 1,000,000) and the nodes each document's aliases stand for when resolved (`MaxAliasExpansion`, 100,000). The load measures expansion without expanding anything, so a "billion laughs" input fails quickly, and an alias inside its own anchor fails as unbounded. Exceeding a limit throws `YamlLimitExceededException` with the limit's name. A `CancellationToken` passed to the reader constructor or `LoadAsync` is checked at bounded intervals.
