# OfficeIMO.Yaml

OfficeIMO's own YAML 1.2 engine. It has no runtime dependencies beyond `OfficeIMO.Core` and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

The package provides `YamlReader`, a forward-only pull reader over YAML 1.2 text.

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

A `%YAML 1.1` directive selects `Yaml11` for that document. `YamlReaderOptions.Schema` forces one schema for every document, for tools that assume YAML 1.1 without declaring it.

## Comments

Set `YamlReaderOptions.IncludeComments` to receive `YamlTokenType.Comment` events in source order. Comments cost nothing when the option is off.

## Errors and limits

Malformed input throws `YamlException` with a stable `Code` (see `YamlErrorCodes`), line, column and offset. `YamlReaderOptions` bounds nesting depth (`MaxDepth`, 256), scalar length (`MaxScalarLength`, 16 MiB of characters) and optionally the whole input (`MaxInputBytes`). Exceeding a limit throws `YamlLimitExceededException` with the limit's name. A `CancellationToken` passed to the constructor is checked at bounded intervals.
