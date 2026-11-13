# TOML engine design

`OfficeIMO.Toml` is OfficeIMO's own TOML 1.1 engine. It reads, edits and writes TOML without changing anything the caller didn't change, and has no runtime dependencies beyond `OfficeIMO.Core`. It follows the [YAML engine design](officeimo.yaml-engine-design.md) wherever the formats allow. This document owns the architecture; the [package README](../OfficeIMO.Toml/README.md) owns the supported API, and open work belongs in the [roadmap](ROADMAP.md#toml-engine).

## Goals

1. **Lossless by construction.** Loading and saving an unchanged file reproduces it byte for byte. An edit changes only the text of what was edited; comments, blank lines, alignment, quoting, key order, line endings, the BOM and the original spelling of values (`0xFF`, `1_000`, `'literal'`, `1979-05-27 07:32:00Z`) survive everywhere else.
2. **Complete TOML 1.1.** [toml-test](https://github.com/toml-lang/toml-test) is the oracle. Cases the engine does not pass go in a checked-in exception list that only shrinks; today it is empty.
3. **Performance from the first commit.** A single pass over the input, no per-token allocation, one object per node, and values decoded only when read.
4. **Safe on hostile input.** Bounded depth, node count, string length and input size, plus cancellation.
5. **One OfficeIMO vocabulary.** The [unified API design](officeimo.api-design.md) applies as written.

Tomlyn (BSD-2, TOML 1.1, with a trivia-preserving syntax tree) is the main .NET TOML library. We own ours for API consistency with the other OfficeIMO formats and for no runtime dependencies. Tomlyn and Tomlet appear only in the opt-in benchmark project.

## Public API

- **Lifecycle:** `TomlDocument.Create()`, `Create(path)`, `Load(path|stream, options)`, `LoadAsync`, `Parse(text, options)`, `Save()`, `Save(path|stream)`, `SaveAsync`, `ToString()`. `TomlDocument` implements `IDisposable` so `using` works as for every format. Options are `TomlLoadOptions` (with `UntrustedDefaults`), `TomlReaderOptions` and `TomlWriterOptions`; there is no save options type until a save setting has a scenario.
- **Authoring:** one object model. `Add`, `AddTable`, `AddInlineTable`, `AddArray`, `AddTableArray`, `Insert` and `Set` return the node they create or change, and `SetComment` returns its node, so calls chain.
- **Paths:** `Select`, `Set` and `Remove` take the same grammar as YAML paths and as the paths `OfficeIMO.Reader.Toml` reports, so a Reader row's path edits the same value.
- **Errors:** `TomlException` carries a stable `TomlErrorCodes` code and a position. TOML has no recoverable findings, so there are no load diagnostics; every rule violation fails the load, as the specification requires.
- **Positions:** 0-based offsets, 1-based lines and columns, like YAML.

## Layers

```
TomlSerializer (+ source generator)     objects ⇄ TOML                       [T4]
TomlDocument / TomlNode                 lossless editable model              [T2, T3]
TomlReader / TomlWriter                 forward-only pull tokens / emitter   [T1, T3]
Scalars (internal)                      validation and lazy decoding         [T1]
Source text (internal)                  decoded once, kept for lossless save [T1]
```

### Source text

TOML is UTF-8. A file or stream is decoded once with a strict decoder, so invalid bytes fail with their line. A UTF-8 BOM is accepted and recorded so a save writes it back. Line endings are never normalized in the source; multi-line string values normalize `\r\n` to `\n`.

### Reader

`TomlReader` is a class over `string` or `ReadOnlyMemory<char>` with an explicit state machine, so deep input cannot overflow the stack. Tokens are header brackets, key segments (a dotted key is several `Key` tokens), values, array and inline-table brackets and, on request, comments. The reader holds only offsets, line counters and a small context stack, so it allocates nothing per token; `GetString()` decodes on demand and `ValueSpan` exposes the source.

The reader checks the whole TOML 1.1 syntax: control characters, `\r` outside `\r\n`, escapes and their scalar values, string delimiters and quote runs, integer ranges and bases, float forms, date and time fields with leap years, optional seconds, and where newlines, commas and trailing commas may appear. Semantic rules that need the whole document stay in the model, as duplicate keys stay outside `Utf8JsonReader`.

### Document model

- `TomlTable` (root, header, implicit, dotted, inline or array item; `TableKind` tells which), `TomlTableArray`, `TomlArray` and `TomlValue`. Values keep their source span and decode on first read; a table is one logical object even when headers and dotted keys spread it over the file.
- The tree builder applies the table rules: unique keys; a `[header]` only for a new or implicit table; dotted keys only through tables they created themselves; inline tables complete as written; static arrays and arrays of tables never mixed. It counts nodes and depth, counting each dotted key segment.
- Comments and whitespace stay in the source. `LeadingComments` and `TrailingComment` read them lazily.

### Editing and saving

Saving walks only nodes marked changed and turns each change into a text edit at a source position; everything else is copied from the source. That is the whole minimal-diff mechanism:

- **Changed value:** its span is replaced. An equal value keeps its spelling; otherwise the current style is kept when it can hold the new value (base, literal quoting, multi-line form, date-time separator).
- **New key/value:** a line after the table's last line in its section, with that line's indentation and `=` spacing. A dotted table gets a dotted key next to its siblings; an implicit table gets its own header before its first sub-table.
- **New table or array of tables:** a header section after the last section of the parent's subtree, so `[[x]]` and sub-table headers keep their meaning. A blank line separates sections when the file does that.
- **Arrays and inline tables:** new items follow the container's layout (a new indented line in a multi-line array, keeping its trailing-comma style; `, item` on one line otherwise). Removed items take exactly one separator; an item alone on its line takes its line.
- **Removed key or table:** the key/value line with the comment lines directly above it, or every section of the removed subtree, wherever it is in the file.
- **Comments:** only the comment text after a key/value pair or header changes.

New text uses the file's line endings, its key separator and its comment spacing. New strings follow the quoting of their siblings. Output stays readable by TOML 1.0 parsers: no `\e` or `\x` escapes, and times always have seconds.

### Serializer

`TomlSerializer` maps .NET objects to TOML documents with `YamlSerializer`'s design, so a model maps to both formats the same way:

- **One vocabulary.** The mapping attributes (`OfficePropertyName`, `OfficeIgnore`, `OfficeRequired`, `OfficeConstructor`, `OfficeDerivedType`, `OfficePolymorphic`), `OfficeNamingPolicy` and `OfficeIgnoreCondition` live in `OfficeIMO.Core`, which both engines reference. Each format keeps its own options, context, `[TomlSerializable]` root attribute and errors. `[TomlInlineTable]` is the one TOML-only attribute.
- **One shape rule set, two metadata sources.** Reflection metadata and the source generator classify types with the same rules as YAML: the runtime rules are shared source (`OfficeIMO.SharedSource/Serialization`), and the generator's parser and emitter are shared source (`OfficeIMO.SharedSource/SerializerGeneration`) compiled into each format's generator with its own diagnostics (`TOML001`–`TOML005`). The generator ships inside the `OfficeIMO.Toml` package, and `buildTransitive/OfficeIMO.Toml.targets` turns reflection off for `PublishAot` and `PublishTrimmed` applications.
- **Reading** loads a `TomlDocument`, so every table rule, limit and position is the model's. TOML tables can be defined out of order (`[a.b]` before `[a]`, dotted keys, headers spread over the file), so the serializer maps the finished tree instead of streaming tokens into objects. Values decode from their source text by target type, without the boxed `TomlValue.Value`. Errors are `TomlException`s with the node's position and path.
- **Writing** goes straight to text through the formatting code `TomlWriter` and model editing use. Each table writes its key/value pairs first and its tables after them as `[header]` and `[[header]]` sections; a header is written only when its table writes a pair or ends empty, so tables that hold only sub-tables need none. Inside arrays and inline tables everything is inline. Untyped and polymorphic values choose their layout by value.
- **Null** has no TOML form. Null properties and dictionary values are left out by default (`TomlNullHandling.Omit`), or fail (`Error`); null array items always fail. Polymorphic values carry their discriminator as a key, `$type` unless named.

### Writer

`TomlWriter` is a forward-only emitter, modeled on `Utf8JsonWriter`, sharing its formatting code with model editing so a value is spelled the same whichever API writes it. It writes to a `TextWriter` or a UTF-8 stream and rejects calls TOML does not allow at that point.

## Resource and security limits

| Limit | Default | `UntrustedDefaults` |
|---|---|---|
| `MaxDepth` (tables, arrays and inline tables, counting dotted segments) | 256 | 64 |
| `MaxNodes` | 1,000,000 | 200,000 |
| `MaxStringLength` (one string or key token, UTF-16 units) | 16 MiB | 1 MiB |
| `MaxInputBytes` | none | 16 MiB |

Exceeding a limit throws `TomlLimitExceededException` with the limit's name. Loading and the reader check cancellation at bounded intervals.

## Package boundaries

| Package | Depends on | Owns |
|---|---|---|
| `OfficeIMO.Toml` | `OfficeIMO.Core` | Reader, writer, model, editing, serializer, generator (built from `OfficeIMO.Toml.SourceGeneration`, packed as an analyzer) |
| `OfficeIMO.Reader.Toml` | `OfficeIMO.Reader.Core`, `OfficeIMO.Toml` | Path/type/value chunks and Markdown projection |

Targets: netstandard2.0, net472, net8.0 and net10.0, with `System.Memory` and `System.Buffers` as the only polyfills.

## Qualification

- **toml-test 2.2.0, TOML 1.1.0:** every valid case decodes to its expected JSON and saves back byte for byte; every invalid case is rejected with `TomlException`. The harness compares values with its own independent date-time and float parsing.
- **Real-world corpus** with recorded provenance (`pyproject.toml`, `Cargo.toml`, `Cargo.lock`, `uv.lock`, ruff, Hugo, Netlify, Starship): every file round-trips byte for byte, and scripted edits change exactly the expected lines.
- **Product tests** cover what the oracle cannot: limits, cancellation, encodings, error codes and positions, edit rules, the reader's token contract and the writer.
- **Serializer:** generated and reflection metadata write the same TOML for every model, read it back losslessly, and Tomlyn reads every output to the same values the model sees. `OfficeIMO.Toml.AotSmoke` publishes the generated serializer with NativeAOT and no trim or AOT warnings.
- **Benchmarks** (opt-in, `OfficeIMO.Toml.Benchmarks`) compare parse, load and unchanged save with Tomlyn and Tomlet on corpus workloads, and model deserialization and serialization with Tomlyn's reflection and generated serializer, after checking that every library sees the same values.
