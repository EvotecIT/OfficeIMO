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

`TomlReader` is a class over `string`, `ReadOnlyMemory<char>`, a `TextReader` or a UTF-8 `Stream`, with an explicit state machine, so deep input cannot overflow the stack. Tokens are header brackets, key segments (a dotted key is several `Key` tokens), values, array and inline-table brackets and, on request, comments. The reader holds only offsets, line counters and a small context stack, so it allocates nothing per token; `GetString()` decodes on demand and `ValueSpan` exposes the source.

Stream and `TextReader` input go through a pooled sliding buffer, as `YamlReader`'s do, for lock files too large to hold as one string. Until the input ends, the scanner sees the buffered text only through its last line feed. Every TOML token except a multi-line string ends on its line, so only two places can reach the edge of the visible text: skipping blank lines and comments, which happens between tokens, and a multi-line string, which keeps its start in the buffer while more is read. Everything else scans exactly as in memory, which is why tokens, positions and errors match the in-memory reader wherever the buffer boundaries fall. The buffer drops consumed text when that frees half of it and otherwise doubles, so it grows only while one line or multi-line string is longer than it; a multi-line string stops growing it at `MaxStringLength`. UTF-8 is decoded strictly by the reader's own decoder, which stops exactly before an invalid or truncated sequence, so the error carries the line and column of the first bad byte. `TomlDocument` keeps loading the whole text into one string, since lossless saving needs it; `OfficeIMO.Reader.Toml` projects the model's tables in document order and so loads it too.

The reader checks the whole TOML 1.1 syntax: control characters, `\r` outside `\r\n`, escapes and their scalar values, string delimiters and quote runs, integer ranges and bases, float forms, date and time fields with leap years, optional seconds, and where newlines, commas and trailing commas may appear. Semantic rules that need the whole document stay in the model, as duplicate keys stay outside `Utf8JsonReader`.

### Document model

- `TomlTable` (root, header, implicit, dotted, inline or array item; `TableKind` tells which), `TomlTableArray`, `TomlArray` and `TomlValue`. Values keep their source span and decode on first read; a table is one logical object even when headers and dotted keys spread it over the file.
- The tree builder applies the table rules: unique keys; a `[header]` only for a new or implicit table; dotted keys only through tables they created themselves; inline tables complete as written; static arrays and arrays of tables never mixed. It counts nodes and depth, counting each dotted key segment.
- Comments and whitespace stay in the source. `LeadingComments` and `TrailingComment` read them lazily.
- A loaded value costs one 56-byte object. A node keeps its key, its source offsets and one reference that is its parent, or the document for the root and for nodes not yet added; the document is reached through the parents. The end of a key is found from the value's start rather than stored, and nesting depth is tracked by the builder, since only loading checks it. Comments set by editing and nodes removed by editing live in side tables on the document. The items of an inline table or array go to an array of their exact count when it closes. Short keys and string values without escapes come from a per-thread cache of recent strings, which removes most key allocations in lock files.

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
- **Stream input:** every toml-test case and corpus file reads through 2- to 64-character buffers, one byte or character per read, with the same tokens, positions and errors as in memory.
- **Large corpus** (opt-in): 5,433 real-world files (236 MB) pinned by source URL and SHA-256 in `OfficeIMO.TestAssets/Toml/LargeCorpus/sources.json` and fetched into a local cache that is never committed: crates.io manifests as published (vendored in Firefox), workspace manifests and `Cargo.lock` files (polkadot-sdk, Zed, rustc, Bevy, Deno, windows-rs), `pyproject.toml` files (DataDog integrations, LlamaIndex, uv, ruff, PDM), `uv.lock` and `poetry.lock` files (LlamaIndex, LangChain, pydantic, uv, ruff, Poetry's fixtures), editor and prompt configuration (Helix, Starship) and Hugo sites and themes. Every file is loaded, compared with Tomlyn's data, saved back byte for byte, read through the stream reader with the same tokens as in memory, and edited with a fixed script (set a value, add a key, remove a key, add a table) whose saved text must reload to the edited model and change only the lines involved. A seeded mutation run (truncation, byte flips, line edits, invalid UTF-8, control characters, deep arrays, inline tables, dotted keys and headers, huge strings, key and table floods, unclosed and quote-run strings, lone carriage returns) feeds hostile variants of a corpus sample to the loader with `UntrustedDefaults` and to the stream reader: only `TomlException` may escape, nothing may run past its time limit, allocation stays within budget and every mutant that loads saves back byte for byte. [Running them](#running-the-large-corpus) has the commands.
- **Product tests** cover what the oracle cannot: limits, cancellation, encodings, error codes and positions, edit rules, the reader's token contract and the writer.
- **Serializer:** generated and reflection metadata write the same TOML for every model, read it back losslessly, and Tomlyn reads every output to the same values the model sees. `OfficeIMO.Toml.AotSmoke` publishes the generated serializer with NativeAOT and no trim or AOT warnings.
- **Benchmarks** (opt-in, `OfficeIMO.Toml.Benchmarks`) compare parse, load and unchanged save with Tomlyn and Tomlet on corpus workloads, and model deserialization and serialization with Tomlyn's reflection and generated serializer, after checking that every library sees the same values.

### Running the large corpus

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --fetch-corpus
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --large-corpus
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --mutations --seed 20261113
```

`--fetch-corpus` downloads every pinned file from `raw.githubusercontent.com` at its pinned commit and checks its SHA-256; `--refresh-corpus` re-lists the sources' trees and re-pins them (set `GITHUB_TOKEN` to raise the API rate limit). The cache is `--cache`, else `OFFICEIMO_TOML_CORPUS_CACHE`, else `toml-corpus` under the temporary folder. Reports (`large-corpus.md`, `large-corpus-failures.tsv`, `mutations.md` and failing mutants) go to `reports` in the cache or `--output-dir`. `--file` runs one file; `--kind` and `--sample` narrow the mutation run, whose mutants depend only on the seed, file, kind and variant.

Last run (macOS arm64, .NET 10): all 5,422 files that load save back byte for byte, read the same through the stream reader and pass the edit script. Eleven files are rejected: ten are invalid on purpose (template placeholders such as `{check_name}`, and invalid-input fixtures of clippy and Poetry, which Tomlyn rejects too), and one is clippy's duplicate-key fixture, which TOML forbids and Tomlyn accepts. Five files load to data that differs from Tomlyn's: each reopens an array of tables (`[[example]]`, `[[grammar]]`) after other tables, where Tomlyn reorders or drops tables and OfficeIMO keeps source order. The run found two editing defects, both fixed: removing the last key of a table made only by dotted keys or deeper headers left an empty table in the model that the saved text did not have, and removing the last line of a file without a final line break left an empty line before text added at the end. The mutation run (400 files, 18 kinds, 3 variants, 21,600 mutants) had no failures; the slowest mutant took 364 ms.

## Performance

`OfficeIMO.Toml.Benchmarks` (BenchmarkDotNet, default job) on an Apple M4, macOS 27.0.1, .NET 10.0.15 Arm64, 13 November 2026. Ratios are to Tomlyn 2.10.1 (lower is better); `WithValues` also decodes every value, as Tomlyn's loader does. The earlier run is the engine before the stream reader and the allocation work in this section; the toml-test suite, the committed and large corpora and a differential run of both builds (data, positions, comments, saved bytes, reader tokens and scripted edits over 6,131 files) give identical results before and after.

| Parse | OfficeIMO time | Tomlyn time | Time ratio | OfficeIMO allocated | Tomlyn allocated | Allocation ratio (earlier) |
|---|---:|---:|---:|---:|---:|---:|
| SmallConfig | 9.9 μs | 36.5 μs | 0.27 | 1.64 KB | 8.28 KB | 0.20 (0.51) |
| CargoWorkspace | 63.4 μs | 226.5 μs | 0.28 | 22.6 KB | 68.8 KB | 0.33 (0.95) |
| LockFile | 570 μs | 1,434 μs | 0.40 | 603 KB | 807 KB | 0.75 (1.00) |
| Corpus | 854 μs | 2,456 μs | 0.35 | 796 KB | 1,145 KB | 0.70 (0.96) |

| Load | OfficeIMO time (`WithValues`) | Tomlyn time | Time ratio | OfficeIMO allocated (`WithValues`) | Tomlyn allocated | Allocation ratio (earlier) |
|---|---:|---:|---:|---:|---:|---:|
| SmallConfig | 10.1 μs (12.3 μs) | 49.4 μs | 0.21 (0.25) | 7.38 KB (8.80 KB) | 14.53 KB | 0.51 / 0.61 (0.92 / 1.13) |
| CargoWorkspace | 81.8 μs (97.7 μs) | 267.4 μs | 0.31 (0.37) | 124.0 KB (143.9 KB) | 169.0 KB | 0.73 / 0.85 (1.13 / 1.30) |
| LockFile | 581 μs (864 μs) | 1,551 μs | 0.37 (0.56) | 467.8 KB (1,093.2 KB) | 1,121.3 KB | 0.42 / 0.97 (0.67 / 1.25) |
| Corpus | 1,887 μs (1,564 μs) | 2,687 μs | 0.70 (0.58) | 973.8 KB (1,749.4 KB) | 1,841.4 KB | 0.53 / 0.95 (0.82 / 1.26) |

Before this work OfficeIMO already parsed and loaded faster, but allocated more than Tomlyn when loading every workload with values, and as much when parsing the lock file. The changes, all internal: the reader and the decoder reuse short repeated keys and values from a per-thread cache; tables no longer keep a parallel key array; a node's parent and document share one field, the key end and nesting depth are computed instead of stored, and edit-only state moved to side tables on the document, taking a value node from 80 to 56 bytes and a table from 112 to 72; inline tables and arrays get exact-size storage when they close; and tables, arrays and arrays of tables enumerate with struct enumerators. Unchanged load and save runs at 0.04 to 0.20 of Tomlyn's syntax-tree round trip in time and allocation. The Corpus load times vary more between runs than the others on this machine. Windows and Linux runs are still open on the [roadmap](ROADMAP.md#toml-engine).

```sh
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --validate-workloads
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Toml.Benchmarks -- --filter "*TomlParseBenchmarks*" "*TomlLoadBenchmarks*"
```
