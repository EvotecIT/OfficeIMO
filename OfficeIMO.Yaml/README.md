# OfficeIMO.Yaml

OfficeIMO's own YAML 1.2 engine. It has no runtime dependencies beyond `OfficeIMO.Core` and targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10.

The package provides `YamlDocument`, a lossless document model you can edit, `YamlReader`, a forward-only pull reader, and `YamlWriter`, a forward-only emitter.

## Loading a document

```csharp
using OfficeIMO.Yaml;

using YamlDocument document = YamlDocument.Load("deployment.yaml");
YamlMapping spec = document.Root!.AsMapping()["spec"].AsMapping();
int replicas = spec["replicas"].GetValue<int>();
string? image = document.Select("spec.template.spec.containers[0].image")?.AsString();
document.Save();   // unchanged: the same bytes as the input
```

`YamlDocument.Load` reads a file or stream, `LoadAsync` does the same asynchronously, and `Parse` takes a string; each accepts a `CancellationToken`. Use `YamlLoadOptions.UntrustedDefaults` for files from outside. `Save()` writes back to the file the document was loaded from; `Save(path)`, `Save(stream)` and `SaveAsync` write elsewhere. The encoding (UTF-8, UTF-16 or UTF-32, either byte order) and byte-order mark are detected and kept, so `Save`, `SaveAsync` and `ToString()` reproduce an unchanged input byte for byte, including comments, blank lines, quoting, key order, line endings and the original spelling of every scalar.

`Documents` lists every document of a multi-document stream as a `YamlDocumentPart` with its root, `%YAML` version, `%TAG` directives and `---`/`...` markers; `Root` is the first document's root. Nodes are `YamlMapping` (ordered entries; string-key lookup), `YamlSequence`, `YamlScalar` and `YamlAlias`. They store offsets into the retained source rather than copies, and scalar values are decoded on first use.

- **Typed reads** go through the document's schema: `GetValue<T>()` and `TryGetValue<T>()` for `string`, `bool`, `int`, `long`, `double`, `float`, `decimal` and their nullable forms; `AsString()` for the decoded text; `IsNull`; `YamlScalar.Kind` for what a scalar resolves to (null, boolean, integer, float, string, timestamp, or a custom tag). `AsMapping()`, `AsSequence()` and `AsScalar()` throw an exception naming the actual node kind. Reading never changes the source text.
- **Aliases** stay `YamlAlias` nodes in the tree. Typed reads and `Target` look through them to the anchored node.
- **Merge keys** (`<<`) apply to key lookups and paths, without being flattened into the tree: under the YAML 1.1 schema, and also under the default Core schema unless the load is strict (see [YAML as real files use it](#yaml-as-real-files-use-it)). A mapping's own keys win, then merged mappings in order; repeated `<<` keys form one merge list.
- **Paths**: `Select("spec.containers[0].image")` returns a node or `null`, and `node.GetPath()` returns a node's path. Use `["key"]` with JSON escapes for keys containing dots or brackets, `[n]` for sequence items, and `$[n]` on a multi-document `YamlDocument` to pick a document. The grammar matches the paths `OfficeIMO.Reader.Yaml` reports. Steps match key text, so values under collection keys (`[a, b]: x`) are reached by enumerating the mapping, and of keys with equal text but different types (`1` and `"1"`) a path finds the one the duplicate-key policy picks.
- **Comments**: `CommentBefore` returns the comment lines above a node's entry, `Comment` the comment at the end of its line, and `SpaceBefore` whether a blank line precedes it. A mapping value shares its key's lines. `YamlDocumentPart.CommentBefore` and `CommentAfter` hold a document's header (such as a license) and footer.
- **Walking**: `node.Descendants()` enumerates a node and everything below it depth first, without allocating per node.
- **Duplicate keys** fail the load by default, as the specification requires. `YamlLoadOptions.DuplicateKeys` set to `FirstWins` or `LastWins` loads such files, keeps every entry and reports each repeat in `Diagnostics`.
- **Custom tags** such as `!Ref` are kept verbatim in `Tag` and never fail a load.

## Editing

```csharp
using YamlDocument document = YamlDocument.Load("deployment.yaml");
document.Set("spec.replicas", 3);                       // only "3" changes in the file
YamlMapping container = document.Select("spec.template.spec.containers[0]")!.AsMapping();
container.Add("imagePullPolicy", "Always").SetComment("pinned by CI");
container.Remove("command");                            // its comments and lines go with it
document.Add("spec.template.spec.containers[0].ports", new YamlMapping { { "containerPort", 8080 } });
document.Save();
```

Saving copies every untouched part of the source and writes only what changed, so comments, blank lines, quoting, indentation, key order and spellings such as `TRUE` or `0x1F` survive everywhere else.

- **Values**: `mapping["key"] = value` or `mapping.Set(key, value)`, `sequence[i] = value`, and `scalar.SetValue(...)` take strings, numbers, booleans, `null` or nodes. A value equal to the current one under the document's schema keeps its source text: setting `true` on `TRUE` changes nothing. A changed string keeps its style when that style can represent it, and otherwise takes the most conservative one that can.
- **Entries and items**: `Add`, `Insert`, `Remove`, `RemoveAt` and `RenameKey` on `YamlMapping`; `Add`, `Insert` and `RemoveAt` on `YamlSequence`. `Add` and `Insert` return the node they placed, so configuration chains: `SetComment`, `SetCommentBefore`, `SetSpaceBefore`, `SetAnchor` and `SetTag` return the node.
- **Paths**: `document.Set(path, value)` creates missing mappings on the way, `document.Add(path, value)` appends to the sequence at a path (creating it), and `document.Remove(path)` returns `false` when nothing is there. A path that goes through a scalar throws.
- **New nodes** follow their neighbors: the document's mapping indentation, sequence style (`key:\n- item` or `key:\n  - item`), quoting convention and line endings. Strings are quoted when a YAML 1.2 or YAML 1.1 reader could read them as anything but a string (`yes`, `on`, `012`, `1:20`).
- **Removing** an entry removes the comment lines above it, its lines and its trailing comment; the neighbors keep theirs. A collection left empty is written as `[]` or `{}`.
- **Anchors and aliases**: `SetAnchor` renames an anchor together with its aliases, `new YamlAlias(node)` refers to an anchored node, and removing a node that aliases still use fails. Editing a node that aliases share records a diagnostic, because the change shows through every alias. Setting a key that only a `<<` merge provides adds an entry of its own; merged mappings are never edited.
- **New documents**: `YamlDocument.Create()` starts an empty stream; set `Root`, use path `Set`, or `AddDocument` for multi-document streams.
- **Reformatting**: `Save(path, new YamlSaveOptions { Reformat = true })` writes the whole document anew through the emitter to normalize its layout, keeping values, anchors, tags and comments.

## Writing YAML

```csharp
using var writer = new YamlWriter(stream);   // UTF-8 without a byte-order mark
writer.WriteStartMapping();
writer.WriteKey("name");
writer.WriteScalar("web");
writer.WriteKey("enabled");
writer.WriteScalar("yes");                   // written 'yes', so YAML 1.1 readers keep it a string
writer.WriteKey("ports");
writer.WriteStartSequence(YamlStyle.Flow);
writer.WriteScalar(80);
writer.WriteScalar(443);
writer.WriteEndSequence();
writer.WriteEndMapping();
```

`YamlWriter` writes to a stream through a pooled UTF-8 buffer or to a `TextWriter`, and allocates nothing per call beyond that buffer. It covers mappings, sequences, keys, scalars with a requested style, aliases, anchors, tags, comments and document markers, and its output is deterministic. `YamlWriterOptions` sets the mapping indent, sequence indent and dash offset (2, 2 and 0 write `key:\n- item`; 2, 4 and 2 write `key:\n  - item`) and the line break. Long lines are never folded, empty collections are written as `[]` and `{}`, and multi-line strings become literal block scalars.

## YAML as real files use it

By default the reader and the loader read three forms the way real files use them and libyaml (behind PyYAML and Ruby) and SharpYaml read them, where the YAML 1.2.2 specification rejects them or reads them differently. Lines inside a flow collection and quoted-scalar continuation lines may be indented no deeper than the enclosing block collection, and `<<` merges under the Core schema:

```yaml
args: [
  --verbose
]                     # a flow collection's lines at the key's indentation
message: 'first line
second line'          # a quoted scalar's continuation at the key's indentation
job:
  <<: *defaults       # merge keys without a %YAML 1.1 directive (docker-compose, GitLab CI)
  <<: *extra          # repeated merge keys add to one merge list
```

Each accepted indentation is a `yaml-lenient-indentation` warning with its line and column, once per outermost flow collection and once per quoted scalar, in `YamlReader.Diagnostics` and `YamlDocument.Diagnostics`. A document part that uses merge keys without a `%YAML 1.1` directive gets one `yaml-merge-key-without-directive` information note in `YamlDocument.Diagnostics`. Indentation that a tab provides is still rejected. `YamlReader.Diagnostics` is empty and allocation-free until a finding occurs.

Set `Strict` on `YamlReaderOptions` or `YamlLoadOptions` for the specification's exact rules: both indentation forms are syntax errors, and `<<` is an ordinary key unless the YAML 1.1 schema applies, so a repeated one is a duplicate key.

Saving copies accepted text that an edit does not own unchanged. New and edited text is always written in the valid form.

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

Malformed input throws `YamlException` with a stable `Code` (see `YamlErrorCodes`), line, column and offset; when the source text is at hand, the message ends with the offending line and a caret under the column. `YamlReaderOptions` and `YamlLoadOptions` bound nesting depth (`MaxDepth`, 256), scalar length (`MaxScalarLength`, 16 MiB of characters) and optionally the whole input (`MaxInputBytes`). `YamlLoadOptions` also bounds the nodes of each document (`MaxNodes`, 1,000,000) and the nodes each document's aliases stand for when resolved (`MaxAliasExpansion`, 100,000). The load measures expansion without expanding anything, so a "billion laughs" input fails quickly, and an alias inside its own anchor fails with `YamlErrorCodes.RecursiveAlias`. Exceeding a limit throws `YamlLimitExceededException` with the limit's name. A `CancellationToken` passed to the reader constructor or `LoadAsync` is checked at bounded intervals.
