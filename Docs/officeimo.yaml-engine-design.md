# YAML engine design

`OfficeIMO.Yaml` is OfficeIMO's own YAML engine. It reads, edits, and writes YAML without changing anything the caller didn't change. The package has no runtime dependencies beyond `OfficeIMO.Core`, like `OfficeIMO.CSV`. This document owns the architecture. The package README will own the supported API, and open work belongs in the [roadmap](ROADMAP.md#yaml-engine).

## Goals

1. **Lossless by construction.** Loading and saving an unchanged file reproduces it byte for byte. An edit changes only the text of the edited node; comments, blank lines, indentation, quoting, key order, line endings, the BOM and the original spelling of scalars (`TRUE`, `yes`, `0x1F`, `~`) survive everywhere else.
2. **Complete YAML 1.2.2 syntax.** The [yaml-test-suite](https://github.com/yaml/yaml-test-suite) is the oracle. Every case with a known exception is listed in a checked-in exception file with its reason. That list only shrinks.
3. **Performance from the first commit.** Low allocation and a single pass over the input. The pull reader allocates nothing per event, and each layer can be used without paying for the layers above it.
4. **Safe on hostile input.** The engine has bounded depth, node counts, alias expansion and scalar sizes, plus cancellation. Tags never activate .NET types.
5. **One OfficeIMO vocabulary.** The package follows the [unified API design](officeimo.api-design.md) from its first release; see [Public API](#public-api). The package is new, so its API is shaped for that contract now. Earlier YAML behavior in `OfficeIMO.Reader.Yaml` changes where the new design is better, and `MIGRATION.md` records the upgrade actions.

Replacing YamlDotNet in `OfficeIMO.Reader.Yaml` is an outcome of this design, not a constraint on it. The Reader adapter adapts to the engine, never the other way round.

## Public API

The C# rules and PowerShell rules in the [API design](officeimo.api-design.md) apply as written. For YAML they mean:

- **Lifecycle:**
  ```csharp
  using var doc = YamlDocument.Create();
  using var doc = YamlDocument.Load("values.yaml", YamlLoadOptions.UntrustedDefaults);
  await using var doc = await YamlDocument.LoadAsync(stream, options);
  YamlDocument doc = YamlDocument.Parse(text);
  doc.Save();                 // to the path it was loaded from
  doc.Save("copy.yaml");
  await doc.SaveAsync(stream);
  ```
  - `YamlDocument` implements `IDisposable`, so `using` works as it does for every format.
  - Options are `YamlLoadOptions` (with `UntrustedDefaults`), `YamlSaveOptions`, `YamlReaderOptions` and `YamlWriterOptions`.
- **Authoring:** one object model.
  - `Add*` and `Insert*` methods return the node they create.
  - Mutators that configure a node return that node, so calls chain naturally: `mapping.Add("image", "nginx").SetComment("pinned by CI")`.

  A separate fluent builder is a thin wrapper over these same names, never a second model.
- **Conversion grammar:** when routes are added, they follow the shared `To{Target}`/`SaveAs{Target}` forms. Examples are `ToJson()` and the JSON-to-YAML route on the JSON side, plus the Markdown front matter integration. YAML has no format-engine dependencies of its own.
- **Diagnostics:** recoverable findings are `OfficeDiagnostic` values whose codes come from the Core diagnostic catalog. Exceptions keep their stable `YamlErrorCodes` for syntax and limit failures.
- **Positions:** C# indexes are 0-based like every .NET collection; PowerShell positions are 1-based like every PSWriteOffice command.
- **Visibility:** public only what the README and examples use. The package adopts the public API analyzer baseline (`Build/ApiMap/PublicApi.targets`) before its first release, so later surface changes are deliberate.
- **AOT:** `IsAotCompatible` and the trimming analyzers are on for net8.0 and net10.0.

PSWriteOffice gets YAML commands built on the same public API, with `Yaml` as the format name:

| Command | Job |
|---|---|
| `New-OfficeYaml` / `Open-OfficeYaml` / `Save-OfficeYaml` / `Close-OfficeYaml` | The standard document lifecycle, including the content-block form that saves on exit |
| `Get-OfficeYamlNode -Select` / `Set-OfficeYamlNode -Select -Value` / `Remove-OfficeYamlNode -Select` | Read and edit by path inside a live document. `-Select` uses the C# path grammar; `-Path` stays the file, as everywhere |
| `Import-OfficeYaml -Path` / `Export-OfficeYaml -InputObject -Path` | Objects from and to files through the serializer |
| `ConvertFrom-OfficeYaml` / `ConvertTo-OfficeYaml` | Text to objects and objects to text, following PowerShell's `ConvertFrom-Json`/`ConvertTo-Json` pattern |
| `Test-OfficeYaml` | Validate without changing the file, reporting diagnostics |
| `Merge-OfficeDocument` | The shared cross-format command; for YAML it overlays files with the [merge](#merging) rules |

Typical script:

```powershell
Open-OfficeYaml -Path .\deployment.yaml {
    Set-OfficeYamlNode -Select 'spec.replicas' -Value 3
}
```

The PowerShell commands follow the C# API. Any capability they need that the library lacks becomes a library feature.

## Non-goals

- Unity's YAML dialect and other formats that only look like YAML. They need a concrete consumer and a separate decision.
- Arbitrary type activation from tags (`!!python/object`, `!<tag:...>` mapped to CLR types). Tags map only to types the caller registers.
- A YAML-to-anything conversion zoo. Conversions go through Workflows routes when a consumer needs them.

## Layers

Each layer depends only on the layers below it. A caller that needs only events never builds a tree, and a caller that needs a tree never pays for serialization.

```
YamlSerializer (+ source generator)     objects ⇄ YAML                       [Y4]
YamlDocument / YamlNode                 lossless editable model              [Y2, Y3]
YamlReader / YamlWriter                 forward-only pull events / emitter   [Y1, Y3]
Scanner (internal)                      tokens with positions + trivia flags [Y1]
Source text (internal)                  decoded once, kept for lossless save [Y1]
```

### Source text

- Encoding detection follows the specification. UTF-8 is the default, and a BOM or the first bytes select UTF-16 or UTF-32. The detected encoding and BOM presence are recorded so a save reproduces them.
- Input is decoded once into a single `char` buffer. `YamlDocument` retains that buffer as its source of truth for untouched text. Retaining it is the cheapest lossless representation: every node is two integers into it.
- `YamlReader` over a `Stream`/`TextReader` uses a pooled sliding buffer instead. It retains only the window it needs, and the window grows only while one token is longer than the buffer.
- Line endings are not normalized in the source. Scalar *values* fold line breaks as the specification requires, while the source keeps the original `\r\n`/`\n`.
- The scanner reads UTF-16 only; there is no separate UTF-8 byte scanner. Decoding UTF-8 input costs about 2% or less of deserializing it in the serializer workloads, and the scanner's per-character work does not depend on the character width, so a second scanner would double the code that conformance and exact spans depend on for no measured gain.

### Scanner (internal)

- A `ref struct` over `ReadOnlySpan<char>` that produces `YamlToken` structs: kind, start, length and flags. Token creation never allocates.
- Whitespace, comments and line breaks are not separate token objects. They are the gaps between token spans, recoverable from the source on demand. Flags mark tokens whose scalar needs decoding (escapes, folding, multi-line plain scalars) so the common case skips decoding entirely.
- Hot loops use `IndexOfAny`/`SearchValues` on net8.0+ (vectorized) and a scalar fallback on netstandard2.0/net472, the same split as the CSV parser. No regular expressions, no LINQ and no closures in the scanner or parser.
- Indentation tracking, simple-key candidates (with the specification's 1024-character limit) and flow-level state live in small pooled stacks.

### Parser and `YamlReader`

- The parser is an explicit state machine over the specification's productions. There is no recursion, so deep inputs can't overflow the stack, and `MaxDepth` is enforced at one point.
- `YamlReader` is the public pull API, modeled on `Utf8JsonReader`/`XmlReader`. `Read()` advances to the next event: stream, document or collection start/end; scalar; alias.
  - Properties: `TokenType`, `Style`, `Tag`, `Anchor`, `Start`/`End` positions (offset, line, column), `ValueSpan` (raw source) and `GetString()` (decoded and cached).
  - Typed accessors resolve through the active schema: `TryGetInt64`, `TryGetDouble`, `TryGetBoolean`, `TryGetDecimal`, `IsNull`.
  - It allocates nothing per event unless the caller asks for a string.
- `YamlReaderOptions` holds the schema, limits, `Strict` and `IncludeComments` (comment events are off by default and cost nothing when off). The duplicate-key policy is a load option, because only the model compares keys.
- Errors throw `YamlException` with a stable code, line, column and offset. `YamlException.ToDiagnostic()` gives consumers (the Reader adapter, Markdown front matter, PowerShell) the failure as an error `OfficeDiagnostic`: its code follows the catalog's `Yaml.{area}.{rule}` → `yaml-{area}-{rule}` pattern, except that a strict duplicate key keeps `yaml-duplicate-key`, the code of the non-strict warning. Recoverable findings are `OfficeDiagnostic` values: `YamlReader.Diagnostics` returns a shared empty list until the scanner records one, so a clean read allocates nothing for it, and `YamlDocument.Diagnostics` collects the reader's findings with the model's own.

### Document model

```csharp
YamlDocument doc = YamlDocument.Load("deployment.yaml");
YamlMapping spec = doc.Root.AsMapping()["spec"].AsMapping();
spec["replicas"] = 3;                      // only "replicas: <value>" text changes
doc.Save("deployment.yaml");
```

- `YamlDocument` is the file, matching `CsvDocument` and `WordDocument`. It exposes `Documents` (`YamlDocumentPart`: directives, explicit `---`/`...` markers and a `Root`) and `Root` as a shortcut to the first part.
- Node types:
  - `YamlNode` is the abstract base, holding style, tag, anchor and source span.
  - `YamlMapping` holds an ordered list of `YamlMappingEntry`, each with a key node and a value node, so complex keys are representable. It has a string-key lookup cached on first use.
  - `YamlSequence`.
  - `YamlScalar` holds raw text, style and a lazily decoded, cached value.
  - `YamlAlias` points to its anchor. Aliases stay aliases in the model; resolution is a view.
- Nodes store spans into the retained source. Trivia (comments, blank lines, indentation) stay in the source gaps and aren't materialized unless asked for: `node.CommentBefore`, `node.Comment` and `node.SpaceBefore` read the gaps lazily, and `YamlDocumentPart.CommentBefore`/`CommentAfter` hold a document's header and footer. `YamlDocumentPart.Comment` is the comment at the end of the `---` line when nothing else starts there; when the root or its properties start on that line, the comment is the root's. The model allocates about one object per node, with nothing for trivia.
- **Source positions:** `node.Start`/`End` turn the stored offsets into line and column through one line table the document builds on first use, so loading allocates nothing for positions. `node.IsModified` reads the edit flags: whether the node or anything inside it changed since load, so a consumer knows whether the position still describes the text that saving writes. Nodes whose content comes from editing have no position.
- Typed reads go through the schema: `scalar.GetValue<T>()`, `TryGetValue<T>`, `AsString()`. Reading never changes raw text.
- **Assigning a value:**
  - If the new value equals the old one under the active schema, the raw text is kept, so assigning `true` to a scalar spelled `TRUE` changes nothing.
  - Otherwise the scalar is re-rendered. It keeps the old style if that style can represent the new value; otherwise it uses the most conservative style that round-trips.
- **Inserting a new node:** the node takes its style from its siblings: indentation width, block or flow, sequence indentation style ("- " under a key or indented), and the quoting convention. The document's mapping indent, sequence indent and dash offset (the writer's three numbers) and its dominant line break are read once, from a bounded sample, when the first edit needs them.
- **Saving:** untouched nodes are written as copies of their source text, and only dirty subtrees are emitted through the same emitter core as `YamlWriter`. That is the whole minimal-diff mechanism. There is no reformat pass unless the caller asks for one (`YamlSaveOptions.Reformat`). Editing never re-parses, and the save walk descends only into nodes that contain an edit.
- **Trivia ownership.** Every edit owns an exact region of the source:
  - The comment lines directly above an entry, with no blank line between, and the blank lines above them belong to the line where the entry starts. A comment block that a blank line separates from the entry belongs to no entry: in real configuration files it is a section note or a commented-out example, such as the disabled `nodeAffinity` blocks in Helm values. Its owner is the outermost node starting there: a mapping value shares its key's lines, and the first entry of a block collection without properties that is a sequence item or the root shares the collection's.
  - Above the first entry of a document without `---`, comment lines that a blank line separates from it are the document's header. Comment lines before `---` are also the header; comment lines after the root's last line are the footer. Setting the `---` line's comment owns the rest of that line; a replaced comment keeps the spacing before it.
  - A changed scalar, alias or flow collection that stays on its line owns its own span, properties included; its trailing comment stays.
  - A mapping value that changes between inline and block layout owns the text from the key's `:` to the end of its last line, and the comment on the key's line is written again after the new value or on the key line. A block collection replacing a sequence item or root drops the replaced value's trailing comment, because it has no line of its own.
  - Removing a block entry removes its blank lines, its attached comment lines, its lines and its trailing comment; neighbors and comment blocks that belong to no entry keep theirs. Below a document header the blank lines stay, so the header stays a header. A removed first entry that shares its line with a dash takes the text up to the next entry, which moves onto the dash line. A removed flow item takes the separator before it, or after it when it is first.
  - An inserted block entry goes above the blank lines and comments of the entry it precedes, or after the last entry's last line. A collection left with none of its loaded children is written whole in place of its own span, and an empty block collection becomes `[]` or `{}`.
- **Never** (bugs seen in other tools, made structurally impossible by copying source text): no `!!merge` tag is added to `<<` keys, untouched sequence indentation is never normalized, and a comment never moves to a different node on save. The scripted corpus edits check all three on every saved result.
- **Path addressing** for configuration edits: `doc.Select("spec.template.spec.containers[0].image")` returns a node, `doc.Set(path, value)` creates missing mappings on the way, `doc.Add(path, value)` appends to a sequence, and `doc.Remove(path)` returns `false` when nothing is there; traversing a scalar throws. Paths don't evaluate expressions. They share their grammar with the paths `OfficeIMO.Reader.Yaml` reports, so a Reader chunk path can be fed back to edit the same node. Steps match key text, so values under collection keys are reached by enumerating the mapping.
- Anchors and aliases:
  - Renaming an anchor updates its aliases.
  - Editing through an alias edits the anchored node and is reported once per anchored node (`yaml-aliased-node-edited`). Setting a key that only a merge provides adds an own entry; merged mappings are never edited.
  - `<<` merge keys (YAML 1.1 schema, and Core unless strict; see [Schemas and versions](#schemas-and-versions)) are resolved in the typed view, never flattened in the tree.

### Merging

Layered configuration is a common YAML job: Helm values overlays, per-environment settings, and defaults plus overrides. Today people do it with `yq` and lose formatting. The engine merges an overlay into a base document through the editing API, so the result is a minimal diff of the base file.

```csharp
using var baseDoc = YamlDocument.Load("values.yaml");
using var prod = YamlDocument.Load("values.prod.yaml");
YamlMergeReport report = baseDoc.Merge(prod);
baseDoc.Save("values.merged.yaml");
```

- `YamlDocument.Merge(overlay, options)` merges each document part by position. `YamlMapping.Merge(overlay, options)` merges any subtree, for example after `Select`.
- **Mappings** merge key by key, recursively.
- **Scalars:** overlay scalars replace base scalars through the normal assignment rule, so an equal value keeps the base spelling.
- **New keys** from the overlay are appended with the overlay's comments and in the base document's inferred style.
- **Sequences** follow `YamlMergeOptions.Sequences`:
  - `Replace` is the default and matches Helm and `yq`'s `*` operator.
  - `Append` appends the overlay's items.
  - `MergeByKey(keyName)` merges items whose mapping has the same key value (for example containers by `name`) and appends the rest.
- **Null in the overlay:**
  - With `NullRemovesKey` (Helm's convention), an explicit null removes the base key.
  - By default a null is an ordinary value that replaces.
- **Kind conflicts:** when one side is a mapping and the other a scalar, or similar, the overlay wins, and the conflict is reported.
- **Aliases and anchors:** overlay values that land on an alias edit the anchored node, as any edit through an alias does, and are reported. Anchors and aliases from the overlay are copied with their names, and a name that collides with a base anchor is renamed and reported.
- **Placeholders:** an empty `{}` or `[]` in block context takes the overlay's block layout, because Helm charts use them as placeholders for values an overlay fills in.
- **Merge keys:** the overlay is read through its typed view, so the entries its `<<` keys provide count as its own. In the base, a key only a merge provides gets an entry of its own; merged mappings are never edited or flattened.
- **Aliases that cannot follow their anchor:** an overlay alias whose anchored node merged into an existing base node, or that would come before its anchor in the base, is written as a copy of the anchored node and reported (`yaml-merge-alias-expanded`), because the merged base node no longer holds the overlay's value.
- **Empty overlay documents** change nothing. With `NullRemovesKey`, a null for a key the base lacks adds nothing. A key that a null cannot remove (a merge provides it, or aliases refer into it) stays and is reported (`yaml-merge-key-not-removed`).
- **`YamlMergeReport`** lists every added, replaced and removed path, plus diagnostics for conflicts (`yaml-merge-kind-conflict`), alias edits (`yaml-aliased-node-edited`) and renamed anchors (`yaml-merge-anchor-renamed`). `Merge-OfficeDocument` writes it to the pipeline.
- **Qualification:** committed base and overlay pairs (`OfficeIMO.TestAssets/Yaml/Merge`) must change only the lines the report accounts for, load in YamlDotNet to the result of an independent Helm-rules merge of YamlDotNet-loaded data, load alike in SharpYaml, and merge to the same bytes every time.

### Writer

- `YamlWriter` is the forward-only emitter, modeled on `Utf8JsonWriter`. It has `WriteStartMapping`/`WriteEndMapping`, `WriteStartSequence`, `WriteKey`, `WriteScalar(value, style)`, `WriteComment`, `WriteAlias`, anchors, tags and document markers. It writes UTF-8 through the shared `PooledUtf8TextWriter` or to a `TextWriter`.
- Its output is deterministic: the same calls always produce the same bytes.
- `YamlWriterOptions` uses ruamel.yaml's three indentation numbers, which cover Kubernetes, Actions and compose styles:
  - mapping indent 2;
  - sequence indent 2 and dash offset 0, the compact `key:\n- a` most real files use (4 and 2 write `key:\n  - a`);
  - LF line endings (or CRLF).

  The writer always writes block style unless a flow collection is requested, never folds lines, and ends the stream with a line break.

  Each option exists because a real convention needs it (Kubernetes, Actions, compose and Helm differ in sequence indentation). Don't add an option without a corpus file that needs it.
- **Quoting safety:** by default a new string scalar is quoted when *any* supported schema (1.2 Core or 1.1) would read it as something other than a string. That covers `yes`, `no`, `on`, `off`, `y`, `n`, `~`, `null`, `1e3`, `0o17`, `012`, `1:20` and `.inf`. A file written by OfficeIMO therefore means the same thing to YAML 1.1 readers such as PyYAML, Ruby's Psych and older Go libraries. The 1.1 rules are the ones those readers implement (PyYAML and libyaml, go-yaml v2, Psych), not the type repository's looser float pattern, so versions such as `1.29.2` and addresses such as `10.0.0.1` stay plain. That is the most common real-world YAML bug, and we avoid it by default.

### Schemas and versions

- **Schemas:** `Failsafe`, `Json`, `Core` (the default) and `Yaml11`.
- `Yaml11` adds the 1.1 type rules:
  - boolean variants: `y`/`n`/`yes`/`no`/`on`/`off` in all three casings;
  - integers: sexagesimal (`1:20`), `0b` binary and `0`-prefixed octal;
  - timestamps;
  - `<<` merge keys.
- **Version selection:**
  - A `%YAML 1.1` directive selects `Yaml11` for that document part unless the caller forces a schema.
  - Without a directive, `Core` applies.
  - `YamlReaderOptions.Schema` overrides both, for ecosystems (Ansible, older Kubernetes tooling) that assume 1.1 without declaring it.
- **Typed values never rewrite raw text.** Reading `on` as a boolean under `Yaml11` and saving keeps `on`.
- **De-facto YAML by default; `Strict` for the specification.** Decision (Przemysław, 2026-11-10): the default reads what real files use and what libyaml (PyYAML, Ruby) and SharpYaml accept; YamlDotNet accepts the flow form but rejects most of the quoted form. `Strict` on `YamlReaderOptions` and `YamlLoadOptions` (default `false`) restores the specification's exact rules. One switch covers both kinds of leniency, because no consumer has needed one without the other; forcing `Schema = Yaml11` already gives merge keys with strict syntax. Without `Strict`:
  - Lines inside a flow collection, such as a closing `]` under its key, and quoted-scalar continuation lines may be indented no deeper than the enclosing block collection when spaces indent them. A tab in that indentation, and a document marker at column 0, still fail. The large corpus has 28 such files from Azure Pipelines, GitHub Actions, SAM, Argo, bitnami and istio; the default mode loads all of them with data equal to SharpYaml's (and YamlDotNet's on the 15 it accepts) and saves them back byte for byte.
  - `<<` is a merge key under `Core` as under `Yaml11`, in the typed view only and never flattened, so repeated `<<` keys (GitLab's `Deploy.gitlab-ci.yml`) are one merge list rather than duplicates. `Json` and `Failsafe` keep `<<` an ordinary key.
  - Each accepted indentation is a `yaml-lenient-indentation` warning with its position, once per outermost flow collection and once per quoted scalar. Merge keys resolved without a `%YAML 1.1` directive or a forced schema are one `yaml-merge-key-without-directive` information note per document part, so compose files are not noisy.
  - yaml-test-suite runs in both modes. Strict mode passes every case. The default mode also accepts exactly three error cases, listed with reasons in the reader conformance tests: `9C9N` and `VJP3/00` (flow lines at the key's indentation) and `QB6E` (quoted continuation at the key's indentation). `DK95/01` stays rejected because a tab indents its continuation.
  - Saving copies source text no edit owns, so accepted text round-trips byte for byte. The writer and every edit write valid indentation; a re-emitted region takes the valid form.
- **Duplicate keys** follow the same split. Decision (Przemysław, 2026-11-11): without `Strict`, the last key wins and each repeat is a `yaml-duplicate-key` warning with its position; with `Strict`, a repeat fails with `Yaml.Load.DuplicateKey`, as the specification requires. `YamlLoadOptions.DuplicateKeys` (`Error`, `FirstWins`, `LastWins`) overrides both. Repeated `<<` merge keys are not duplicates where merge keys apply.
  - The tree keeps every entry, so saving reproduces the file byte for byte. Lookups, paths, edits by key and the serializer resolve to the winning entry; `Remove` by key removes only that entry, so an earlier one becomes visible.
  - Last-wins matches what most readers do with real files. Measured on `a: 1
b: 2
a: 3`: PyYAML 6.0.3, go-yaml v2.4.0 and the YamlDotNet 18 deserializer return `a: 3` silently; YamlDotNet's `YamlStream` and `WithDuplicateKeyChecking`, SharpYaml 3.17 (unless `DuplicateKeyHandling` is set), ruamel.yaml 0.19 and go-yaml v3.0.1 fail. Jekyll and Hugo front matter also keep the last value. The large corpus has 19 files (23 repeats) that only load this way: Azure Pipelines templates, Kubernetes and Argo examples, SAM templates and a Prometheus test file. All of them load, save back byte for byte, and load to the same entries as SharpYaml's representation model.
  - The serializer has no diagnostics list, so it applies last-wins silently unless `YamlReaderOptions.Strict` is set.
- **Custom tags** (`!Ref`, `!Sub`, `!reference`) are preserved verbatim. A caller can register a resolver for a tag. Unknown tags never fail a load.

### Serializer

`YamlSerializer` converts between .NET objects and YAML. It follows `System.Text.Json`'s design, because .NET developers already know that API. The package README lists the supported types and scalar formats.

- **API**:
  - `Serialize`/`Deserialize` over strings, `Stream`, `TextReader`/`TextWriter`, `YamlReader`/`YamlWriter` and `YamlNode`, in generic and `Type` forms.
  - `YamlSerializerOptions` holds the naming policy, `DefaultIgnoreCondition`, `ReferenceHandling`, `ReaderOptions` (schema, strictness and every limit), `WriterOptions` and `TypeInfoResolver`. Options become read-only on first use and cache their metadata.
  - Overloads without metadata carry `RequiresUnreferencedCode`/`RequiresDynamicCode`. Overloads with `YamlTypeInfo<T>` or a `YamlSerializerContext` use none.
- **One converter engine, two metadata sources.**
  - The source generator emits calls to `YamlMetadataServices`: creators, constructor parameters, property getters and setters as delegates, and derived-type registrations.
  - Reflection metadata builds the same values through the same methods. Both apply one set of shape rules: `YamlTypeShapes` at run time, mirrored by the generator's parser. So both write the same YAML, which the round-trip tests check.
- **Generator packaging**:
  - The generator ships inside `OfficeIMO.Yaml` (`analyzers/dotnet/cs`), as `System.Text.Json`'s does, so its version cannot drift from the runtime.
  - `buildTransitive/OfficeIMO.Yaml.targets` turns reflection off for `PublishAot` and `PublishTrimmed` applications through the `OfficeIMO.Yaml.YamlSerializer.IsReflectionEnabledByDefault` switch (a `FeatureSwitchDefinition` on .NET 9+).
  - Without metadata, such an application gets an exception that names the context to add.
  - Shapes the serializer cannot handle are build errors at the context (`YAML001`–`YAML005`), never a runtime fallback.
  - Init-only setters and constructors of types with `required` members use `UnsafeAccessor` where the consumer's runtime has it, and reflection on .NET Framework and .NET Standard consumers, which have no trimming.
- **Streaming**:
  - Reading goes through `YamlReader` events and writing through `YamlWriter` calls. No tree is built, and object keys are matched without allocating where the runtime allows (`Dictionary.AlternateLookup` on .NET 9+).
  - `YamlNode` input is replayed as recorded events.
  - `SerializeToNode` builds a detached tree through an internal node-building mode of `YamlWriter`, so the same converters produce text and nodes.
- **Anchors and aliases**:
  - Reading records the events of an anchored node only while that node is read, so input without anchors costs nothing extra.
  - An alias replays the recorded events. Replayed nodes count against `MaxAliasExpansion`, and an alias inside its own anchor fails with `RecursiveAlias`.
  - Merge keys fill only the members a mapping does not set.
  - Writing has no reference handling by default and fails on a cycle. With `Preserve`, a pre-pass finds the objects reached more than once. Only those get anchors (`&id001`), and later references become aliases. Reading maps each alias back to the anchored instance, and objects register before their members are read, so cycles resolve.
- **Polymorphism** goes only through `[YamlDerivedType]` registrations on the base type:
  - The discriminator is a local tag (`!cat`, the YAML-native form other libraries map with tag registrations), or with `[YamlPolymorphic(DiscriminatorPropertyName = "kind")]` the first mapping key.
  - Reading finds a property discriminator anywhere: the usual first-key case streams, and otherwise the mapping is captured and replayed.
  - A tag never activates a type by itself; unknown discriminators fail with `UnknownDiscriminator`.
- **Scalars** are read by target type: the decoded text is parsed with the document schema's rules whatever its quoting. Strings are written with the writer's YAML 1.1-safe quoting. Timestamps are written plain in ISO 8601.
  - **Flags enums** are one comma-separated scalar (`Read, Write`). That is the spelling of `Enum.ToString`/`Enum.Parse`, which YamlDotNet and SharpYaml also read. It keeps zero and composite members as single values, and reading also accepts a sequence of names.
- **Untyped values**:
  - `object` reads as `Dictionary<string, object?>` and `List<object?>`. Scalars are resolved by the schema: `null`, `bool`, `long`, `BigInteger` beyond 64 bits, `double`, or `string`; timestamps stay strings.
  - The dictionary enumerates in file order because entries are only added. A PowerShell `ConvertFrom-OfficeYaml` builds its ordered hashtable or `PSCustomObject` by walking it, so an ordered dictionary type in the C# API would add a type without a consumer.
  - Writing dispatches on the runtime type and handles dictionaries (generic and `IDictionary`, which covers PowerShell's `Hashtable` and `[ordered]`), enumerables and scalars without metadata.
- **Errors**: input that does not fit throws `YamlException` with `TypeMismatch`, `MissingRequiredMember`, `DuplicateKey` or `UnknownDiscriminator`. The position points at the node, also inside a loaded document. The message names the path.
- **Dependencies**: there is no dependency on `System.Text.Json`. The attributes are the package's own.

## Resource and security limits

`YamlReaderOptions` and `YamlLoadOptions` carry these limits:

| Limit | Default | Purpose |
|---|---|---|
| `MaxDepth` | 256 | Bounds parser stacks |
| `MaxNodes` | 1,000,000 | Bounds model size |
| `MaxAliasExpansion` | 100,000 visited nodes | Stops "billion laughs" alias bombs in resolved views and the serializer |
| `MaxScalarLength` | 16 MiB | Bounds a single value |
| `MaxInputBytes` | none (caller-set) | Bounds whole input |

Exceeding a limit throws `YamlLimitExceededException` with the limit's name and the source position. Async loads honor cancellation, and the synchronous parser checks a `CancellationToken` at bounded intervals. The serializer never instantiates a type that was not registered or generated.

## Package boundaries

| Package | Depends on | Owns |
|---|---|---|
| `OfficeIMO.Yaml` | `OfficeIMO.Core` (diagnostics, shared IO) | Scanner, reader, writer, model, schemas, serializer, generator (built from `OfficeIMO.Yaml.SourceGeneration`, packed as an analyzer) |
| `OfficeIMO.Reader.Yaml` | `OfficeIMO.Reader.Core`, `OfficeIMO.Yaml` | Chunking and Markdown projection only; YamlDotNet removed |
| `OfficeIMO.Markdown` | `OfficeIMO.Core`, `OfficeIMO.Yaml` | Front matter parsed and written through `YamlDocument` |

- Targets: netstandard2.0, net8.0, net10.0 and net472, matching `OfficeIMO.CSV`. On netstandard2.0/net472 the only polyfill package is `System.Buffers`/`System.Memory`.
- Third-party YAML libraries (YamlDotNet, SharpYaml, others) appear only in `OfficeIMO.Yaml.Tests` and `OfficeIMO.Yaml.Benchmarks`, as oracles and comparison lanes.

## Performance

- The benchmark project exists before the engine (Y0). It covers small configuration files, a ~10 MB multi-document Kubernetes stream, deep nesting, big block scalars, flow-heavy input, anchor-heavy compose files and many small files.
- Measured operations:
  - parse-to-events;
  - load-to-model;
  - unchanged load-and-save;
  - single-value edit and save;
  - emit a new document;
  - serialize and deserialize.
- Every lane validates its output. YamlDotNet and SharpYaml v3 run the same workloads.
- Target: lead both reference libraries in elapsed time and allocation on parse-to-events and load-to-model, and stay competitive on serialization. Measurement stays opt-in (`Category=Performance`), outside pull-request correctness gates, per the repository's benchmark rules.
- Design rules that keep the target reachable:
  - Tokens and events are structs.
  - Nodes hold spans into the source, not copied strings.
  - Scalar decoding is lazy, and plain ASCII scalars without escapes never get an intermediate copy.
  - Mapping keys go through a small string cache, as `CsvParser.StringCache` does.
  - Buffers are pooled.
  - The model allocates nothing for trivia.
  - Shortcuts past the general path (the quoting decision by text shape, one-step serializer entries, short decimals, the plain-stop search) each have a test that compares them with the general path.

## Qualification

- **yaml-test-suite:**
  - events match `test.event` in both modes;
  - error cases fail in strict mode, and the default mode accepts only the listed de-facto cases;
  - loaded values equal `in.json`;
  - `in.yaml` round-trips byte for byte;
  - emitted output re-parses to the same events as `out.yaml`/`emit.yaml`.
- **Real-world corpus** with recorded provenance:
  - every file round-trips byte for byte;
  - scripted edits change only the edited lines (diff-checked);
  - output from OfficeIMO parses to the same data in YamlDotNet and SharpYaml, plus PyYAML/ruamel.yaml in an opt-in run, because independent readers accept what we write.
- **Hostile input:** limit tests at the parser boundary, plus an opt-in mutation run that must finish without crashes or hangs.
- **Targets:** net472/net8.0/net10.0 test lanes, an AotSmoke consumer for the reader, model and generated serializer, and a package consumer for the final package.

## Delivery order

1. **Y0:** test data, harness and comparison runners.
2. **Y1:** scanner, parser and `YamlReader`.
3. **Y2:** `YamlDocument` model, schemas and lossless save.
4. **Y5a:** `OfficeIMO.Reader.Yaml` on the engine. The first real consumer validates the read API early, and YamlDotNet is removed.
5. **Y3:** editing, `YamlWriter`, new-document emission, then Markdown front matter through the engine.
6. **Y4:** serializer and source generator.

Each milestone lands on `feature/yaml-engine` with its tests and benchmarks. The ready-for-review PR goes out when a usable milestone set (at least Y0–Y2 and Y5a) is complete.
