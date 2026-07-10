# OfficeIMO.AsciiDoc

`OfficeIMO.AsciiDoc` is a dependency-free, source-preserving AsciiDoc parser, semantic model, writer, and explicit preprocessing engine.

The experimental Phase 1 profile covers headings and metadata, typed inline formatting and references, ordered/unordered/description/compound lists, admonitions, variable-length delimited blocks, structured PSV/CSV/TSV/DSV tables, document attributes, substitution plans, conditionals, safe includes with line/tag selection, and caller-registered directives. Unsupported constructs remain in the lossless source tree, and unchanged input writes back character-for-character.

```csharp
using OfficeIMO.AsciiDoc;

AsciiDocParseResult result = AsciiDocDocument.Parse(source);
AsciiDocDocument document = result.Document;

AsciiDocHeading title = document.Blocks.OfType<AsciiDocHeading>().First();
title.Title = "Updated title";

string updated = document.ToAsciiDoc(AsciiDocWriterMode.Preserve);
```

Processing is opt-in and separate from parsing:

```csharp
AsciiDocProcessingResult processed = AsciiDocProcessor.Process(
    source,
    new AsciiDocProcessorOptions {
        // Null keeps includes disabled. A resolver must be supplied explicitly.
        IncludeResolver = null
    });
```

## Boundaries

- The native parser and writer do not use external packages or executables.
- Parsing never reads includes or executes registered directives. Only the explicit processor can do so, under caller-supplied policy and hard limits.
- The built-in include resolver is root-confined and rejects remote, absolute, traversal, and symbolic-link escape by default.
- The implementation does not claim every Asciidoctor substitution, macro, extension, nested table-cell, or rendering behavior.
- Preserve mode reuses original source for every unchanged subtree.
- Canonical mode emits stable OfficeIMO formatting for recognized semantic nodes.
- Character source, whitespace, and line endings are lossless. Original file encoding and BOM bytes are not retained by `Load`/`Save`.

See the [AsciiDoc support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.asciidoc-support-matrix.md) for the feature-level contract.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
