# OfficeIMO.AsciiDoc

`OfficeIMO.AsciiDoc` is a dependency-free, source-preserving AsciiDoc parser, semantic model, writer, and explicit preprocessing engine.

The bounded profile covers headings and metadata, typed inline formatting and references, footnotes, bibliography anchors, ordered/unordered/description/callout/compound lists, admonitions, variable-length delimited blocks, structured PSV/CSV/TSV/DSV tables, document attributes, substitution plans, conditionals, safe includes with line/tag selection, and caller-registered directives. Unsupported constructs remain in the lossless source tree, and unchanged input writes back character-for-character.

```csharp
using OfficeIMO.AsciiDoc;

AsciiDocDocument document = AsciiDocDocument.Parse(source);
AsciiDocParseResult result = AsciiDocDocument.ParseResult(source);

AsciiDocHeading title = document.Blocks.OfType<AsciiDocHeading>().First();
title.Title = "Updated title";

string updated = document.ToAsciiDoc(AsciiDocWriterMode.Preserve);
```

## Create and edit structure

```csharp
AsciiDocDocument authored = AsciiDocDocument.Create()
    .AddHeading(1, "Getting started")
    .AddParagraph("A *formatted* paragraph.")
    .Add("* First item\n* Second item");

AsciiDocBlock paragraph = authored.Blocks.OfType<AsciiDocParagraph>().First();
authored.Move(paragraph, 0);
authored.Remove(paragraph);
```

`Add` and `Insert` accept native AsciiDoc fragments. `AddParagraph` accepts one nonempty paragraph with native inline syntax and escapes block-start markers; use `Add` for multiple paragraphs. Move/remove operations keep a block's bound metadata and list attachments together and reject insertion points that would split those units.

Compound blocks expose an editable `Body`. Table cells with style `a` expose a parsed AsciiDoc `Body`; ordinary cells expose `Inlines`. Scalar and nested edits share the effective content used by native writing and adapters.

`AsciiDocReferenceCatalog.Create(document)` collects explicit anchors, generated section targets, footnotes, and reference diagnostics. `GetBlockId(block)` returns the catalog's ID without changing the source. Generated IDs honor source-order `sectids`, `idprefix`, and `idseparator` attributes, visible title formatting and labeled URLs, Unicode, and duplicate suffixes. Explicit IDs take precedence. Unsupported title substitutions and empty IDs have diagnostics; supply an explicit ID when those titles need processor-independent links. Rebuild catalogs after editing titles, anchors, attributes, footnotes, or document order.

`AsciiDocCalloutCatalog.Create(document)` correlates callout explanations with markers in the preceding listing or literal block.

Processing is opt-in and separate from parsing:

```csharp
AsciiDocProcessingResult processed = AsciiDocProcessor.Process(
    source,
    new AsciiDocProcessorOptions {
        // Null keeps includes disabled. A resolver must be supplied explicitly.
        IncludeResolver = null
});
```

Select a named parsing profile explicitly when the caller must pin the semantic contract. Both profiles are lossless; `OfficeIMO` exposes typed common constructs, while `PreserveOnly` identifies a preservation-oriented pipeline. Neither profile reads includes or runs extensions during parsing.

```csharp
AsciiDocDocument preserved = AsciiDocDocument.Parse(
    source,
    AsciiDocParseOptions.CreateProfile(AsciiDocDocumentProfile.PreserveOnly));

AsciiDocProcessorOptions bounded = AsciiDocProcessorOptions.CreateProfile(
    AsciiDocDocumentProfile.OfficeIMO);
bounded.IncludeResolver = new AsciiDocRootedFileIncludeResolver(contentRoot);
```

The rooted resolver denies remote, absolute, traversal, and symbolic-link escape targets by default. Attributes, conditionals, line/tag include selection, and caller-registered directives are processed only by `AsciiDocProcessor`, under its include depth/count/character and extension invocation limits.

`processed.SourceMap` maps offsets in `ProcessedSource` back to the original root/include file and line, including nested includes and line/tag selection. Processing diagnostics retain original line numbers. Use the map to locate parser or conversion diagnostics from the expanded document:

```csharp
foreach (AsciiDocDiagnostic diagnostic in processed.Document.Diagnostics) {
    int offset = diagnostic.Span.Start.Offset;
    AsciiDocSourceMapping? origin = processed.SourceMap.Find(offset);
    if (origin != null) {
        string? file = origin.SourceName;
        AsciiDocSourcePosition position = origin.GetSourcePosition(offset);
    }
}
```

Unchanged line slices have exact positions. Heading-level offsets, inline conditional replacements, generated extension output, and inserted line endings refer to their producing source range with `IsExact = false`. The map describes the processing snapshot; subsequent edits do not update it.

## Current limits

- The native parser and writer do not use external packages or executables.
- Parsing never reads includes or executes registered directives. Only the explicit processor can do so, under caller-supplied policy and hard limits.
- The built-in include resolver is root-confined and rejects remote, absolute, traversal, and symbolic-link escape by default.
- The implementation does not claim every Asciidoctor substitution, macro, extension, advanced table layout, or rendering behavior.
- Preserve mode reuses original source for every unchanged subtree.
- Canonical mode emits stable OfficeIMO formatting for recognized semantic nodes.
- Character source, whitespace, and line endings are lossless. Original file encoding and BOM bytes are not retained by `Load`/`Save`.
- Path saves stage complete output beside the destination and require atomic publication. If the filesystem cannot atomically replace an existing file, the save fails and preserves that file.

Path and stream loads share bounded, BOM-aware decoding. Caller-selected encodings take precedence: a matching preamble is removed, and bytes that resemble another encoding's BOM are decoded with the selected encoding. Parsing, loading, and explicit processing accept cancellation; input, inline, table, expansion, include, and output limits are enforced by their options.

See the [AsciiDoc support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.asciidoc-support-matrix.md) for the feature-level contract.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.

## Dependency footprint

- **External:** None; no Asciidoctor process or parser package.
- **OfficeIMO:** `OfficeIMO.Core`. Parsing, source preservation, processing limits, and writing are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.AsciiDoc` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
