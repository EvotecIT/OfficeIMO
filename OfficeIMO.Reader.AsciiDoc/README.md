# OfficeIMO.Reader.AsciiDoc

`OfficeIMO.Reader.AsciiDoc` provides a modular `.adoc`, `.asciidoc`, and `.asc` handler backed by the dependency-free, lossless `OfficeIMO.AsciiDoc` engine.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.AsciiDoc;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAsciiDocHandler()
    .Build();
IReadOnlyList<ReaderChunk> chunks = reader.Enumerate("guide.adoc").ToList();
```

The handler emits deterministic block-aware chunks with source lines, heading paths, typed Markdown projections, structured table content, compound-list ownership, and parser/conversion warnings. Whole-document mode includes attached list content once, without duplicating it as a top-level block. Parsing and writing remain owned by `OfficeIMO.AsciiDoc`; this package only adapts the native model to Reader contracts.

`OfficeIMO.Reader.All` registers this handler.

## Registration

| Item | Value |
|---|---|
| Handler id | `officeimo.reader.asciidoc` |
| Extensions | `.adoc`, `.asciidoc`, `.asc` |
| Reader kind | `ReaderInputKind.AsciiDoc` |
| Input limit | `ReaderOptions.MaxInputBytes` when set, plus the parser's `MaxInputLength` |

`AddAsciiDocHandler(options, replaceExisting: true)` snapshots the options at registration.

## Options

| `ReaderAsciiDocOptions` | Default | Effect |
|---|---|---|
| `ChunkByBlock` | `true` | One chunk per top-level block. When `false`, the whole document is one projection split only by `ReaderOptions.MaxChars`. |
| `IncludeComments` | `false` | Emits line and block comments as chunks. |
| `IncludeAttributes` | `false` | Emits attribute entries such as `:toc:` as chunks. |
| `IncludeDiagnostics` | `true` | Adds parser and Markdown-conversion diagnostics to chunk `Warnings` as `CODE: message`. |
| `ParseOptions` | `new AsciiDocParseOptions()` | Parser profile and limits (see below). |
| `MarkdownOptions` | `new AsciiDocMarkdownExportOptions()` | Controls the Markdown projection, for example attribute expansion and preserving unsupported constructs as source. |

| `AsciiDocParseOptions` | Default |
|---|---|
| `Profile` | `AsciiDocDocumentProfile.OfficeIMO` |
| `MaxInputLength` | 64 MiB of characters |
| `MaxBlockCount` | 1,000,000 |
| `MaxInlineNestingDepth` | 64 |
| `MaxInlineNodeCount` | 1,000,000 |
| `MaxTableCellCount` | 1,000,000 |
| `MaxTableColumnCount` | 10,000 |
| `MaxTableSpan` | 10,000 |

Exceeding a parser limit fails the read rather than silently returning partial content.

## Output

- **Chunk ids:** `asciidoc-<block>` in block mode (`asciidoc-0`, `asciidoc-2`, ...); a block split by `MaxChars` continues as `asciidoc-4-part-2`.
- **Location:** `StartLine`/`EndLine` from the lossless source spans, `HeadingPath` such as `Guide > Install`, `BlockAnchor` `asciidoc-block-<n>`, and `SourceBlockKind`: `heading`, `paragraph`, `ordered-list`, `unordered-list`, `callout-list`, `description-list`, `admonition`, `table`, `delimited-<kind>` (for example `delimited-listing`) or `block-macro`.
- **Text and Markdown:** `Text` is plain text with attributes resolved; `Markdown` is the typed `OfficeIMO.AsciiDoc.Markdown` projection, except when `MaxChars` splits a chunk, in block mode or in document mode (`ChunkByBlock = false`): each part then carries its plain text as `Markdown`, so structure such as listing fences is not kept.
- **Warnings:** diagnostics that fall inside the block's span, plus `AsciiDoc content was split due to ReaderOptions.MaxChars.` for split blocks.

## Example output

```asciidoc
= Guide

== Install

NOTE: Requires .NET 8 or later.
```

produces three chunks in block mode:

| `Id` | `SourceBlockKind` | `HeadingPath` | Lines | `Markdown` |
|---|---|---|---|---|
| `asciidoc-0` | `heading` | `Guide` | 1 | `# Guide` |
| `asciidoc-2` | `heading` | `Guide > Install` | 3 | `## Install {#_install}` |
| `asciidoc-4` | `admonition` | `Guide > Install` | 5 | `> [!NOTE]` callout with `Requires .NET 8 or later.` |
Chunk ids follow the parser's source block index rather than a running count, so they stay stable when comments or attribute entries are filtered out.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core`, `OfficeIMO.AsciiDoc`, and `OfficeIMO.AsciiDoc.Markdown`; parsing stays in the native AsciiDoc package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
