# OfficeIMO.Reader.Markdown

Markdown support for `OfficeIMO.Reader.Core`, backed by the typed `OfficeIMO.Markdown` parser.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddMarkdownHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("README.md");
foreach (ReaderChunk chunk in result.Chunks) {
    Console.WriteLine($"{chunk.Id} [{chunk.Location.HeadingPath}] lines {chunk.Location.StartLine}-{chunk.Location.EndLine}");
}
```

The adapter preserves source spans and heading paths and projects Markdown tables and supported visual fences into Reader's structured models. `OfficeIMO.Reader.All` registers this handler.

## Registration

| Item | Value |
|---|---|
| Handler id | `officeimo.reader.markdown` |
| Extensions | `.md`, `.markdown`, `.mdown`, `.mkd`, `.mdx` |
| Reader kind | `ReaderInputKind.Markdown` |
| Default input limit | 64 MiB when `ReaderOptions.MaxInputBytes` is not set |
| Output | Deterministic |

`AddMarkdownHandler(options, replaceExisting)` snapshots the options at registration; later changes to the same `ReaderMarkdownOptions` instance do not affect the built reader.

## Options

| `ReaderMarkdownOptions` | Default | Effect |
|---|---|---|
| `ChunkByHeadings` | `true` | Starts a new chunk at every heading. When `false`, blocks are packed into chunks by size only. |
| `ParserOptions` | `null` (default `MarkdownLoadOptions`, the OfficeIMO profile) | Parser profile and switches. Use `MarkdownLoadOptions.CreateCommonMarkProfile()`, `CreateGitHubFlavoredMarkdownProfile()` or `CreatePortableProfile()` for standards-only parsing. The value is cloned at registration. |

Shared `ReaderOptions` also apply:

| `ReaderOptions` | Default | Effect on Markdown |
|---|---|---|
| `MaxChars` | `8000` | Target chunk size. Values below 256 are treated as 256. |
| `MaxTableRows` | `200` | Rows kept per projected table. |
| `MaxInputBytes` | `null` (handler default) | Input size limit enforced before parsing. |

## What each chunk contains

- `Id` in the form `markdown:<file name>:<index>`, for example `markdown:README.md:0002`.
- `Text` and `Markdown` with the normalized Markdown of the blocks in the chunk.
- `Location.StartLine` and `EndLine` from the parser's source spans, `HeadingPath` such as `Guide > Install`, `HeadingSlug`, `SourceBlockKind` (`heading`, `paragraph`, `table`, `code`, ...) and `BlockAnchor`.
- `Tables` for GFM tables and `ix-dataview` JSON fences (`columns` + `rows`, or `records`), with column profiles.
- `Visuals` for fenced `mermaid`, `chart`/`ix-chart` and `network`/`ix-network`/`visnetwork` blocks. Visual content is carried as source text with a payload hash; nothing is rendered or executed.

## Limits and warnings

- Blocks are never cut in the middle. A single block larger than `MaxChars` becomes its own chunk with the warning `A single markdown block exceeded MaxChars and was preserved as one chunk.`
- Text is decoded as UTF-8, honoring a byte-order mark when present.
- Links and images are not fetched.

## Example output

For this input:

```markdown
# Guide

Intro paragraph.

## Install

| Package | Command |
|---|---|
| Reader | `dotnet add package OfficeIMO.Reader.All` |
```

the default options produce two chunks:

| `Id` | `HeadingPath` | Lines | Content |
|---|---|---|---|
| `markdown:guide.md:0000` | `Guide` | 1-3 | Heading and intro paragraph |
| `markdown:guide.md:0001` | `Guide > Install` | 5-9 | Heading and table; `Tables` holds one table with columns `Package` and `Command` |

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Markdown`.
