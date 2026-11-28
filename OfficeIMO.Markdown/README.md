# OfficeIMO.Markdown

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markdown)](https://www.nuget.org/packages/OfficeIMO.Markdown)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markdown)

Build, parse and render Markdown in .NET. `OfficeIMO.Markdown` has a fluent builder, a reader that parses Markdown into a typed document model with source spans, and an HTML renderer, all first-party with no external runtime dependencies.

## Install

```sh
dotnet add package OfficeIMO.Markdown
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Markdown/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.Markdown;

var document = MarkdownDocument.Create()
    .FrontMatter(new { title = "Release notes", tags = new[] { "docs", "reporting" } })
    .H1("Release notes")
    .P("Typed Markdown builder, reader and HTML renderer for .NET.")
    .Ul(list => list
        .Item("Build Markdown with a fluent API")
        .Item("Parse Markdown into typed blocks and inlines")
        .Item("Render HTML fragments or full documents"));

File.WriteAllText("Release-notes.md", document.ToMarkdown());
File.WriteAllText("Release-notes.html", document.ToHtml(new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document }));

MarkdownDocument parsed = MarkdownDocument.Parse(File.ReadAllText("Release-notes.md"));
Console.WriteLine($"{parsed.Blocks.Count} blocks");
```
<!-- /powerforge:code -->

The chaining methods are shortcuts for model methods such as `AddHeading`, `AddParagraph`, `AddList` and `AddTable`, which return the block they append.

## What it does

- **Build** documents with headings, paragraphs, lists, task lists, tables, code blocks, callouts, details, front matter, footnotes, tables of contents and semantic fenced blocks.
- **Parse** Markdown into a typed `MarkdownDocument` with blocks, inlines, anchors, source spans, YAML or TOML front matter and transform diagnostics, following CommonMark and GitHub Flavored Markdown.
- **Typed tables** from objects with `TableAuto`, with column selection, renaming and alignment.
- **Render HTML** fragments or full documents with profiles, CSS delivery, syntax highlighting, Mermaid, chart and math assets, and safe defaults.
- **AST snapshots** with stable block ids for editors, chat transcripts and document hosts.
- **Provenance**: `MarkdownProvenance.Inspect` and `Remove` find and remove C2PA carriers.
- **AOT-friendly** typed selectors and explicit overloads, with reflection helpers kept for convenience.

The [Markdown guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.markdown-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| PDF | [OfficeIMO.Markdown.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markdown.Pdf/README.md) |
| HTML to Markdown | [OfficeIMO.Markdown.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markdown.Html/README.md) |
| Word, both directions | [OfficeIMO.Word.Markdown](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.Markdown/README.md) |
| Live rendering in WebView and browser hosts | [OfficeIMO.MarkdownRenderer](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.MarkdownRenderer/README.md) |
| Semantic authoring for Office outputs | [OfficeIMO.Markup](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markup/README.md) |

## Limits

The [compatibility matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.markdown.compatibility-matrix.md) lists CommonMark and GFM coverage and the extensions, and [extension authoring](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.markdown.extension-authoring.md) shows how to add your own syntax.

## Learn more

- [OfficeIMO.Markdown on officeimo.com](https://officeimo.com/products/markdown/) and the [Markdown docs](https://officeimo.com/docs/markdown/).
- [Markdown guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.markdown-guide.md) and [examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Markdown).
- Command line: `officeimo convert notes.md notes.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core`, plus `OfficeIMO.Yaml` and `OfficeIMO.Toml` for front matter. No external dependencies.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |
| Remove | 1 | 0 | 0 | 0 | 0 | 0 |
| Convert | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
