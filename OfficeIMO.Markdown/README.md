# OfficeIMO.Markdown - Markdown builder, reader, and HTML renderer

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Markdown)](https://www.nuget.org/packages/OfficeIMO.Markdown)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Markdown?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Markdown)

`OfficeIMO.Markdown` is the core Markdown package in the OfficeIMO family. It builds Markdown, parses Markdown into a typed document model, projects native AST snapshots for hosts, and renders HTML without external runtime dependencies.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys).

## Install

```powershell
dotnet add package OfficeIMO.Markdown
```

## Quick start

```csharp
using OfficeIMO.Markdown;

var document = MarkdownDoc.Create()
    .FrontMatter(new { title = "OfficeIMO.Markdown", tags = new[] { "docs", "reporting" } })
    .H1("OfficeIMO.Markdown")
    .P("Typed Markdown builder, reader, and HTML renderer for .NET.")
    .Ul(list => list
        .Item("Build Markdown with a fluent API")
        .Item("Parse Markdown into typed blocks and inlines")
        .Item("Render HTML fragments or full documents"))
    .Code("powershell", "dotnet add package OfficeIMO.Markdown");

string markdown = document.ToMarkdown();
string htmlFragment = document.ToHtmlFragment();
string htmlDocument = document.ToHtmlDocument();
```

## What it does

- Builds Markdown documents with headings, paragraphs, lists, task lists, tables, code blocks, callouts, details, front matter, footnotes, table-of-contents helpers, and semantic fenced blocks.
- Parses Markdown into a typed `MarkdownDoc` object model with headings, blocks, inlines, anchors, source spans, front matter, and transform diagnostics.
- Renders HTML fragments or full HTML documents with configurable profiles, CSS delivery, Prism options, Mermaid/chart/math shell assets, and safe defaults.
- Provides native AST projection for editor, transcript, chat, and document hosts that need stable block ids and DTO-style snapshots.
- Supports AOT-sensitive code paths through typed selectors and explicit overloads while leaving reflection helpers available for convenience scenarios.

## Common tasks

### Tables from typed objects

```csharp
var results = new[] {
    new { Name = "Alice", Role = "Dev", Score = 98.5 },
    new { Name = "Bob", Role = "Ops", Score = 91.0 }
};

var document = MarkdownDoc.Create()
    .H2("Team")
    .Table(table => table.FromSequence(results,
        ("Name", item => item.Name),
        ("Role", item => item.Role),
        ("Score", item => item.Score)));
```

### Parse and inspect headings

```csharp
var parsed = MarkdownReader.Parse(File.ReadAllText("README.md"));

foreach (var heading in parsed.GetHeadingInfos()) {
    Console.WriteLine($"{heading.Level}: {heading.Text} -> {heading.Anchor}");
}
```

### Native AST snapshot

```csharp
MarkdownNativeDocument native = MarkdownNativeDocument.Parse("""
# Investigation

See **CPU** in [dashboard](https://example.com).
""");

var snapshot = native.ToSnapshot();
```

### Front matter, task lists, callouts, details, and TOC

```csharp
var document = MarkdownDoc.Create()
    .FrontMatter(new {
        title = "Deployment checklist",
        owner = "Platform",
        tags = new[] { "runbook", "deployment" }
    })
    .H1("Deployment checklist")
    .Toc(options => {
        options.Title = "Contents";
        options.MinLevel = 2;
        options.MaxLevel = 3;
    })
    .H2("Pre-flight")
    .Ul(list => list
        .ItemTask("Packages published", done: true)
        .ItemTask("Change approved")
        .ItemTask("Rollback plan attached"))
    .Callout("warning", "Freeze window", "Do not deploy outside the approved window.")
    .Details("Rollback commands", body => body
        .Code("powershell", "dotnet nuget locals all --clear"));

string markdown = document.ToMarkdown();
```

### HTML fragments, full documents, and parts

```csharp
var document = MarkdownDoc.Create()
    .H1("Status")
    .P(p => p.Text("Generated ").Bold("today").Text(" for the portal."))
    .TableAuto(table => table
        .Headers("Service", "Status", "Latency")
        .Row("API", "Green", "42")
        .Row("Jobs", "Yellow", "210"));

string fragment = document.ToHtmlFragment();
string fullPage = document.ToHtmlDocument();
HtmlRenderParts parts = document.ToHtmlParts();
```

### Parse, inspect, and rewrite a document

```csharp
var parsed = MarkdownReader.Parse(File.ReadAllText("README.md"),
    MarkdownReaderOptions.CreateOfficeIMOProfile());

if (!parsed.HasHeading("Install")) {
    parsed.H2("Install")
          .Code("powershell", "dotnet add package OfficeIMO.Markdown");
}

foreach (var heading in parsed.GetHeadingInfos()) {
    Console.WriteLine($"{heading.Level}: {heading.Text} -> #{heading.Anchor}");
}

File.WriteAllText("README.normalized.md", parsed.ToMarkdown());
```

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Markdown.Html](../OfficeIMO.Markdown.Html/README.md) | HTML to Markdown document conversion. |
| [OfficeIMO.Markdown.Pdf](../OfficeIMO.Markdown.Pdf/README.md) | Markdown to PDF through `OfficeIMO.Pdf`. |
| [OfficeIMO.MarkdownRenderer](../OfficeIMO.MarkdownRenderer/README.md) | WebView/browser-friendly shell rendering and incremental updates. |
| [OfficeIMO.Word.Markdown](../OfficeIMO.Word.Markdown/README.md) | Word to/from Markdown conversion. |
| [OfficeIMO.Markup](../OfficeIMO.Markup/README.md) | Markdown-inspired semantic authoring for OfficeIMO document outputs. |

## Boundaries

- `OfficeIMO.Markdown` owns the Markdown model, parser, writer, and HTML renderer.
- PDF output belongs in `OfficeIMO.Markdown.Pdf`.
- HTML ingestion belongs in `OfficeIMO.Markdown.Html`.
- Host shell behavior belongs in `OfficeIMO.MarkdownRenderer` and host-specific plug-ins.

## Deeper docs

- [Correctness roadmap](../Docs/officeimo.markdown.correctness-roadmap.md)
- [Correctness backlog](../Docs/officeimo.markdown.correctness-backlog.md)
- [Extension authoring](../Docs/officeimo.markdown.extension-authoring.md)
- [Benchmarks](../OfficeIMO.Markdown.Benchmarks/README.md)

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None; no Markdig or other Markdown parser.
- **OfficeIMO:** `OfficeIMO.Drawing`. The AST, parser, builder, transformations, and HTML renderer are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
