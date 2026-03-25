---
title: "OfficeIMO.Markdown"
description: "Build, parse, and render Markdown with a typed AST and fluent API. Multiple reader profiles, HTML rendering, and zero dependencies."
layout: product
product_color: "#7c3aed"
install: "dotnet add package OfficeIMO.Markdown"
nuget: "OfficeIMO.Markdown"
docs_url: "/docs/markdown/"
api_url: "/api/markdown/"
---

## Why OfficeIMO.Markdown?

OfficeIMO.Markdown is a purpose-built Markdown engine for .NET. It gives you a strongly typed AST, a fluent builder for document construction, and multiple reader profiles so you can parse CommonMark, GFM, or OfficeIMO-flavored Markdown with one library. Every node carries source span information, making it ideal for tooling, linters, and editor integrations.

## Features

- **Fluent document builder** -- construct Markdown documents programmatically with a chainable API
- **Typed block & inline AST model** -- headings, paragraphs, lists, tables, code blocks, emphasis, links, and images as strongly typed objects
- **Source spans for every node** -- line, column, and offset tracking for diagnostics and editor support
- **Multiple reader profiles** -- OfficeIMO, CommonMark, GFM, and Portable profiles with configurable strictness
- **HTML rendering** -- emit fragment or full-page HTML with customizable templates
- **Front matter support** -- parse YAML and TOML front matter into typed dictionaries
- **TOC helpers & callouts** -- generate table of contents from headings and render note/warning/tip callouts
- **Tables from objects** -- build Markdown tables directly from collections with column selectors
- **Input normalization presets** -- normalize line endings, whitespace, and encoding before parsing
- **Post-parse document transforms** -- rewrite, filter, or augment the AST after parsing
- **Extension API** -- register custom block and inline parsers for domain-specific syntax
- **Zero external dependencies** -- ships as a single assembly with no third-party references

## Quick start

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDoc.Create()
    .FrontMatter(new { title = "Release Notes", tags = new[] { "docs", "release" } })
    .H1("Release Notes")
    .P("Version 3.0 introduces several improvements.")
    .H2("New Features")
    .Ul(ul => ul
        .Item("Fluent document builder")
        .Item("Typed AST and traversal helpers")
        .Item("HTML fragment and full-page rendering"))
    .H2("Performance")
    .Table(t => t
        .Headers("Benchmark", "v2.4", "v3.0")
        .Row("Parse 10K lines", "48 ms", "21 ms")
        .Row("Render to HTML", "35 ms", "14 ms")
        .Row("Round-trip fidelity", "97%", "100%")
        .AlignNumericRight())
    .Code("powershell", "dotnet add package OfficeIMO.Markdown");

string markdown = doc.ToMarkdown();
string html = doc.ToHtmlFragment();

var parsed = MarkdownReader.Parse(
    markdown,
    MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());

// Inspect the AST
foreach (var block in parsed.DescendantsAndSelf())
{
    Console.WriteLine(block.GetType().Name);
}
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Markdown runs on Windows, Linux, and macOS. It is AOT-compatible and trimming-safe.

## Related guides

| Guide | Description |
|-------|-------------|
| [Markdown documentation](/docs/markdown/) | Start with the package overview and document model. |
| [Builder API](/docs/markdown/builder/) | Compose documents fluently with headings, tables, and callouts. |
| [PSWriteOffice Markdown cmdlets](/docs/pswriteoffice/markdown/) | Generate Markdown from PowerShell objects and scripts. |
| [Word to Markdown](/docs/converters/word-markdown/) | Convert between Word documents and Markdown workflows. |
