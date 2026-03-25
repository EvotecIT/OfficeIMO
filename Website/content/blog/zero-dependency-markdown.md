---
title: "Zero-Dependency Markdown: Why We Built OfficeIMO.Markdown"
description: "A deep dive into the design philosophy behind OfficeIMO.Markdown, its typed AST, builder API, and why we chose zero external dependencies over Markdig."
date: 2025-08-01
tags: [markdown, design, aot]
categories: [Deep Dive]
author: "Przemyslaw Klys"
---

Markdown is everywhere: README files, documentation sites, CMS content, chat messages. Yet most .NET Markdown libraries either produce raw HTML strings or depend on a large pipeline of extensions. When we needed Markdown support inside OfficeIMO for document-to-Markdown conversion, we decided to build our own. Here is why.

## Why Not Markdig?

Markdig is an excellent library. It is fast, well-tested, and extensible. But it was designed primarily as a Markdown-to-HTML renderer. Its AST is optimised for rendering, not for programmatic inspection or round-trip transformation. It also brings a dependency graph that, while small, conflicts with our goal of NativeAOT readiness and single-assembly deployment.

We needed a Markdown layer that:

1. **Parses into a strongly typed AST** where every node is a concrete C# type.
2. **Builds Markdown programmatically** without string concatenation.
3. **Has zero external dependencies** so it can be trimmed and AOT-compiled.
4. **Supports round-trip fidelity** so you can parse a document, transform it, and emit Markdown that preserves the original formatting choices.

## The Typed AST

Every Markdown construct maps to a sealed class:

```csharp
using OfficeIMO.Markdown;

MarkdownDocument doc = MarkdownParser.Parse(input);

foreach (MarkdownBlock block in doc.Blocks)
{
    switch (block)
    {
        case MarkdownHeading h:
            Console.WriteLine($"H{h.Level}: {h.InlineText}");
            break;
        case MarkdownParagraph p:
            Console.WriteLine($"Paragraph: {p.InlineText}");
            break;
        case MarkdownCodeBlock cb:
            Console.WriteLine($"Code ({cb.Language}): {cb.Code.Length} chars");
            break;
        case MarkdownTable t:
            Console.WriteLine($"Table: {t.Rows.Count} rows");
            break;
    }
}
```

Pattern matching on sealed types gives you exhaustiveness checking at compile time. If we add a new block type in a future release, your switch expression will produce a compiler warning, not a silent bug.

## The Builder API

Creating Markdown is just as clean:

```csharp
var builder = new MarkdownBuilder();

builder.AddHeading("Release Notes", level: 2);
builder.AddParagraph("Version 1.4.0 ships with the following changes:");
builder.AddUnorderedList(new[]
{
    "Parallel AutoFit in OfficeIMO.Excel",
    "Cross-platform PDF conversion",
    "Improved table border handling"
});
builder.AddCodeBlock("csharp", "var doc = WordDocument.Create(\"demo.docx\");");

string markdown = builder.ToString();
```

The builder handles blank-line separation, fence formatting, and list indentation so you never have to think about whitespace rules.

## Transformation Pipeline

Because the AST is mutable, you can write transformation passes:

```csharp
var doc = MarkdownParser.Parse(File.ReadAllText("README.md"));

// Bump all headings down one level
foreach (var heading in doc.Blocks.OfType<MarkdownHeading>())
{
    heading.Level = Math.Min(heading.Level + 1, 6);
}

// Remove code blocks in a specific language
doc.Blocks.RemoveAll(b => b is MarkdownCodeBlock cb && cb.Language == "diff");

string output = doc.ToMarkdown();
```

This is the kind of structural manipulation that is awkward with a string-based or HTML-centric library.

## AOT and Trimming

Because OfficeIMO.Markdown uses no reflection, no `System.Linq.Expressions`, and no runtime code generation, it is fully compatible with `PublishTrimmed` and `PublishAot`. The entire assembly trims to under 80 KB.

## When to Use Markdig Instead

If you need full CommonMark compliance, GFM autolinks, or an extensive plugin ecosystem, Markdig is the better choice. OfficeIMO.Markdown targets the practical subset of Markdown used in documentation and reports: headings, paragraphs, lists, tables, code blocks, and inline emphasis. For that scope, it is smaller, faster, and easier to reason about.
