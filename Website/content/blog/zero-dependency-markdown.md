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

string input = File.ReadAllText("README.md");
MarkdownDoc doc = MarkdownReader.Parse(input);

foreach (var block in doc.TopLevelBlocks)
{
    switch (block)
    {
        case HeadingBlock h:
            Console.WriteLine($"H{h.Level}: {h.Text}");
            break;
        case ParagraphBlock p:
            Console.WriteLine($"Paragraph with {p.Inlines.Nodes.Count} inline nodes");
            break;
        case CodeBlock cb:
            Console.WriteLine($"Code ({cb.Language}): {cb.Content.Length} chars");
            break;
        case TableBlock t:
            Console.WriteLine($"Table: {t.Rows.Count} rows");
            break;
    }
}
```

Pattern matching on sealed types gives you exhaustiveness checking at compile time. If we add a new block type in a future release, your switch expression will produce a compiler warning, not a silent bug.

## The Builder API

Creating Markdown is just as clean:

```csharp
using OfficeIMO.Markdown;

var markdown = MarkdownDoc.Create()
    .H2("Release Notes")
    .P("Version 1.4.0 ships with the following changes:")
    .Ul(ul => ul
        .Item("Parallel AutoFit in OfficeIMO.Excel")
        .Item("Cross-platform PDF conversion")
        .Item("Improved table border handling"))
    .Code("csharp", "var doc = WordDocument.Create(\"demo.docx\");")
    .ToMarkdown();
```

The builder handles blank-line separation, fence formatting, and list indentation so you never have to think about whitespace rules.

## Transformation Pipeline

Because the AST is mutable, you can write transformation passes:

```csharp
var doc = MarkdownReader.Parse(File.ReadAllText("README.md"));

// Bump all headings down one level
foreach (var heading in doc.DescendantHeadings())
{
    heading.Level = Math.Min(heading.Level + 1, 6);
}

// Remove code blocks in a specific language
doc.TopLevelBlocks.RemoveAll(b => b is CodeBlock cb && cb.Language == "diff");

string output = doc.ToMarkdown();
```

This is the kind of structural manipulation that is awkward with a string-based or HTML-centric library.

## AOT and Trimming

Because OfficeIMO.Markdown keeps its dependency surface small and avoids the heavier Open XML document stack, it is one of the lower-risk packages in the repo for trimming- or AOT-sensitive deployments. Even so, you should still validate your own publish configuration rather than treating size or startup numbers as universal guarantees.

## When to Use Markdig Instead

If you need full CommonMark compliance, GFM autolinks, or an extensive plugin ecosystem, Markdig is the better choice. OfficeIMO.Markdown targets the practical subset of Markdown used in documentation and reports: headings, paragraphs, lists, tables, code blocks, and inline emphasis. For that scope, it is smaller, faster, and easier to reason about.
