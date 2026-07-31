---
title: "Zero-Dependency Markdown: Inside OfficeIMO.Markdown"
description: "A deep dive into OfficeIMO.Markdown's typed model, builder API, source-aware editing, and dependency-light deployment for .NET applications."
date: 2025-08-01
tags: [markdown, design, aot]
categories: [Deep Dive]
author: "Przemyslaw Klys"
---

Markdown is everywhere: README files, documentation sites, CMS content, and chat messages. OfficeIMO needs Markdown as an editable document model as well as an interchange format, so `OfficeIMO.Markdown` owns parsing, a typed semantic tree, source-aware editing, writing, and HTML projection in one first-party package.

## Why a first-party document model?

The package is designed around four current contracts:

1. **Parses into a strongly typed AST** where every node is a concrete C# type.
2. **Builds Markdown programmatically** without string concatenation.
3. **Keeps third-party dependencies out of the core model** and includes a NativeAOT publish-and-run smoke for fluent rendering.
4. **Supports round-trip fidelity** so you can parse a document, transform it, and emit Markdown that preserves the original formatting choices.

## The Typed AST

Supported Markdown constructs map to typed public block and inline classes:

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

Pattern matching makes structural inspection explicit. Applications should retain a default branch because optional extensions can introduce additional block or inline types.

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

OfficeIMO.Markdown's checked-in NativeAOT executable composes and renders a document, and CI republishes that scenario on Linux. That is a concrete baseline, not a blanket guarantee for every extension or consumer graph, so add your actual parsing and rendering paths before production.

## Current boundaries

CommonMark, GFM, OfficeIMO profile behavior, optional extension families, and source-model limits are tracked independently. The generated inventories record 651 of 652 CommonMark `0.31.2` examples and all 52 tracked GFM fixtures as matching the current contracts. Optional families remain explicitly `Covered`, `Partial`, `Intentional`, `Gap`, or `Unsupported`; applications should choose a parser profile and review the compatibility matrix for the exact boundary they depend on.

## Continue with

- [OfficeIMO.Markdown](/products/markdown/) for the package overview and runtime guidance.
- [Markdown documentation](/docs/markdown/) for the document model, parser profiles, and renderer surface.
- [Builder API guide](/docs/markdown/builder/) if you want to generate Markdown fluently.
- [Markdown compatibility matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.markdown.compatibility-matrix.md) for standards profiles, extension families, and source-model boundaries.
- [AOT and trimming guidance](/docs/advanced/aot-trimming/) for deployment notes across the package family.
