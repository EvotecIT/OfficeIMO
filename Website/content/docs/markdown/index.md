---
title: Markdown
description: "Overview of the OfficeIMO.Markdown package — fluent builder, typed reader/AST, and HTML renderer. Includes examples and API links."
order: 40
---

The `OfficeIMO.Markdown` package provides a focused Markdown toolkit for .NET with three main capabilities:

1. **Fluent Builder** — Compose Markdown documents programmatically with a chainable API.
2. **Reader / AST** — Parse Markdown text into a typed Abstract Syntax Tree (AST) for analysis and transformation.
3. **HTML Renderer** — Convert Markdown documents to HTML with GitHub-like output, syntax highlighting, and table-of-contents generation.

The package has **zero external dependencies** and targets .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2.

## Key Classes

| Class | Description |
|-------|-------------|
| `MarkdownDocument` | Builds, parses, loads and saves the typed Markdown document model. |
| `HeadingBlock` | AST node representing a heading (H1--H6). |
| `ParagraphBlock` | AST node for a paragraph with inline formatting. |
| `CodeBlock` | Fenced or indented code block with language annotation. |
| `TableBlock` | AST node for a GFM-style table. |
| `ListBlock` | Ordered or unordered list with nested items. |
| `CalloutBlock` | Admonition / callout block (Note, Warning, Tip, etc.). |
| `QuoteBlock` | Block quotation. |
| `ImageBlock` | Image reference with alt text, title, and optional size hints. |
| `FrontMatterBlock` | YAML front matter parsed into typed entries. |
| `MarkdownHtmlExportOptions` | Configuration for the HTML rendering pipeline. |

## Creating a Markdown Document

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDocument.Create()
    .H1("Project README")
    .P("A short description of the project.")
    .H2("Installation")
    .Code("bash", "dotnet add package MyProject")
    .H2("Usage")
    .P("Import the namespace and create an instance.")
    .Code("csharp", "var client = new MyClient();")
    .H2("License")
    .P("MIT License");

var markdown = doc.ToMarkdown();
Console.WriteLine(markdown);
```

## Parsing Markdown

```csharp
string source = File.ReadAllText("README.md");
MarkdownDocument doc = MarkdownDocument.Parse(source);

// Enumerate headings
foreach (var heading in doc.DescendantHeadings()) {
    Console.WriteLine($"H{heading.Level}: {heading.Text}");
}

// Find a specific heading
var installSection = doc.FindHeading("Installation");
```

## Rendering to HTML

```csharp
var doc = MarkdownDocument.Create()
    .H1("Hello")
    .P("World");

var options = new MarkdownHtmlExportOptions {
    Style = MarkdownHtmlStyle.GitHub,
    InjectTocAtTop = true,
    Prism = new MarkdownPrismOptions { Enabled = true }
};

string html = doc.ToHtml(options);
```

## Document Scaffolds

OfficeIMO.Markdown includes pre-built scaffolds for common document types:

```csharp
// Generate a README scaffold
var readme = ReadmeScaffold.Readme("MyProject", opts => {
    opts.Description = "A cross-platform library for data processing.";
    opts.License = "MIT";
    opts.Author = "Your Name";
});

File.WriteAllText("README.md", readme.ToString());

// Generate a CHANGELOG scaffold
var changelog = MarkdownChangelogScaffold.Changelog("MyProject");

// Generate a CONTRIBUTING guide
var contributing = MarkdownContributingScaffold.Contributing("MyProject");
```

## Further Reading

- [Builder API](/docs/markdown/builder) — Detailed guide to the fluent builder methods.
- [Word to Markdown Conversion](/docs/converters/word-markdown) — Convert between Word documents and Markdown.
