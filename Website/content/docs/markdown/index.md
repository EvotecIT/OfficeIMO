---
title: Markdown
description: Overview of the OfficeIMO.Markdown package -- fluent builder, typed reader/AST, and HTML renderer.
order: 40
---

# Markdown

The `OfficeIMO.Markdown` package provides a complete Markdown toolkit for .NET with three main capabilities:

1. **Fluent Builder** -- Compose Markdown documents programmatically with a chainable API.
2. **Reader / AST** -- Parse Markdown text into a typed Abstract Syntax Tree (AST) for analysis and transformation.
3. **HTML Renderer** -- Convert Markdown documents to HTML with GitHub-like output, syntax highlighting, and table-of-contents generation.

The package has **zero external dependencies** and targets .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2.

## Key Classes

| Class | Description |
|-------|-------------|
| `MarkdownDoc` | Root document class and fluent API entry point. |
| `MarkdownReader` | Static parser that converts Markdown text into a `MarkdownDoc` AST. |
| `HeadingBlock` | AST node representing a heading (H1--H6). |
| `ParagraphBlock` | AST node for a paragraph with inline formatting. |
| `CodeBlock` | Fenced or indented code block with language annotation. |
| `TableBlock` | AST node for a GFM-style table. |
| `ListBlock` | Ordered or unordered list with nested items. |
| `CalloutBlock` | Admonition / callout block (Note, Warning, Tip, etc.). |
| `QuoteBlock` | Block quotation. |
| `ImageBlock` | Image reference with alt text, title, and optional size hints. |
| `FrontMatterBlock` | YAML front matter parsed into typed entries. |
| `HtmlOptions` | Configuration for the HTML rendering pipeline. |

## Creating a Markdown Document

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDoc.Create()
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
MarkdownDoc doc = MarkdownReader.Parse(source);

// Enumerate headings
foreach (var heading in doc.DescendantHeadings()) {
    Console.WriteLine($"H{heading.Level}: {heading.Text}");
}

// Find a specific heading
var installSection = doc.FindHeading("Installation");
```

## Rendering to HTML

```csharp
var doc = MarkdownDoc.Create()
    .H1("Hello")
    .P("World");

var options = new HtmlOptions {
    Style = HtmlStyle.GitHub,
    InjectTocAtTop = true,
    Prism = new PrismOptions { Enabled = true }
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
var changelog = ChangelogScaffold.Changelog("MyProject");

// Generate a CONTRIBUTING guide
var contributing = ContributingScaffold.Contributing("MyProject");
```

## Further Reading

- [Builder API](/docs/markdown/builder) -- Detailed guide to the fluent builder methods.
- [Word to Markdown Conversion](/docs/converters/word-markdown) -- Convert between Word documents and Markdown.
