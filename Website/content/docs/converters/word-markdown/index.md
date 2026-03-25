---
title: Word to Markdown
description: Bidirectional conversion between Word documents and Markdown using OfficeIMO.Word.Markdown.
order: 71
---

# Word to Markdown Conversion

The `OfficeIMO.Word.Markdown` package provides bidirectional conversion between Word documents and Markdown. It builds on `OfficeIMO.Word`, `OfficeIMO.Markdown`, and `OfficeIMO.Word.Html` to handle headings, formatting, tables, lists, images, and more.

## Installation

```bash
dotnet add package OfficeIMO.Word.Markdown
```

This package depends on `OfficeIMO.Word`, `OfficeIMO.Word.Html`, `OfficeIMO.Markdown`, and `OfficeIMO.Markdown.Html`.

## Word to Markdown

### Convert to Markdown String

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");

string markdown = document.ToMarkdown();
Console.WriteLine(markdown);
```

### Async Conversion

```csharp
string markdown = await document.ToMarkdownAsync();
```

### Convert to Markdown AST

Get a typed `MarkdownDoc` AST for further processing:

```csharp
using OfficeIMO.Markdown;

MarkdownDoc markdownDoc = document.ToMarkdownDocument();

// Inspect the structure
foreach (var heading in markdownDoc.DescendantHeadings()) {
    Console.WriteLine($"H{heading.Level}: {heading.Text}");
}

// Render to string
string output = markdownDoc.ToString();
```

### Save as Markdown File

```csharp
document.SaveAsMarkdown("report.md");

// Async version
await document.SaveAsMarkdownAsync("report.md");
```

### Save to a Stream

```csharp
using var stream = new MemoryStream();
document.SaveAsMarkdown(stream);
```

### Conversion Options

```csharp
var options = new WordToMarkdownOptions {
    // Configure heading mapping, image handling, etc.
};

string markdown = document.ToMarkdown(options);
```

## Markdown to Word

### Create a Word Document from Markdown String

```csharp
using OfficeIMO.Word.Markdown;

string markdown = @"
# Project Report

## Summary

This report covers **Q4 performance**.

| Metric | Value |
|--------|-------|
| Revenue | $1.2M |
| Users | 5,000 |

## Conclusion

Overall positive results.
";

using var document = markdown.LoadFromMarkdown();
document.Save("from-markdown.docx");
```

### Create from a Markdown File

```csharp
using var document = WordMarkdownConverterExtensions.LoadFromMarkdown(
    "README.md"
);
document.Save("readme.docx");
```

### Create from a Stream

```csharp
using var stream = File.OpenRead("document.md");
using var document = stream.LoadFromMarkdown();
document.Save("output.docx");
```

### Async Markdown to Word

```csharp
using var document = await "README.md".LoadFromMarkdownAsync();
document.Save("readme.docx");
```

### Create from a MarkdownDoc AST

If you have already parsed or built a `MarkdownDoc`, convert it directly:

```csharp
using OfficeIMO.Markdown;

var md = MarkdownDoc.Create()
    .H1("Generated Report")
    .P("Created programmatically.")
    .Table(t => t
        .Header("Item", "Count")
        .Row("Widgets", "42")
        .Row("Gadgets", "17")
    );

using var document = md.ToWordDocument();
document.Save("generated.docx");
```

### Markdown to Word Options

```csharp
var options = new MarkdownToWordOptions {
    // Configure default styles, fonts, etc.
};

using var document = markdown.LoadFromMarkdown(options);
```

## HTML via Markdown

Convert Word to HTML by going through the Markdown pipeline. This can produce cleaner HTML than the direct Word-to-HTML converter for simple documents:

```csharp
using var document = WordDocument.Load("report.docx");

// Word -> Markdown -> HTML
string html = document.ToHtmlViaMarkdown();
```

For an HTML fragment (no `<html>`, `<head>`, `<body>` wrapper):

```csharp
string fragment = document.ToHtmlFragmentViaMarkdown();
```

Save the HTML output to a file:

```csharp
document.SaveAsHtmlViaMarkdown("report.html");
```

### HTML to Word via Markdown

Convert HTML to Word by first converting to Markdown, then to Word:

```csharp
string html = "<h1>Hello</h1><p>World</p>";

using var document = html.LoadFromHtmlViaMarkdown();
document.Save("from-html-via-md.docx");
```

From a stream:

```csharp
using var htmlStream = File.OpenRead("page.html");
using var document = htmlStream.LoadFromHtmlViaMarkdown();
document.Save("output.docx");
```

## Element Mapping

The converter maps between Word and Markdown elements:

| Word Element | Markdown Element |
|-------------|-----------------|
| Heading1--Heading6 styles | `# ` through `###### ` |
| Normal paragraph | Plain paragraph text |
| Bold run | `**bold**` |
| Italic run | `*italic*` |
| Bold + Italic run | `***bold italic***` |
| Strikethrough | `~~strikethrough~~` |
| Inline code style | `` `code` `` |
| Hyperlink | `[text](url)` |
| Table | GFM pipe table |
| Bulleted list | `- item` |
| Numbered list | `1. item` |
| Image | `![alt](path)` |
| Horizontal rule | `---` |
| Block quote style | `> text` |

## Round-Trip Fidelity

While the converters handle the most common elements, some formatting may be lost during round-trip conversion (Word -> Markdown -> Word or vice versa). Complex features like nested tables, custom table styles, SmartArt, and embedded OLE objects do not have Markdown equivalents and will be simplified or omitted.
