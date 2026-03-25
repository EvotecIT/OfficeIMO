---
title: Word to HTML
description: Bidirectional conversion between Word documents and HTML using OfficeIMO.Word.Html.
order: 70
---

# Word to HTML Conversion

The `OfficeIMO.Word.Html` package provides bidirectional conversion between Word documents and HTML. It uses [AngleSharp](https://anglesharp.github.io/) for HTML parsing and DOM manipulation, and supports both synchronous and asynchronous workflows.

## Installation

```bash
dotnet add package OfficeIMO.Word.Html
```

This package depends on `OfficeIMO.Word`, `AngleSharp`, and `AngleSharp.Css`.

## Word to HTML

### Convert to HTML String

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using var document = WordDocument.Load("report.docx");

// Convert to a full HTML document
string html = document.ToHtml();
Console.WriteLine(html);
```

### Async Conversion

```csharp
string html = await document.ToHtmlAsync();
```

### Save as HTML File

```csharp
document.SaveAsHtml("report.html");

// Async version
await document.SaveAsHtmlAsync("report.html");
```

### Save to a Stream

```csharp
using var stream = new MemoryStream();
document.SaveAsHtml(stream);
```

### Conversion Options

Customize the conversion with `WordToHtmlOptions`:

```csharp
var options = new WordToHtmlOptions {
    // Include or exclude document properties
    IncludeDocumentProperties = true,

    // Control image handling
    EmbedImages = true,  // Base64-encode images inline

    // CSS options
    IncludeStyles = true
};

string html = document.ToHtml(options);
```

## HTML to Word

### Create a Word Document from HTML

```csharp
using OfficeIMO.Word.Html;

string html = @"
<html>
<body>
    <h1>Report Title</h1>
    <p>This is a <strong>bold</strong> and <em>italic</em> paragraph.</p>
    <table>
        <tr><th>Name</th><th>Value</th></tr>
        <tr><td>Alpha</td><td>100</td></tr>
        <tr><td>Beta</td><td>200</td></tr>
    </table>
</body>
</html>";

using var document = html.LoadFromHtml();
document.Save("from-html.docx");
```

### Async HTML to Word

```csharp
using var document = await html.LoadFromHtmlAsync();
document.Save("from-html.docx");
```

### HTML to Word Options

```csharp
var options = new HtmlToWordOptions {
    // Default font settings
    DefaultFontFamily = "Calibri",
    DefaultFontSize = 11,

    // Page setup
    DefaultPageWidth = 12240,   // Letter width in twips
    DefaultPageHeight = 15840,  // Letter height in twips
};

using var document = html.LoadFromHtml(options);
```

## Adding HTML to an Existing Document

You can inject HTML content into specific parts of an existing Word document:

### Append HTML to Document Body

```csharp
using var document = WordDocument.Create("mixed.docx");
document.AddParagraph("Native OfficeIMO paragraph");

// Append HTML-sourced content
document.AddHtmlToBody("<p>This came from <strong>HTML</strong>.</p>");
document.Save();
```

### Add HTML to Headers

```csharp
document.AddHeadersAndFooters();
document.AddHtmlToHeader(
    "<p style='text-align: center;'>Company Header</p>",
    HeaderFooterValues.Default
);
```

### Add HTML to Footers

```csharp
document.AddHtmlToFooter(
    "<p style='font-size: 8pt; color: gray;'>Confidential</p>",
    HeaderFooterValues.Default
);
```

### Async Versions

All `AddHtml*` methods have async counterparts:

```csharp
await document.AddHtmlToBodyAsync("<p>Async HTML content</p>");
await document.AddHtmlToHeaderAsync("<p>Header from HTML</p>");
await document.AddHtmlToFooterAsync("<p>Footer from HTML</p>");
```

## Supported HTML Elements

The converter handles the following HTML elements:

| HTML Element | Word Equivalent |
|-------------|-----------------|
| `<h1>` -- `<h6>` | Heading1 -- Heading6 paragraph styles |
| `<p>` | Normal paragraph |
| `<strong>`, `<b>` | Bold formatting |
| `<em>`, `<i>` | Italic formatting |
| `<u>` | Underline formatting |
| `<s>`, `<del>` | Strikethrough |
| `<table>`, `<tr>`, `<td>`, `<th>` | Word tables with cells |
| `<ul>`, `<ol>`, `<li>` | Bulleted and numbered lists |
| `<img>` | Inline images |
| `<a>` | Hyperlinks |
| `<br>` | Line breaks |
| `<hr>` | Horizontal rules |
| `<code>`, `<pre>` | Monospace formatting |

## CSS Style Mapping

The converter maps common CSS properties to Word formatting:

- `font-family` maps to Word font family
- `font-size` maps to Word font size
- `font-weight: bold` maps to bold
- `font-style: italic` maps to italic
- `text-align` maps to paragraph alignment
- `color` maps to text color
- `background-color` maps to paragraph shading
- `text-decoration` maps to underline/strikethrough
