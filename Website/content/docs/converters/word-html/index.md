---
title: Word to HTML
description: Bidirectional conversion between Word documents and HTML using OfficeIMO.Word.Html.
order: 70
---

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
using System.Net;

string html = """
&lt;html&gt;
&lt;body&gt;
    &lt;h1&gt;Report Title&lt;/h1&gt;
    &lt;p&gt;This is a &lt;strong&gt;bold&lt;/strong&gt; and &lt;em&gt;italic&lt;/em&gt; paragraph.&lt;/p&gt;
    &lt;table&gt;
        &lt;tr&gt;&lt;th&gt;Name&lt;/th&gt;&lt;th&gt;Value&lt;/th&gt;&lt;/tr&gt;
        &lt;tr&gt;&lt;td&gt;Alpha&lt;/td&gt;&lt;td&gt;100&lt;/td&gt;&lt;/tr&gt;
        &lt;tr&gt;&lt;td&gt;Beta&lt;/td&gt;&lt;td&gt;200&lt;/td&gt;&lt;/tr&gt;
    &lt;/table&gt;
&lt;/body&gt;
&lt;/html&gt;
""";

using var document = WebUtility.HtmlDecode(html).LoadFromHtml();
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
| `&lt;h1&gt;` -- `&lt;h6&gt;` | Heading1 -- Heading6 paragraph styles |
| `&lt;p&gt;` | Normal paragraph |
| `&lt;strong&gt;`, `&lt;b&gt;` | Bold formatting |
| `&lt;em&gt;`, `&lt;i&gt;` | Italic formatting |
| `&lt;u&gt;` | Underline formatting |
| `&lt;s&gt;`, `&lt;del&gt;` | Strikethrough |
| `&lt;table&gt;`, `&lt;tr&gt;`, `&lt;td&gt;`, `&lt;th&gt;` | Word tables with cells |
| `&lt;ul&gt;`, `&lt;ol&gt;`, `&lt;li&gt;` | Bulleted and numbered lists |
| `&lt;img&gt;` | Inline images |
| `&lt;a&gt;` | Hyperlinks |
| `&lt;br&gt;` | Line breaks |
| `&lt;hr&gt;` | Horizontal rules |
| `&lt;code&gt;`, `&lt;pre&gt;` | Monospace formatting |

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
