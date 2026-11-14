---
title: "Convert Word to HTML and HTML to Word"
description: "Convert DOCX to HTML, build Word documents from HTML, and add HTML content to the body, headers, and footers of an existing document with OfficeIMO.Word.Html."
order: 70
packages:
  - OfficeIMO.Word.Html
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeWordHtmlOptions
  - New-OfficeHtmlWordOptions
commands:
  - officeimo convert
search_terms:
  - docx to html
  - html to docx
  - word to html
  - html to word
  - add html to word
verified: 2026-11-14
---

`OfficeIMO.Word.Html` converts in both directions. `ToHtml` and `SaveAsHtml` turn a `WordDocument` into HTML; `HtmlConversionDocument.Parse(html).ToWord()` builds a Word document from HTML; `AddHtml` inserts HTML into an existing document. HTML parsing uses AngleSharp. Check the output against your own fidelity needs when layouts are complex.

## Convert a Word document to HTML

`ToHtml` is an in-memory conversion. `SaveAsHtml` and `SaveAsHtmlAsync` write to a file or stream and return a conversion report.

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using WordDocument document = WordDocument.Load("report.docx");

string html = document.ToHtml();
Console.WriteLine(html.Length);

document.SaveAsHtml("report.html");
await document.SaveAsHtmlAsync("report-async.html");

using var stream = new MemoryStream();
document.SaveAsHtml(stream);
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\report.docx -To Html -OutputPath .\report.html
```

### CLI

```sh
officeimo convert report.docx report.html
```

## Control the HTML output

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using WordDocument document = WordDocument.Load("report.docx");

var options = new WordHtmlExportOptions {
    IncludeFontStyles = true,
    IncludeListStyles = true,
    EmbedImagesAsBase64 = true,
    IncludeDefaultCss = true
};

string html = document.ToHtml(options);
File.WriteAllText("report.html", html);
```

### PowerShell

```powershell
$options = New-OfficeWordHtmlOptions -IncludeFontStyles -IncludeListStyles -IncludeDefaultCss
Convert-OfficeDocument -Path .\report.docx -To Html -OutputPath .\report.html -Options $options
```

## Create a Word document from HTML

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

string html = """
    <h1>Report Title</h1>
    <p>This is a <strong>bold</strong> and <em>italic</em> paragraph.</p>
    <table>
        <tr><th>Name</th><th>Value</th></tr>
        <tr><td>Alpha</td><td>100</td></tr>
        <tr><td>Beta</td><td>200</td></tr>
    </table>
    """;

HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

using WordDocument document = source.ToWord(new HtmlWordExportOptions {
    FontFamily = "Calibri",
    IncludeListStyles = true,
    BasePath = AppContext.BaseDirectory
});
document.Save("from-html.docx");
```

`HtmlWordExportOptions.CreateTrustedDocumentProfile()` returns options for HTML you trust, and `ToWordAsync` loads external resources asynchronously:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse(File.ReadAllText("page.html"));
HtmlWordExportOptions options = HtmlWordExportOptions.CreateTrustedDocumentProfile();

using WordDocument document = (await source.ToWordAsync(options)).Value;
document.Save("page.docx");
```

### PowerShell

```powershell
$options = New-OfficeHtmlWordOptions -StylesheetPath .\report.css
Convert-OfficeDocument -Path .\report.html -To Word -OutputPath .\report.docx -Options $options
```

### CLI

```sh
officeimo convert report.html report.docx
```

## Add HTML to an existing document

`AddHtml` appends HTML-sourced content to the body, a table cell, a header, or a footer, and returns the new elements. Each method has an async counterpart.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using WordDocument document = WordDocument.Create("mixed.docx");
document.AddParagraph("Native OfficeIMO paragraph");
document.AddHtml("<p>This came from <strong>HTML</strong>.</p>");

WordSection section = document.Sections[0];

section.PageSetup.Header.AddHtml(
    HtmlConversionDocument.Parse("<p style='text-align: center;'>Company Header</p>"),
    section: section);
await section.PageSetup.Footer.AddHtmlAsync(
    HtmlConversionDocument.Parse("<p style='font-size: 8pt; color: gray;'>Confidential</p>"),
    section: section);

document.Save();
```

## Supported HTML elements and CSS

| HTML element | Word result |
|---|---|
| `<h1>` to `<h6>` | Heading1 to Heading6 paragraph styles |
| `<p>` | Normal paragraph |
| `<strong>`, `<b>` | Bold |
| `<em>`, `<i>` | Italic |
| `<u>` | Underline |
| `<s>`, `<del>` | Strikethrough |
| `<table>`, `<tr>`, `<td>`, `<th>` | Word table |
| `<ul>`, `<ol>`, `<li>` | Bulleted and numbered lists |
| `<img>` | Inline image |
| `<a>` | Hyperlink |
| `<br>` | Line break |
| `<hr>` | Horizontal rule |
| `<code>`, `<pre>` | Monospace text |

CSS `font-family`, `font-size`, `font-weight`, `font-style`, `text-align`, `color`, `background-color`, and `text-decoration` map to the matching Word run and paragraph formatting.
