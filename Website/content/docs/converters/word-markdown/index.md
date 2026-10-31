---
title: Word to Markdown
description: Convert Word documents to a typed Markdown model, and build Word documents from Markdown.
order: 71
---

`OfficeIMO.Word.Markdown` converts between `WordDocument` and `MarkdownDocument`. The returned model owns its save operations; call `ToMarkdown()` on that model when you need a string.

```bash
dotnet add package OfficeIMO.Word.Markdown
```

## Word to Markdown

Load a Word document, convert it, and save or inspect the Markdown model:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown();
Console.WriteLine(markdown.ToMarkdown());
markdown.Save("report.md");
await markdown.SaveAsync("report-copy.md");

foreach (var heading in markdown.DescendantHeadings()) {
    Console.WriteLine($"H{heading.Level}: {heading.Text}");
}
```

Conversion runs in memory. `SaveAsync` handles asynchronous output. A caller-owned stream remains open after saving:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown();
using var stream = File.Create("report-stream.md");
markdown.Save(stream);
```

Export options control Markdown features such as underline and highlighting:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown(new WordMarkdownExportOptions {
    EnableUnderline = true,
    EnableHighlight = true
});
Console.WriteLine(markdown.ToMarkdown());
```

## Markdown to Word

Parse Markdown into its native model, then convert that model to Word:

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

const string source = "# Summary\n\nThis report covers **Q4 performance**.\n\n| Metric | Value |\n| --- | --- |\n| Revenue | $1.2M |\n| Users | 5,000 |";
var options = new WordMarkdownImportOptions();
using var markdown = MarkdownDocument.Parse(source, options.CreateReaderOptions());
using var document = markdown.ToWord(options);
document.Save("from-markdown.docx");
```

Use the same conversion for a loaded file or caller-owned stream:

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

using var markdown = await MarkdownDocument.LoadAsync("input.md");
using var document = markdown.ToWord();
document.Save("from-file.docx");

using var stream = File.OpenRead("input.md");
using var fromStream = MarkdownDocument.Load(stream);
using var streamedDocument = fromStream.ToWord();
streamedDocument.Save("from-stream.docx");
```

## HTML through Markdown

Add the HTML adapter for conversions between HTML and the Markdown model:

```bash
dotnet add package OfficeIMO.Markdown.Html
```

Save a Word document's Markdown model as a complete HTML page:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown();
markdown.SaveAsHtml("report.html");
```

For HTML to Word, convert the HTML model to Markdown first:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

var html = HtmlConversionDocument.Parse("<h1>Hello</h1><p>World</p>");
using var markdown = html.ToMarkdown();
using var document = markdown.ToWord();
document.Save("from-html.docx");
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
