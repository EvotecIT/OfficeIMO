---
title: "Convert Between Word and Markdown"
description: "Convert DOCX to a typed Markdown model and save it as Markdown or HTML, and build Word documents from Markdown with OfficeIMO.Word.Markdown."
order: 71
packages:
  - OfficeIMO.Word.Markdown
  - OfficeIMO.Markdown
  - OfficeIMO.Markdown.Html
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeWordMarkdownOptions
  - New-OfficeMarkdownWordOptions
commands:
  - officeimo convert
limits: "Nested tables, custom table styles, SmartArt, and embedded OLE objects have no Markdown equivalent and are simplified or left out."
search_terms:
  - docx to markdown
  - markdown to docx
  - word to md
  - md to word
verified: 2026-11-14
---

`OfficeIMO.Word.Markdown` converts between `WordDocument` and the `MarkdownDocument` model from `OfficeIMO.Markdown`. `ToMarkdown()` on a Word document returns the model, which saves itself; call `ToMarkdown()` on the model when you need a string. `OfficeIMO.Markdown.Html` adds HTML on either side of that model.

## Convert a Word document to Markdown

### C#

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

Saving to a stream leaves the stream open:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown();
using var stream = File.Create("report-stream.md");
markdown.Save(stream);
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\report.docx -To Markdown -OutputPath .\report.md
```

### CLI

A `.md` destination uses the Reader's Markdown projection; [Convert documents with the CLI](/docs/cli/convert/#markdown-and-json-output) covers its options.

```sh
officeimo convert report.docx report.md
```

## Keep underline and highlighting

Markdown has no standard underline or highlight. Export options turn them on as extensions.

### C#

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

### PowerShell

```powershell
$options = New-OfficeWordMarkdownOptions -EnableUnderline -EnableHighlight
Convert-OfficeDocument -Path .\report.docx -To Markdown -OutputPath .\report.md -Options $options
```

## Create a Word document from Markdown

Parse Markdown into its model, then convert the model to Word.

### C#

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

const string source = "# Summary\n\nThis report covers **Q4 performance**.\n\n| Metric | Value |\n| --- | --- |\n| Revenue | $1.2M |\n| Users | 5,000 |";
var options = new MarkdownWordExportOptions();
using var markdown = MarkdownDocument.Parse(source, options.CreateReaderOptions());
using var document = markdown.ToWord(options);
document.Save("from-markdown.docx");
```

The same conversion works for a file or a stream you own:

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

### PowerShell

```powershell
$options = New-OfficeMarkdownWordOptions -Template .\Templates\Runbook.dotx -BookmarkName Body
Convert-OfficeDocument -Path .\runbook.md -To Word -OutputPath .\runbook.docx -Options $options
```

### CLI

```sh
officeimo convert notes.md notes.docx
```

## Convert through Markdown to and from HTML

Save the Markdown model of a Word document as a complete HTML page:

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Load("report.docx");
using var markdown = document.ToMarkdown();
markdown.SaveAsHtml("report.html");
```

For HTML to Word through Markdown, convert the HTML to the Markdown model first:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

var html = HtmlConversionDocument.Parse("<h1>Hello</h1><p>World</p>");
using var markdown = html.ToMarkdown();
using var document = markdown.ToWord();
document.Save("from-html.docx");
```

## How Word elements map to Markdown

| Word element | Markdown |
|---|---|
| Heading1 to Heading6 styles | `#` to `######` |
| Normal paragraph | Plain paragraph |
| Bold run | `**bold**` |
| Italic run | `*italic*` |
| Bold and italic run | `***bold italic***` |
| Strikethrough | `~~strikethrough~~` |
| Inline code style | `` `code` `` |
| Hyperlink | `[text](url)` |
| Table | GFM pipe table |
| Bulleted list | `- item` |
| Numbered list | `1. item` |
| Image | `![alt](path)` |
| Horizontal rule | `---` |
| Block quote style | `> text` |
