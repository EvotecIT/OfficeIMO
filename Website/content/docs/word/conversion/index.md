---
title: Convert Word to PDF, HTML, and Markdown
description: Convert DOCX files to PDF, HTML, Markdown, ODT, and RTF, and create Word documents from HTML or Markdown, with OfficeIMO.Word converters, PSWriteOffice, and the officeimo CLI.
order: 20
packages:
  - OfficeIMO.Word
  - OfficeIMO.Word.Pdf
  - OfficeIMO.Word.Html
  - OfficeIMO.Word.Markdown
  - OfficeIMO.Word.OpenDocument
  - OfficeIMO.Word.Rtf
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeWordPdfOptions
  - New-OfficeWordMarkdownOptions
limits: PDF output is laid out by OfficeIMO's own renderer; complex layouts, floating objects, and fonts missing on the machine can paginate differently from Word, and the conversion report lists what was approximated.
search_terms:
  - docx to pdf
  - word to pdf
  - docx to html
  - docx to markdown
  - markdown to docx
  - html to word
commands:
  - officeimo convert
verified: 2026-11-14
---

Each destination format has its own package so an application only takes the converter it ships. Install `OfficeIMO.Word` plus the converter, then call the `SaveAs...` extension method it adds to `WordDocument`. Every conversion returns an `OfficeConversionReport`; its `Diagnostics` list content that was approximated or dropped, so check them, or set `FailOnLoss` on the options to throw instead.

| Destination | Package | Method |
|---|---|---|
| PDF | `OfficeIMO.Word.Pdf` | `SaveAsPdf`, `ToPdf` |
| HTML | `OfficeIMO.Word.Html` | `SaveAsHtml`, `ToHtml` |
| Markdown | `OfficeIMO.Word.Markdown` | `SaveAsMarkdown`, `ToMarkdown` |
| OpenDocument (ODT) | `OfficeIMO.Word.OpenDocument` | `SaveAsOpenDocument` |
| RTF | `OfficeIMO.Word.Rtf` | `SaveAsRtf` |

## Convert a Word document to PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using WordDocument document = WordDocument.Load("contract.docx");
OfficeConversionReport report = document.SaveAsPdf("contract.pdf");

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\contract.docx -OutputPath .\contract.pdf
Get-ChildItem .\contracts -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory .\pdf -SummaryVariable summary
$summary
```

### CLI

```sh
officeimo convert contract.docx contract.pdf
officeimo convert contract.docx contract.pdf --require-no-loss
```

## Stop when the PDF would lose content

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using WordDocument document = WordDocument.Load("contract.docx");
document.SaveAsPdf("contract.pdf", new WordPdfExportOptions {
    FailOnLoss = true,
    FontFamily = "Arial"
});
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\contract.docx -OutputPath .\contract.pdf -FailOnLoss
```

## Convert a Word document to HTML

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

using WordDocument document = WordDocument.Load("newsletter.docx");
document.SaveAsHtml("newsletter.html");

string fragment = document.ToHtml();
Console.WriteLine(fragment.Length);
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\newsletter.docx -To Html -OutputPath .\newsletter.html
```

## Create a Word document from HTML

`HtmlConversionDocument` parses the HTML once; `SaveAsWord` writes the DOCX. To add HTML into a document you are already building, use `document.AddHtml(html)`.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

HtmlConversionDocument html = HtmlConversionDocument.Parse(
    "<h1>Release notes</h1><p>Version <strong>2.0</strong> is available.</p><ul><li>Faster load</li></ul>");
html.SaveAsWord("release-notes.docx");

using WordDocument document = WordDocument.Create("mixed.docx");
document.AddParagraph("Generated summary").Style = WordParagraphStyles.Heading1;
document.AddHtml("<p>Imported <em>HTML</em> content.</p>");
document.Save();
```

### PowerShell

```powershell
'<h1>Review</h1>', '<p>Approved.</p>' | Convert-OfficeDocument -Text { $_ } -From Html -To Word -OutputPath .\review.docx
```

### CLI

```sh
officeimo convert release-notes.html release-notes.docx
```

## Convert a Word document to Markdown

Headings, lists, tables, links, emphasis, and images map to Markdown. Features with no Markdown equivalent, such as page layout or text boxes, are reported in the diagnostics.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using WordDocument document = WordDocument.Load("handbook.docx");
OfficeConversionReport report = document.SaveAsMarkdown("handbook.md");
Console.WriteLine($"{report.Diagnostics.Count} diagnostics");

MarkdownDocument markdown = document.ToMarkdown();
Console.WriteLine(markdown.ToMarkdown());
```

### PowerShell

```powershell
$options = New-OfficeWordMarkdownOptions -ImageExportMode File -ImageDirectory .\handbook-images
Convert-OfficeDocument -Path .\handbook.docx -To Markdown -OutputPath .\handbook.md -Options $options
```

### CLI

```sh
officeimo convert handbook.docx handbook.md --assets handbook-images
```

## Create a Word document from Markdown

### C#

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

MarkdownDocument markdown = MarkdownDocument.Parse("""
    # Release notes

    Version **2.0** adds:

    - Faster load
    - Dark theme
    """);
markdown.SaveAsWord("release-notes.docx");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\release-notes.md -To Word -OutputPath .\release-notes.docx
```

### CLI

```sh
officeimo convert release-notes.md release-notes.docx
```

## Convert to OpenDocument or RTF

### C#

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using OfficeIMO.Word.Rtf;

using WordDocument document = WordDocument.Load("letter.docx");
document.SaveAsOpenDocument("letter.odt");
document.SaveAsRtf("letter.rtf");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\letter.docx -OutputPath .\letter.odt
```

### CLI

```sh
officeimo convert letter.docx letter.odt
officeimo convert letter.rtf letter.docx
```
