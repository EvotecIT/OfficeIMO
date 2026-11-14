---
title: Convert Markdown to HTML, PDF and Word
description: "Convert Markdown to HTML, PDF and Word, and HTML back to Markdown, with conversion reports that list anything that could not be kept."
order: 43
search_terms:
  - markdown to pdf
  - md to pdf
  - markdown to word
  - md to docx
  - markdown to html
  - html to markdown
packages:
  - OfficeIMO.Markdown
  - OfficeIMO.Markdown.Pdf
  - OfficeIMO.Word.Markdown
  - OfficeIMO.Markdown.Html
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeMarkdownPdfOptions
  - New-OfficeMarkdownWordOptions
  - New-OfficeHtmlMarkdownOptions
commands:
  - officeimo convert
related:
  - /docs/markdown/parse-and-inspect/
  - /docs/converters/word-markdown/
  - /docs/html/html-to-pdf/
limits: Markdown output cannot express merged table cells, cell borders or page layout; those conversions report the loss instead of dropping it silently.
verified: 2026-11-14
---

Every route starts from a `MarkdownDocument`, so you parse once and send the same model to any target. HTML rendering is built into `OfficeIMO.Markdown`. PDF needs `OfficeIMO.Markdown.Pdf`, Word needs `OfficeIMO.Word.Markdown`, and HTML input needs `OfficeIMO.Markdown.Html`. Each `SaveAs…` method returns an `OfficeConversionReport`; set `FailOnLoss` on the options when a conversion must stop instead of writing a lossy file.

```sh
dotnet add package OfficeIMO.Markdown.Pdf
dotnet add package OfficeIMO.Word.Markdown
dotnet add package OfficeIMO.Markdown.Html
```

## Convert Markdown to an HTML page

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;

File.WriteAllText("status.md", "# Status\n\n| Service | State |\n| --- | --- |\n| API | Green |\n");

MarkdownDocument document = MarkdownDocument.Load("status.md");
OfficeConversionReport report = document.SaveAsHtml("status.html", new MarkdownHtmlExportOptions {
    Kind = MarkdownHtmlKind.Document,
    Style = MarkdownHtmlStyle.Clean,
    Title = "Service status"
});

Console.WriteLine($"{report.Target}: {report.OutputPath}");
```

## Convert Markdown to PDF

Front matter can supply the PDF title, author and theme. Local images are read only when you allow them with a resource policy.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;

MarkdownDocument document = MarkdownDocument.Parse("""
---
title: Incident review
author: Operations
---
# Incident review

## Summary

- Detection was automatic.
- Customer impact was limited.

| Task | Owner |
| --- | --- |
| Update monitor | Platform |
""");

OfficeConversionReport report = document.SaveAsPdf("incident-review.pdf", new MarkdownPdfExportOptions {
    UseFrontMatterMetadata = true,
    Theme = MarkdownVisualTheme.Report()
});

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

### PowerShell

```powershell
$options = New-OfficeMarkdownPdfOptions -Title 'Service report' -IncludeLocalImages -BaseDirectory .\Assets
Convert-OfficeDocument -Path .\Report.md -OutputPath .\Report.pdf -Options $options
```

### CLI

```sh
officeimo convert Report.md Report.pdf
```

## Convert Markdown to Word

`ToWord` maps headings, lists, tables, images, code, footnotes and TeX math to native Word content, and turns `[TOC]` markers into a Word table of contents.

### C#

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

MarkdownDocument markdown = MarkdownDocument.Parse("""
# Runbook

[TOC]

## Prepare

Stop the **scheduler** before you start.

## Deploy

1. Copy the package.
2. Restart the service.
""");

using WordDocument word = markdown.ToWord(new MarkdownWordExportOptions {
    AllowLocalImages = true,
    FitImagesToPageContentWidth = true
});
word.Save("runbook.docx");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\Runbook.md -To Word -OutputPath .\Runbook.docx

# Fill a template at a bookmark instead of creating a blank document
$options = New-OfficeMarkdownWordOptions -Template .\Templates\Runbook.dotx -BookmarkName Body
Convert-OfficeDocument -Path .\Runbook.md -To Word -OutputPath .\Runbook.docx -Options $options
```

### CLI

```sh
officeimo convert Runbook.md Runbook.docx
```

## Convert HTML to Markdown

Parse the HTML with `HtmlConversionDocument`, then call `ToMarkdown`. Profiles choose the Markdown dialect: GitHub-flavored, CommonMark or portable.

### C#

```csharp
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse("""
<h1>Release notes</h1>
<p>See the <a href="guide/start">getting started guide</a>.</p>
<table><tr><th>Area</th><th>State</th></tr><tr><td>PDF</td><td>Done</td></tr></table>
<div class="cookie-banner">Accept cookies</div>
""");

HtmlMarkdownExportOptions options = HtmlMarkdownExportOptions.CreateGitHubFlavoredMarkdownProfile();
options.BaseUri = new Uri("https://example.com/docs/");
options.ExcludeSelectors.Add(".cookie-banner");

using MarkdownDocument document = source.ToMarkdown(options);
document.Save("release-notes.md");
Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
$options = New-OfficeHtmlMarkdownOptions -Portable -LineEnding LF
Convert-OfficeDocument -Path .\Report.html -To Markdown -OutputPath .\Report.md -Options $options
```

### CLI

```sh
officeimo convert Report.html Report.md
```

## Use one theme for HTML, PDF and Word

A `MarkdownVisualTheme` carries fonts, colors, spacing and table styling into every adapter, so the three outputs look alike.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

MarkdownVisualTheme theme = MarkdownVisualTheme.FromOfficeTheme(OfficeTheme.Boardroom)
    .WithColorScheme(MarkdownColorSchemeKind.Emerald);

using MarkdownDocument document = MarkdownDocument.Parse("# Quarterly report\n\n| Region | Sales |\n| --- | ---: |\n| North | 120 |\n");

document.SaveAsHtml("report.html", new MarkdownHtmlExportOptions { Kind = MarkdownHtmlKind.Document, Theme = theme });
document.SaveAsPdf("report.pdf", new MarkdownPdfExportOptions { Theme = theme });

using WordDocument word = document.ToWord(new MarkdownWordExportOptions { Theme = theme });
word.Save("report.docx");
```
