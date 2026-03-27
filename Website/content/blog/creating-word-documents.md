---
title: "Creating Word Documents Without Microsoft Word Installed"
description: "A step-by-step tutorial showing how to create professional Word documents with paragraphs, tables, and images using OfficeIMO.Word in C#."
date: 2025-02-10
tags: [word, docx, getting-started]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

One of the most common questions we hear is: "Can I really generate Word documents on a server without installing Office?" The answer is yes, and with **OfficeIMO.Word** it takes fewer lines of code than you might expect.

## Installation

Add the NuGet package to your project:

```bash
dotnet add package OfficeIMO.Word
```

That is it. No COM references, no Office PIA assemblies, no registry hacks.

## Creating Your First Document

```csharp
using OfficeIMO.Word;

using var doc = WordDocument.Create("Report.docx");

// Title
var title = doc.AddParagraph("Quarterly Sales Report");
title.Bold = true;
title.FontSize = 24;
title.FontFamily = "Calibri";

// Body text
doc.AddParagraph("This report summarises sales performance for Q1 2025. " +
    "All figures are expressed in USD unless otherwise noted.");
```

The `WordDocument.Create` call initialises an in-memory DOCX package. Nothing is written to disk until you call `Save()`.

## Adding a Table

Tables are first-class citizens in OfficeIMO. You define them with a simple row-and-column API:

```csharp
var table = doc.AddTable(4, 3);

// Header row
table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Revenue";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Growth";

// Data rows
string[][] data = {
    new[] { "North America", "$1.2M", "+8%" },
    new[] { "Europe",        "$0.9M", "+5%" },
    new[] { "Asia-Pacific",  "$0.6M", "+12%" },
};

for (int r = 0; r < data.Length; r++)
    for (int c = 0; c < data[r].Length; c++)
        table.Rows[r + 1].Cells[c].Paragraphs[0].Text = data[r][c];
```

## Inserting an Image

Adding a logo or chart image is straightforward:

```csharp
var paragraph = doc.AddParagraph();
paragraph.AddImage("chart.png", width: 400, height: 250);
```

OfficeIMO embeds the image directly inside the DOCX archive, so the file is fully self-contained.

## Headers and Footers

Professional documents need headers and footers. OfficeIMO handles odd/even and first-page variants:

```csharp
doc.AddHeadersAndFooters();
doc.Header.Default.AddParagraph("Contoso Ltd. — Confidential");
doc.Footer.Default.AddParagraph("Page ");
```

## Saving

When you are finished assembling the document, call `Save`:

```csharp
doc.Save();
```

If you need a `Stream` instead of a file, use `Save(stream)` to write directly to a `MemoryStream` for HTTP responses or blob storage uploads.

## Full Example

```csharp
using OfficeIMO.Word;

using var doc = WordDocument.Create("Report.docx");

var title = doc.AddParagraph("Quarterly Sales Report");
title.Bold = true;
title.FontSize = 24;

doc.AddParagraph("Generated automatically by OfficeIMO.");

var table = doc.AddTable(2, 2);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Metric";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
table.Rows[1].Cells[0].Paragraphs[0].Text = "Total Revenue";
table.Rows[1].Cells[1].Paragraphs[0].Text = "$2.7M";

doc.Save();
Console.WriteLine("Report.docx created successfully.");
```

Run this with `dotnet run` and inspect the resulting file in Word or another OOXML-capable editor. OfficeIMO writes standard Open XML packages, but visual fidelity still depends on the specific viewer and feature set involved.

## Next Steps

In future posts we will cover styles, table of contents generation, mail merge, and batch document assembly. If you hit a question, open a GitHub issue and we will help you out.

## Continue with

- [OfficeIMO.Word](/products/word/) for the package overview and supported document features.
- [Word documentation](/docs/word/) for paragraphs, sections, styles, and layout structure.
- [Tables guide](/docs/word/tables/) if your next step is richer tabular output.
- [PSWriteOffice Word cmdlets](/docs/pswriteoffice/word/) if you want to generate similar documents from PowerShell.
