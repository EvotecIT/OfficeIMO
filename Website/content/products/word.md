---
title: "OfficeIMO.Word"
description: "Create and edit Word documents without Microsoft Office. Paragraphs, tables, images, charts, watermarks, and more from pure .NET."
layout: product
product_color: "#2563eb"
install: "dotnet add package OfficeIMO.Word"
nuget: "OfficeIMO.Word"
docs_url: "/docs/word/"
api_url: "/api/word/"
---

## Why OfficeIMO.Word?

OfficeIMO.Word gives you full control over `.docx` files from any .NET application. No COM interop, no Office installation, no license fees. Build reports, contracts, invoices, and any document you can imagine with a clean, discoverable API.

## Features

- **Paragraphs & text styling** -- fonts, sizes, colors, bold, italic, underline, strikethrough, highlight, and spacing
- **Tables with merge & split** -- horizontal and vertical cell merging, nested tables, 105+ built-in table styles
- **Images** -- insert from file path, stream, Base64 string, or URL with precise positioning and sizing
- **Headers & footers** -- default, first page, and odd/even with text, images, and page numbers
- **Watermarks** -- text and image watermarks with rotation, color, and transparency
- **Table of Contents** -- automatic TOC generation with configurable heading levels
- **Bookmarks & hyperlinks** -- internal cross-references and external links
- **Charts** -- pie, bar, line, area, and combo charts with series data, legends, and axis formatting
- **Content controls** -- checkboxes, drop-down lists, combo boxes, date pickers, and rich text controls
- **Document protection** -- read-only, password protection, and editing restrictions
- **Footnotes & endnotes** -- numbered references with custom formatting
- **Sections & page numbering** -- multiple sections with independent orientation, margins, and numbering

## Quick start

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("Report.docx");

// Add a styled heading
var paragraph = document.AddParagraph("Quarterly Report");
paragraph.Style = WordParagraphStyle.Heading1;
paragraph.Color = SixLabors.ImageSharp.Color.DarkBlue;

// Add body text
document.AddParagraph("This report summarizes key metrics for Q4 2025.")
    .SetBold(false)
    .SetFontSize(12);

// Add a table
var table = document.AddTable(4, 3);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Region";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Revenue";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Growth";
table.Rows[1].Cells[0].Paragraphs[0].Text = "North America";
table.Rows[1].Cells[1].Paragraphs[0].Text = "$4.2M";
table.Rows[1].Cells[2].Paragraphs[0].Text = "+12%";
table.Rows[2].Cells[0].Paragraphs[0].Text = "Europe";
table.Rows[2].Cells[1].Paragraphs[0].Text = "$3.1M";
table.Rows[2].Cells[2].Paragraphs[0].Text = "+8%";
table.Rows[3].Cells[0].Paragraphs[0].Text = "Asia Pacific";
table.Rows[3].Cells[1].Paragraphs[0].Text = "$2.7M";
table.Rows[3].Cells[2].Paragraphs[0].Text = "+15%";
table.Style = WordTableStyle.GridTable4Accent1;

document.Save();
```

## Compatibility

| Target Framework  | Supported |
|-------------------|-----------|
| .NET 10.0         | Yes       |
| .NET 8.0          | Yes       |
| .NET Standard 2.0 | Yes       |
| .NET Framework 4.7.2 | Yes   |

OfficeIMO.Word runs on Windows, Linux, and macOS. It produces standard `.docx` files intended for Word and other OOXML-capable editors.

## Related guides

| Guide | Description |
|-------|-------------|
| [Word documentation](/docs/word/) | Start with the package overview and document structure. |
| [Tables guide](/docs/word/tables/) | Build styled tables, merged cells, and richer layouts. |
| [Word to HTML](/docs/converters/word-html/) | Convert generated documents to and from HTML. |
| [PSWriteOffice Word cmdlets](/docs/pswriteoffice/word/) | Automate Word output from PowerShell scripts. |

## Related packages

| Package | Description |
|---------|-------------|
| [OfficeIMO.Word.Html](/docs/converters/word-html/) | Convert Word documents to and from HTML |
| [OfficeIMO.Word.Markdown](/docs/converters/word-markdown/) | Convert Word documents to and from Markdown |
| [OfficeIMO.Word.Pdf](https://www.nuget.org/packages/OfficeIMO.Word.Pdf) | Export Word documents to PDF |
