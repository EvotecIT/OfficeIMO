---
title: Headers and Footers
description: "Adding headers, footers, and page numbers to Word documents with OfficeIMO.Word. Includes practical examples, limits, and API links."
order: 14
---

OfficeIMO provides full support for document headers and footers, including different content for the first page, odd/even pages, and section-specific headers and footers. Page numbers can be inserted in various formats.

## Page Setup

Running content belongs to a section's `PageSetup`, next to its page size, margins and orientation. `document.PageSetup` is the first section's setup. `Header` and `Footer` create the default header and footer on first use; `FirstPageHeader`/`FirstPageFooter` and `EvenPageHeader`/`EvenPageFooter` enable `DifferentFirstPage` or `DifferentOddAndEvenPages` when they create their story. `GetHeader(type)` and `GetFooter(type)` return an existing story or `null` without creating one.

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("headers.docx");
WordPageSetup setup = document.PageSetup;
```

## Adding Header Content

Each section has three headers: `Header`, `FirstPageHeader`, and `EvenPageHeader`.

```csharp
// Default header (appears on all pages unless overridden)
document.PageSetup.Header.AddParagraph("Company Name - Confidential");

// First page header (different from the rest)
document.PageSetup.DifferentFirstPage = true;
document.PageSetup.FirstPageHeader.AddParagraph("COVER PAGE HEADER");

// Even page header (for double-sided printing)
document.PageSetup.DifferentOddAndEvenPages = true;
document.PageSetup.EvenPageHeader.AddParagraph("Even Page Header");
```

## Adding Footer Content

Footers work identically to headers:

```csharp
document.PageSetup.Footer.AddParagraph("Copyright 2025 - All Rights Reserved");

document.PageSetup.FirstPageFooter.AddParagraph("Draft Document");
```

## Page Numbers

Add page numbers to headers or footers using `AddPageNumber`:

```csharp
// Add page number to the default footer
document.PageSetup.Footer.AddPageNumber(WordPageNumberStyle.Normal);
```

Available `WordPageNumberStyle` values:

| Style | Output |
|-------|--------|
| `WordPageNumberStyle.Normal` | 1, 2, 3... |
| `WordPageNumberStyle.PageOfTotal` | Page 1 of 5 |

### Centering Page Numbers

```csharp
var footerParagraph = document.PageSetup.Footer.AddParagraph();
footerParagraph.ParagraphAlignment = WordParagraphAlignment.Center;
document.PageSetup.Footer.AddPageNumber(WordPageNumberStyle.Normal);
```

## Different First Page

Set `DifferentFirstPage` to use separate header/footer content on the first page of a section:

```csharp
using var document = WordDocument.Create("first-page.docx");
document.PageSetup.DifferentFirstPage = true;

// First page gets a special header
document.PageSetup.FirstPageHeader.AddParagraph("TITLE PAGE");

// All other pages get the default header
document.PageSetup.Header.AddParagraph("Chapter 1 - Introduction");

document.AddParagraph("Title page content...");
document.AddPageBreak();
document.AddParagraph("Second page content...");

document.Save();
```

## Different Odd and Even Pages

For double-sided printing, enable different odd/even page headers:

```csharp
document.PageSetup.DifferentOddAndEvenPages = true;

document.PageSetup.Header.AddParagraph("Odd Page Header");   // Appears on odd pages
document.PageSetup.EvenPageHeader.AddParagraph("Even Page Header");
```

## Section-Specific Headers and Footers

Each section can have independent headers and footers. Add a new section and configure its headers separately:

```csharp
using var document = WordDocument.Create("sections.docx");

// Section 1
document.PageSetup.Header.AddParagraph("Chapter 1");
document.AddParagraph("Chapter 1 content...");

// Section 2 with different header
var section2 = document.AddSection();
section2.PageSetup.Header.AddParagraph("Chapter 2");
document.AddParagraph("Chapter 2 content...");

document.Save();
```

## Adding Images to Headers

```csharp
// Add a logo to the default header using VML
document.Sections[0].PageSetup.Header.AddImageVml("logo.png", width: 120, height: 40);
```

## Removing Headers

Remove all headers from the document:

```csharp
WordHeader.RemoveHeaders(document._wordprocessingDocument);
```

Or remove specific header types:

```csharp
WordHeader.RemoveHeaders(
    document,
    WordHeaderFooterType.Default,
    WordHeaderFooterType.First
);
```

## Header and Footer Paragraphs

Headers and footers support the full paragraph API, including formatted text, hyperlinks, fields, and tab stops:

```csharp
var headerParagraph = document.PageSetup.Header.AddParagraph();
headerParagraph.Text = "Document Title";
headerParagraph.Bold = true;
headerParagraph.FontSize = 10;
headerParagraph.ParagraphAlignment = WordParagraphAlignment.Center;

var footerParagraph = document.PageSetup.Footer.AddParagraph();
footerParagraph.Text = "Printed: " + DateTime.Now.ToString("yyyy-MM-dd");
footerParagraph.FontSize = 8;
footerParagraph.Color = OfficeIMO.OfficeColor.Gray;
```
