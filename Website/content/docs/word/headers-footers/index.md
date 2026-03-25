---
title: Headers and Footers
description: Adding headers, footers, and page numbers to Word documents with OfficeIMO.Word.
order: 14
---

# Headers and Footers

OfficeIMO provides full support for document headers and footers, including different content for the first page, odd/even pages, and section-specific headers and footers. Page numbers can be inserted in various formats.

## Enabling Headers and Footers

Before adding content to headers or footers, you must enable them on the document:

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("headers.docx");

// Enable headers and footers for the current section
document.AddHeadersAndFooters();
```

## Adding Header Content

Each section has three header slots: `Default`, `First`, and `Even`.

```csharp
document.AddHeadersAndFooters();

// Default header (appears on all pages unless overridden)
document.Header.Default.AddParagraph("Company Name - Confidential");

// First page header (different from the rest)
document.DifferentFirstPage = true;
document.Header.First.AddParagraph("COVER PAGE HEADER");

// Even page header (for double-sided printing)
document.DifferentOddAndEvenPages = true;
document.Header.Even.AddParagraph("Even Page Header");
```

## Adding Footer Content

Footers work identically to headers:

```csharp
document.Footer.Default.AddParagraph("Copyright 2025 - All Rights Reserved");

document.Footer.First.AddParagraph("Draft Document");
```

## Page Numbers

Add page numbers to headers or footers using `AddPageNumber`:

```csharp
document.AddHeadersAndFooters();

// Add page number to the default footer
document.Footer.Default.AddPageNumber(WordPageNumberStyle.Normal);
```

Available `WordPageNumberStyle` values:

| Style | Output |
|-------|--------|
| `WordPageNumberStyle.Normal` | 1, 2, 3... |
| `WordPageNumberStyle.PageOfTotal` | Page 1 of 5 |

### Centering Page Numbers

```csharp
var footerParagraph = document.Footer.Default.AddParagraph();
footerParagraph.ParagraphAlignment = JustificationValues.Center;
document.Footer.Default.AddPageNumber(WordPageNumberStyle.Normal);
```

## Different First Page

Set `DifferentFirstPage` to use separate header/footer content on the first page of a section:

```csharp
using var document = WordDocument.Create("first-page.docx");
document.AddHeadersAndFooters();
document.DifferentFirstPage = true;

// First page gets a special header
document.Header.First.AddParagraph("TITLE PAGE");

// All other pages get the default header
document.Header.Default.AddParagraph("Chapter 1 - Introduction");

document.AddParagraph("Title page content...");
document.AddPageBreak();
document.AddParagraph("Second page content...");

document.Save();
```

## Different Odd and Even Pages

For double-sided printing, enable different odd/even page headers:

```csharp
document.DifferentOddAndEvenPages = true;

document.Header.Default.AddParagraph("Odd Page Header");   // Appears on odd pages
document.Header.Even.AddParagraph("Even Page Header");
```

## Section-Specific Headers and Footers

Each section can have independent headers and footers. Add a new section and configure its headers separately:

```csharp
using var document = WordDocument.Create("sections.docx");

// Section 1
document.AddHeadersAndFooters();
document.Header.Default.AddParagraph("Chapter 1");
document.AddParagraph("Chapter 1 content...");

// Section 2 with different header
var section2 = document.AddSection();
document.AddHeadersAndFooters();
document.Header.Default.AddParagraph("Chapter 2");
document.AddParagraph("Chapter 2 content...");

document.Save();
```

## Adding Images to Headers

```csharp
document.AddHeadersAndFooters();

// Add a logo to the default header using VML
document.Sections[0].Header.Default.AddImageVml("logo.png", width: 120, height: 40);
```

## Removing Headers

Remove all headers from the document:

```csharp
WordHeader.RemoveHeaders(document._wordprocessingDocument);
```

Or remove specific header types:

```csharp
WordHeader.RemoveHeaders(
    document._wordprocessingDocument,
    HeaderFooterValues.Default,
    HeaderFooterValues.First
);
```

## Header and Footer Paragraphs

Headers and footers support the full paragraph API -- you can add formatted text, hyperlinks, fields, and tab stops:

```csharp
document.AddHeadersAndFooters();

var headerParagraph = document.Header.Default.AddParagraph();
headerParagraph.Text = "Document Title";
headerParagraph.Bold = true;
headerParagraph.FontSize = 10;
headerParagraph.ParagraphAlignment = JustificationValues.Center;

var footerParagraph = document.Footer.Default.AddParagraph();
footerParagraph.Text = "Printed: " + DateTime.Now.ToString("yyyy-MM-dd");
footerParagraph.FontSize = 8;
footerParagraph.Color = SixLabors.ImageSharp.Color.Gray;
```
