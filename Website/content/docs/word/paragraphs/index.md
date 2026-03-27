---
title: Paragraphs
description: Working with paragraphs in OfficeIMO.Word -- text formatting, alignment, spacing, fonts, and colors.
order: 11
---

# Paragraphs

Paragraphs are the fundamental text containers in a Word document. The `WordParagraph` class provides properties and methods for controlling text content, character formatting, paragraph-level layout, and inline elements such as hyperlinks and images.

## Creating Paragraphs

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("paragraphs.docx");

// Simple text paragraph
document.AddParagraph("Hello, World!");

// Paragraph with chained formatting
var p = document.AddParagraph("Important Notice");
p.Bold = true;
p.FontSize = 16;
p.Color = SixLabors.ImageSharp.Color.Red;

document.Save();
```

## Text Formatting

### Bold, Italic, Underline, Strikethrough

```csharp
var p1 = document.AddParagraph("Bold text");
p1.Bold = true;

var p2 = document.AddParagraph("Italic text");
p2.Italic = true;

var p3 = document.AddParagraph("Underlined text");
p3.Underline = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single;

var p4 = document.AddParagraph("Struck through");
p4.Strike = true;

// Double strikethrough
var p5 = document.AddParagraph("Double strikethrough");
p5.DoubleStrike = true;
```

### Font Family and Size

```csharp
var p = document.AddParagraph("Custom font");
p.FontFamily = "Arial";
p.FontSize = 14;        // in half-points (14 = 7pt visual size)
p.FontSizeComplexScript = 14;
```

### Text Color and Highlighting

```csharp
var p = document.AddParagraph("Colored text");
p.Color = SixLabors.ImageSharp.Color.DarkBlue;

var p2 = document.AddParagraph("Highlighted text");
p2.Highlight = DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues.Yellow;
```

### Caps and Small Caps

```csharp
var p = document.AddParagraph("small caps example");
p.CapsStyle = CapsStyle.SmallCaps;

var p2 = document.AddParagraph("all caps example");
p2.CapsStyle = CapsStyle.Caps;
```

### Subscript and Superscript

```csharp
// Use VerticalTextAlignment for sub/superscript
var p = document.AddParagraph("H");
p.AddText("2").VerticalTextAlignment =
    DocumentFormat.OpenXml.Wordprocessing.VerticalPositionValues.Subscript;
p.AddText("O");
```

## Paragraph Alignment

```csharp
var left = document.AddParagraph("Left aligned");
left.ParagraphAlignment = JustificationValues.Left;

var center = document.AddParagraph("Centered");
center.ParagraphAlignment = JustificationValues.Center;

var right = document.AddParagraph("Right aligned");
right.ParagraphAlignment = JustificationValues.Right;

var justify = document.AddParagraph("Justified text spans the full width...");
justify.ParagraphAlignment = JustificationValues.Both;
```

## Spacing and Indentation

### Line Spacing

```csharp
var p = document.AddParagraph("Double-spaced paragraph");
p.LineSpacing = 480;         // in twips; 240 = single, 480 = double
p.LineSpacingRule = DocumentFormat.OpenXml.Wordprocessing.LineSpacingRuleValues.Auto;
```

### Spacing Before and After

```csharp
var p = document.AddParagraph("Spaced paragraph");
p.SpacingBefore = 200;      // in twips
p.SpacingAfter = 200;
```

### Indentation

```csharp
var p = document.AddParagraph("Indented paragraph");
p.IndentationFirstLine = 720;   // 720 twips = 0.5 inch
p.IndentationBefore = 360;      // left indent
```

## Heading Styles

```csharp
var h1 = document.AddParagraph("Chapter 1");
h1.Style = WordParagraphStyles.Heading1;

var h2 = document.AddParagraph("Section 1.1");
h2.Style = WordParagraphStyles.Heading2;

var h3 = document.AddParagraph("Subsection 1.1.1");
h3.Style = WordParagraphStyles.Heading3;
```

## Hyperlinks

```csharp
// External hyperlink
document.AddHyperLink(
    "Visit OfficeIMO",
    new Uri("https://github.com/EvotecIT/OfficeIMO"),
    addStyle: true
);

// Internal bookmark link
document.AddHyperLink(
    "Go to Chapter 1",
    "chapter1_bookmark",
    addStyle: true
);
```

## Page Breaks

```csharp
document.AddParagraph("Content before break");
document.AddPageBreak();
document.AddParagraph("Content on next page");
```

## Horizontal Lines

```csharp
document.AddHorizontalLine(
    lineType: DocumentFormat.OpenXml.Wordprocessing.BorderValues.Single,
    color: SixLabors.ImageSharp.Color.Gray,
    size: 12
);
```

## Bookmarks

```csharp
// Add a bookmark
document.AddBookmark("chapter1_bookmark");
document.AddParagraph("Chapter 1: Introduction");
```

## Multiple Runs in a Paragraph

A paragraph can contain multiple runs with different formatting:

```csharp
var p = document.AddParagraph();
p.AddText("Normal text, ");
var boldRun = p.AddText("bold text, ");
boldRun.Bold = true;
var italicRun = p.AddText("and italic.");
italicRun.Italic = true;
```

## Paragraph Shading

```csharp
var p = document.AddParagraph("Shaded paragraph");
p.ShadingFill = "FFFF00";     // Yellow background
p.ShadingPattern = DocumentFormat.OpenXml.Wordprocessing.ShadingPatternValues.Clear;
```
