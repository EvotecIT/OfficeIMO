---
title: "Format Text and Paragraphs in Word"
description: "Add paragraphs and headings to DOCX files and set fonts, bold, colors, alignment, spacing, indentation, hyperlinks and page breaks with OfficeIMO.Word."
order: 11
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordParagraph
  - Add-OfficeWordText
  - Add-OfficeWordHeading
  - Add-OfficeWordBookmark
  - Add-OfficeWordBreak
search_terms:
  - word paragraph
  - docx text formatting
  - word heading
  - word hyperlink
  - page break
verified: 2026-11-14
---

A `WordParagraph` holds text runs, and each run carries its own character formatting. `document.AddParagraph("text")` creates a paragraph with one run; `paragraph.AddText` appends more runs. Formatting set on the paragraph object (such as `Bold` or `Color`) applies to its current run, while alignment, spacing, indentation, and style apply to the whole paragraph.

## Add headings and paragraphs

Built-in heading styles are what a table of contents and the navigation pane read, so use them for document structure.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("report.docx");

document.AddParagraph("Service review").Style = WordParagraphStyles.Heading1;
document.AddParagraph("All services were available during the period.");
document.AddParagraph("Availability").Style = WordParagraphStyles.Heading2;
document.AddParagraph("Monthly availability stayed above target.");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\report.docx {
    Add-OfficeWordHeading 'Service review'
    Add-OfficeWordParagraph -Text 'All services were available during the period.'
    Add-OfficeWordHeading 'Availability' -Level 2
    Add-OfficeWordParagraph -Text 'Monthly availability stayed above target.'
}
```

## Mix formatting inside one paragraph

Each `AddText` call returns the new run, so formatting applies only to that piece of text.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("formatting.docx");

WordParagraph paragraph = document.AddParagraph("Status: ");
paragraph.AddText("approved").SetBold().SetColor(OfficeColor.Parse("2E7D32"));
paragraph.AddText(", effective ");
paragraph.AddText("immediately").SetItalic();

WordParagraph note = document.AddParagraph("Read this first");
note.FontFamily = "Arial";
note.FontSize = OfficeLength.Points(14);
note.Underline = WordUnderlineStyle.Single;
note.Highlight = WordHighlightColor.Yellow;

WordParagraph formula = document.AddParagraph("H");
formula.AddText("2").VerticalTextAlignment = WordVerticalTextPosition.Subscript;
formula.AddText("O");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\formatting.docx {
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Status: '
        Add-OfficeWordText 'approved' -Bold -Color '#2E7D32'
        Add-OfficeWordText ', effective '
        Add-OfficeWordText 'immediately' -Italic
    }
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Read this first' -Font Arial -Size 14pt -Underline Single
    }
}
```

## Set alignment, spacing, and indentation

Spacing and indentation take an `OfficeLength`, so the unit is explicit. `WordLineSpacing.Multiple(1.5)` sets line spacing relative to the font; `Exact` and `AtLeast` set a fixed height.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("layout.docx");

WordParagraph title = document.AddParagraph("Annual report");
title.ParagraphAlignment = WordParagraphAlignment.Center;
title.LineSpacingAfter = OfficeLength.Points(12);

WordParagraph body = document.AddParagraph("Justified text fills the full line width and wraps evenly.");
body.ParagraphAlignment = WordParagraphAlignment.Both;
body.LineSpacing = WordLineSpacing.Multiple(1.5);
body.IndentationFirstLine = OfficeLength.Inches(0.5);

WordParagraph quote = document.AddParagraph("Indented quotation.");
quote.IndentationBefore = OfficeLength.Centimeters(1.5);
quote.IndentationAfter = OfficeLength.Centimeters(1.5);

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\layout.docx {
    Add-OfficeWordParagraph -Text 'Annual report' -Alignment Center
    Add-OfficeWordParagraph -Text 'Justified text fills the full line width and wraps evenly.' -Alignment Justify
}
```

## Add hyperlinks and bookmarks

A hyperlink either opens a URL or jumps to a bookmark inside the document.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("links.docx");

document.AddParagraph("Contents");
document.AddHyperLink("Jump to the appendix", "Appendix", addStyle: true);
document.AddHyperLink("Project site", new Uri("https://github.com/EvotecIT/OfficeIMO"), addStyle: true);

document.AddPageBreak();
WordParagraph appendix = document.AddParagraph("Appendix");
appendix.Style = WordParagraphStyles.Heading1;
appendix.AddBookmark("Appendix");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\links.docx {
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Jump to the appendix' -Anchor Appendix
    }
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Project site' -Link 'https://github.com/EvotecIT/OfficeIMO'
    }
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Appendix'
        Add-OfficeWordBookmark -Name Appendix
    }
}
```

## Add page breaks, line breaks, and horizontal lines

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("breaks.docx");

WordParagraph address = document.AddParagraph("Contoso Ltd");
address.AddBreak();
address.AddText("1 Main Street");

document.AddHorizontalLine(WordBorderStyle.Single, OfficeColor.Parse("808080"), 12);
document.AddPageBreak();
document.AddParagraph("This paragraph starts on a new page.");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\breaks.docx {
    Add-OfficeWordParagraph {
        Add-OfficeWordText 'Contoso Ltd'
        Add-OfficeWordBreak
        Add-OfficeWordText '1 Main Street'
    }
    Add-OfficeWordParagraph { Add-OfficeWordBreak -BreakType Page }
    Add-OfficeWordParagraph -Text 'This paragraph starts on a new page.'
}
```

## Set document properties

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("properties.docx");
document.Properties.Title = "Quarterly report";
document.Properties.Creator = "Finance team";
document.AddParagraph("Quarterly report");
document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\properties.docx {
    Set-OfficeWordDocumentProperty -Name Title -Value 'Quarterly report'
    Set-OfficeWordDocumentProperty -Name ReleaseStatus -Value 'Approved' -Custom
    Add-OfficeWordParagraph -Text 'Quarterly report'
}
```
