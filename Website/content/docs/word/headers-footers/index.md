---
title: "Add Headers, Footers and Page Numbers"
description: "Add headers and footers to DOCX files, insert page numbers and \"Page X of Y\", and use a different first page or odd and even pages with OfficeIMO.Word."
order: 14
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordHeader
  - Add-OfficeWordFooter
  - Add-OfficeWordPageNumber
  - Add-OfficeWordSection
  - Set-OfficeWordPageSetup
search_terms:
  - word header
  - word footer
  - page numbers
  - page x of y
  - landscape section
verified: 2026-11-14
---

Headers, footers, page size, margins, and orientation belong to a section, and every document has at least one. `document.PageSetup` is the first section's setup; `section.PageSetup` is any other's. `Header` and `Footer` create the default story on first use. `FirstPageHeader`/`FirstPageFooter` and `EvenPageHeader`/`EvenPageFooter` also turn on `DifferentFirstPage` or `DifferentOddAndEvenPages` when they create their story.

## Add a header and a footer

Headers and footers take the same paragraph API as the document body.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("headers.docx");

WordParagraph header = document.PageSetup.Header.AddParagraph("Contoso - Confidential");
header.ParagraphAlignment = WordParagraphAlignment.Right;
header.FontSize = OfficeLength.Points(9);

document.PageSetup.Footer.AddParagraph("Printed " + DateTime.Today.ToString("yyyy-MM-dd"));

document.AddParagraph("Body text.");
document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\headers.docx {
    Add-OfficeWordSection {
        Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Contoso - Confidential' -Alignment Right }
        Add-OfficeWordFooter { Add-OfficeWordParagraph -Text 'Internal use only' }
        Add-OfficeWordParagraph -Text 'Body text.'
    }
}
```

## Add page numbers

`AddPageNumber(includeTotalPages: true)` writes "1 of 5" style numbering with PAGE and NUMPAGES fields, which Word updates as the document changes.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("numbered.docx");

WordParagraph footer = document.PageSetup.Footer.AddParagraph("Page ");
footer.ParagraphAlignment = WordParagraphAlignment.Center;
footer.AddPageNumber(includeTotalPages: true);

for (int i = 1; i <= 3; i++) {
    document.AddParagraph($"Chapter {i}").Style = WordParagraphStyles.Heading1;
    document.AddPageBreak();
}
document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\numbered.docx {
    Add-OfficeWordSection {
        Add-OfficeWordFooter { Add-OfficeWordPageNumber -IncludeTotal }
        Add-OfficeWordParagraph -Text 'Chapter 1' -Style Heading1
    }
}
```

## Use a different first page

A cover page usually has no header. Setting content on `FirstPageHeader` switches the section to a separate first-page story.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("cover.docx");

document.PageSetup.FirstPageHeader.AddParagraph(string.Empty);
document.PageSetup.Header.AddParagraph("Annual report 2026");

document.AddParagraph("Annual report").Style = WordParagraphStyles.Title;
document.AddPageBreak();
document.AddParagraph("The header appears from page two.");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\cover.docx {
    Add-OfficeWordSection {
        Add-OfficeWordHeader -Type First { Add-OfficeWordParagraph -Text '' }
        Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Annual report 2026' }
        Add-OfficeWordParagraph -Text 'Annual report' -Style Title
    }
}
```

## Use odd and even page headers

For double-sided printing, the default header is used on odd pages and `EvenPageHeader` on even pages.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("book.docx");

document.PageSetup.Header.AddParagraph("Chapter title").ParagraphAlignment = WordParagraphAlignment.Right;
document.PageSetup.EvenPageHeader.AddParagraph("Book title").ParagraphAlignment = WordParagraphAlignment.Left;

for (int i = 0; i < 4; i++) {
    document.AddParagraph($"Page {i + 1}");
    document.AddPageBreak();
}
document.Save();
```

## Change page layout and headers for a new section

A new section starts on a new page and can have its own orientation, size, margins, and headers.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("sections.docx");

document.PageSetup.Header.AddParagraph("Report");
document.AddParagraph("Portrait content.");

WordSection appendix = document.AddSection();
appendix.PageSetup.Orientation = OfficePageOrientation.Landscape;
appendix.PageSetup.Header.AddParagraph("Appendix - wide tables");
appendix.AddParagraph("Landscape content.");

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\sections.docx {
    Add-OfficeWordSection {
        Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Report' }
        Add-OfficeWordParagraph -Text 'Portrait content.'
    }
    Add-OfficeWordSection -BreakType NextPage {
        Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape
        Add-OfficeWordHeader { Add-OfficeWordParagraph -Text 'Appendix - wide tables' }
        Add-OfficeWordParagraph -Text 'Landscape content.'
    }
}
```

## Add an image to a header

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("logo-header.docx");
document.PageSetup.Header.AddParagraph().AddImage("logo.png", OfficeLength.Centimeters(3), OfficeLength.Centimeters(1),
    new WordImageOptions { AltText = "Contoso logo" });
document.AddParagraph("Body text.");
document.Save();
```
