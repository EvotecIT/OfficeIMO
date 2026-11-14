---
title: "Add a Table of Contents and Fields in Word"
description: "Insert a table of contents, page, date and property fields into DOCX files, refresh them on open or in code, and read fields with OfficeIMO.Word."
order: 16
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordTableOfContents
  - Update-OfficeWordTableOfContents
  - Set-OfficeWordTableOfContents
  - Add-OfficeWordField
  - Get-OfficeWordField
  - Update-OfficeWordField
limits: Table of contents entries and page references are calculated by Word when the document is opened; OfficeIMO marks them for refresh rather than paginating the document itself.
search_terms:
  - table of contents
  - word toc
  - word fields
  - date field
  - docproperty field
verified: 2026-11-14
---

A table of contents and fields such as PAGE, DATE, or DOCPROPERTY are Word field codes: the document stores the instruction, and Word calculates the result. OfficeIMO writes the field code and its last known result, and can flag the document so Word refreshes every field when it opens. A table of contents lists paragraphs that use the built-in heading styles.

## Add a table of contents

`AddTableOfContent` inserts the table at the current position. `minLevel` and `maxLevel` choose which heading levels appear, and `UpdateFieldsOnOpen` asks Word to fill in entries and page numbers on open.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("report.docx");

WordTableOfContent toc = document.AddTableOfContent(WordTableOfContentsStyle.Template1, minLevel: 1, maxLevel: 2);
toc.Text = "Contents";
document.Settings.UpdateFieldsOnOpen = true;
document.AddPageBreak();

document.AddParagraph("Executive summary").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Summary text.");
document.AddParagraph("Findings").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Availability").Style = WordParagraphStyles.Heading2;

document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\report.docx {
    Add-OfficeWordTableOfContents -Style Template1
    Set-OfficeWordTableOfContents -Text 'Contents' -TextNoContent 'No headings yet'
    Add-OfficeWordParagraph -Text 'Executive summary' -Style Heading1
    Add-OfficeWordParagraph -Text 'Summary text.'
    Add-OfficeWordParagraph -Text 'Findings' -Style Heading1
    Update-OfficeWordTableOfContents
}
```

## Refresh or rebuild an existing table of contents

`Update` flags the table for refresh when Word opens the file. `RegenerateTableOfContent` removes the table and writes a new one in the same place, which helps after a template's table was damaged or restyled.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("report.docx");

if (document.TableOfContent is WordTableOfContent toc) {
    toc.Update();
} else {
    document.AddTableOfContent(insertAtTop: true);
}
document.Save();

using WordDocument damaged = WordDocument.Load("imported.docx");
damaged.RegenerateTableOfContent();
damaged.Save();
```

### PowerShell

```powershell
$doc = Open-OfficeWord -Path .\report.docx
$doc | Update-OfficeWordTableOfContents -Regenerate
$doc | Save-OfficeWord -Path .\report-rebuilt.docx
Close-OfficeWord -Document $doc
```

## Insert page, date, and property fields

`AddField` appends a field to a paragraph. `WordFieldFormat` applies a format switch, and `customFormat` sets a date picture for DATE and TIME fields.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("fields.docx");
document.Properties.Title = "Service review";

WordParagraph title = document.AddParagraph("Title: ");
title.AddField(WordFieldType.Title);

WordParagraph date = document.AddParagraph("Generated on ");
date.AddField(WordFieldType.Date, customFormat: "yyyy-MM-dd");

WordParagraph author = document.AddParagraph("Author: ");
author.AddField(WordFieldType.Author, WordFieldFormat.Upper);

document.PageSetup.Footer.AddParagraph("Page ").AddPageNumber(includeTotalPages: true);

document.Settings.UpdateFieldsOnOpen = true;
document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\fields.docx {
    Set-OfficeWordDocumentProperty -Name Title -Value 'Service review'
    Add-OfficeWordParagraph { Add-OfficeWordText 'Title: '; Add-OfficeWordField -Type Title }
    Add-OfficeWordParagraph { Add-OfficeWordText 'Generated on '; Add-OfficeWordField -Type Date -CustomFormat 'yyyy-MM-dd' }
    Add-OfficeWordParagraph { Add-OfficeWordText 'Status: '; Add-OfficeWordField -DocumentProperty Category }
    Update-OfficeWordField
}
```

## Update page number fields in code

`UpdateFields` recalculates PAGE and NUMPAGES results so the saved file shows sensible numbers even before Word refreshes it. When the document has a table of contents, it also sets the refresh-on-open flag.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("numbers.docx");
document.AddParagraph("Page ").AddPageNumber(includeTotalPages: true);
document.AddPageBreak();
document.AddParagraph("Second page");

document.UpdateFields();
document.Save();
```

## List the fields in a document

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("fields.docx");
foreach (WordField field in document.Fields) {
    Console.WriteLine($"{field.FieldType}: {field.Field} -> {field.Text}");
}
```

### PowerShell

```powershell
Get-OfficeWordField -Path .\fields.docx -FieldType Date, Title
```
