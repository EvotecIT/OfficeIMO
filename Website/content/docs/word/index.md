---
title: Create and Edit Word Documents in .NET
description: Create, open, read, edit, and save DOCX files with OfficeIMO.Word, PSWriteOffice, and the officeimo CLI, without Microsoft Word installed.
order: 10
packages:
  - OfficeIMO.Word
cmdlets:
  - New-OfficeWord
  - Open-OfficeWord
  - Get-OfficeWordText
  - Export-OfficeWord
search_terms:
  - docx
  - word document
  - create word file
  - read docx
  - doc to docx
commands:
  - officeimo read
verified: 2026-11-14
---

`OfficeIMO.Word` creates, reads, and edits Word documents (`.docx`, `.docm`, `.dotx`) and opens legacy binary `.doc` files into the same model. `WordDocument` is the document; it holds sections, and sections hold paragraphs, tables, images, headers, and footers. It runs anywhere .NET runs, with no Word or COM automation.

The guides in this section cover each part of a document: text and paragraphs, tables, images, headers and footers, lists, fields and the table of contents, templates, review, protection, and conversion.

## Create a document

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("welcome.docx");
document.AddParagraph("Welcome").Style = WordParagraphStyles.Heading1;
document.AddParagraph("This document was generated with OfficeIMO.Word.");
document.Save();
```

### PowerShell

```powershell
New-OfficeWord -Path .\welcome.docx {
    Add-OfficeWordHeading 'Welcome'
    Add-OfficeWordParagraph -Text 'This document was generated with PSWriteOffice.'
}
```

## Open a document and read its text

`Paragraphs` covers the whole body across sections, in document order.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("report.docx");
foreach (WordParagraph paragraph in document.Paragraphs) {
    if (!string.IsNullOrWhiteSpace(paragraph.Text)) Console.WriteLine(paragraph.Text);
}
Console.WriteLine($"{document.Sections.Count} sections, {document.Tables.Count} tables, {document.Images.Count} images");
```

### PowerShell

```powershell
Get-OfficeWordText -Path .\report.docx
```

### CLI

```sh
officeimo read report.docx
```

## Edit a document and save a copy

`Save(path)` writes the edited document to the path you give, so the result can go to a new file.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("report.docx");
document.ReplaceAll("DRAFT", "FINAL");
document.AddParagraph("Approved on " + DateTime.Today.ToString("yyyy-MM-dd"));
document.Save("report-final.docx");
```

### PowerShell

```powershell
Open-OfficeWord -Path .\report.docx {
    Add-OfficeWordParagraph -Text 'Approved'
}
```

## Create a document in memory

Create without a path and save to any stream, for example an HTTP response.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create();
document.AddParagraph("Generated in memory");

using var stream = new MemoryStream();
document.Save(stream);
Console.WriteLine($"{stream.Length} bytes");
```

## Turn a list of objects into a document

`Export-OfficeWord` writes a title and a table from pipeline objects in one call.

### PowerShell

```powershell
Get-Service | Select-Object Name, Status, StartType | Export-OfficeWord -Path .\services.docx -Title 'Services'
```

## Convert a legacy DOC file to DOCX

`Load` reads binary `.doc` files into the same model, so saving to `.docx` converts them.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("old-letter.doc");
document.Save("old-letter.docx");
```
