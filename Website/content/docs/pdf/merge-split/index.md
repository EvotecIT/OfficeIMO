---
title: "Merge, Split, and Select PDF Pages"
description: "Merge PDFs, split a PDF into parts, and extract, delete, reorder, rotate, or duplicate pages with OfficeIMO.Pdf, PSWriteOffice, and the officeimo CLI."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Merge-OfficeDocument
  - Split-OfficeDocument
search_terms:
  - merge pdf
  - split pdf
  - combine pdf
  - extract pages
commands:
  - officeimo workflow assemble
  - officeimo pdf split
  - officeimo pdf extract
verified: 2026-11-14
---

Page operations never change the loaded document in place: each call returns a new `PdfDocument`, so you can chain operations and save the result. Pages are numbered from 1, and page selections use the familiar print-dialog syntax: `"1-3,5"`, `"2"`, `"3,1-2"`.

## Merge PDFs

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("agenda.pdf")
    .MergeWith("minutes.pdf")
    .MergeWith("appendix.pdf")
    .Save("board-pack.pdf");
```

To merge documents that are already loaded, pass them together:

```csharp
using OfficeIMO.Pdf;

PdfDocument[] parts = Directory.GetFiles("weekly", "*.pdf")
    .Order()
    .Select(path => PdfDocument.Load(path))
    .ToArray();

using PdfDocument merged = PdfDocument.Merge(parts);
merged.Save("monthly.pdf");
```

### PowerShell

```powershell
Merge-OfficeDocument -Path .\agenda.pdf, .\minutes.pdf -OutputPath .\board-pack.pdf -PassThru
```

### CLI

`officeimo workflow assemble` combines PDFs, Office documents, images, and folders into one PDF:

```sh
officeimo workflow assemble agenda.pdf minutes.pdf appendix.pdf --output board-pack.pdf
```

## Split a PDF

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("scans.pdf");

IReadOnlyList<PdfDocument> single = pdf.Pages.Split();       // one document per page
IReadOnlyList<PdfDocument> pairs = pdf.Pages.Split(2);       // two pages per document
for (int i = 0; i < pairs.Count; i++) {
    pairs[i].Save($"record-{i + 1}.pdf");
}
```

### PowerShell

```powershell
Get-ChildItem .\Scans -Filter *.pdf | Split-OfficeDocument -OutputDirectory .\Records -PageCount 2 -NoClobber
Split-OfficeDocument -LiteralPath '.\Manual [final].pdf' -OutputDirectory .\Chapters -ByBookmark
```

### CLI

```sh
officeimo pdf split scans.pdf --pages-per-document 2 --output records
```

## Extract, delete, reorder, and rotate pages

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("packet.pdf");

pdf.Pages.Extract("1-3").Save("summary.pdf");
pdf.Pages.Delete("2").Save("without-page-2.pdf");
pdf.Pages.Reorder("2,3,1").Save("reordered.pdf");
pdf.Pages.Rotate(90, "2").Save("rotated.pdf");
pdf.Pages.Duplicate("1").Save("with-copy.pdf");
pdf.Pages.Move(1, "3").Save("page-3-first.pdf");
```

### PowerShell

```powershell
Copy-OfficePdfPage -Path .\packet.pdf -Page '1-3' -OutputPath .\summary.pdf
Remove-OfficePdfPage -Path .\packet.pdf -Page '4-5' -OutputPath .\packet-clean.pdf -PassThru
Open-OfficePdf .\Report.pdf { Remove-OfficePdfPage -Page last }
```

### CLI

```sh
officeimo pdf extract packet.pdf --pages last,1-3 --output selected.pdf
```
