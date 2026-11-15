---
title: "PDF Quick Start"
description: "Create a PDF with headings, paragraphs, a table, and an image, then load it and read its text, with OfficeIMO.Pdf, PSWriteOffice, or the officeimo CLI."
layout: docs
packages:
  - OfficeIMO.Pdf
search_terms:
  - create pdf
  - generate pdf
  - pdf c#
commands:
  - officeimo convert
  - officeimo read
verified: 2026-11-14
---

OfficeIMO.Pdf writes and reads PDF files directly in .NET. It has no native dependencies and does not need a browser, printer driver, or Office installation.

```powershell
dotnet add package OfficeIMO.Pdf
```

## Create a PDF

`PdfDocument.Create` takes a builder callback. Content flows down the page and breaks onto new pages automatically.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;

byte[] logo = File.ReadAllBytes("logo.png");

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .AddParagraph("All scheduled checks passed. Two follow-ups remain open.")
        .AddTable(new[] {
            new[] { "Task", "Owner", "Due" },
            new[] { "Rotate certificates", "Platform", "4 Nov" },
            new[] { "Renew backup licence", "Operations", "12 Nov" } })
        .AddImage(logo, OfficeLength.Points(120), OfficeLength.Points(48), OfficeAlignment.Right)),
        new PdfCreateOptions { CreateOutlineFromHeadings = true })
    .Meta(title: "Service review", author: "Operations")
    .Save("Service-review.pdf");
```

`CreateOutlineFromHeadings` turns headings into bookmarks. For rich text inside a paragraph, use the callback overload:

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .AddParagraph(p => {
            p.AddText("Status: ");
            p.AddText("Healthy").SetBold();
        })))
    .Save("Status.pdf");
```

### PowerShell

```powershell
$services = @(
    [pscustomobject]@{ Name = 'Directory'; Status = 'Healthy'; Incidents = 0 }
    [pscustomobject]@{ Name = 'Mail'; Status = 'Watch'; Incidents = 2 }
)

New-OfficePdf -Path .\Service-review.pdf -CreateOutlineFromHeadings {
    PdfHeading 'Service review'
    PdfParagraph 'All scheduled checks passed. Two follow-ups remain open.'
    PdfTable -InputObject $services
}
```

### CLI

Convert an existing document instead of composing one:

```sh
officeimo convert report.docx report.pdf
officeimo convert notes.md notes.pdf
```

## Load a PDF and read it

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("Service-review.pdf");
Console.WriteLine($"{pdf.Pages.Count} page(s)");
Console.WriteLine(pdf.ExtractText());
```

### PowerShell

```powershell
Read-OfficeDocument -Path .\Service-review.pdf -As Text
```

### CLI

```sh
officeimo read Service-review.pdf --format markdown
```

## Next steps

- [Merge, split, and select pages](/docs/pdf/merge-split/)
- [Extract text and tables](/docs/pdf/extract-text-tables/)
- [Fill and flatten forms](/docs/pdf/forms/)
- [Redact content](/docs/pdf/redaction/)
- [Encrypt and sign](/docs/pdf/encrypt-sign/)
- [Convert to PDF/A](/docs/pdf/pdfa/)
- [Watermarks and stamps](/docs/pdf/watermarks-stamps/)
