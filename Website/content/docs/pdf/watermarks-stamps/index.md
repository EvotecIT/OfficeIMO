---
title: "Watermarks, Stamps, and Metadata"
description: "Add text or image watermarks and stamps to new or existing PDFs and set title, author, and keywords with OfficeIMO.Pdf and PSWriteOffice."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Add-OfficePdfStamp
search_terms:
  - watermark pdf
  - stamp pdf
  - pdf metadata
verified: 2026-11-14
---

A watermark is drawn across the page, usually behind or over the content with low opacity. A stamp places text or an image at a position on selected pages, such as an approval mark or a page label. Positions are measured in points from the bottom-left corner of the page.

## Watermark a new PDF

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Draft service review")
        .AddParagraph("This copy is not final.")),
        new PdfCreateOptions { TextWatermark = new PdfTextWatermark("DRAFT") { Opacity = 0.18 } })
    .Save("draft.pdf");
```

### PowerShell

```powershell
New-OfficePdf -Path .\draft.pdf {
    Add-OfficePdfWatermark -Text 'DRAFT' -Size 72 -Opacity 0.12 -Rotation -35 -Color '#64748B'
    PdfHeading 'Draft service review'
}
```

## Stamp an existing PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;

PdfDocument.Load("report.pdf")
    .Stamp.Text("Approved", new PdfTextStampOptions {
        PageNumbers = new[] { 1 },
        X = OfficeLength.Points(450),
        Y = OfficeLength.Points(36),
        FontSize = OfficeLength.Points(12),
        Color = OfficeColor.FromRgb(22, 101, 52)
    })
    .Save("report-approved.pdf");

PdfDocument.Load("report.pdf")
    .Stamp.TextWatermark("Confidential", new PdfTextStampOptions {
        X = OfficeLength.Points(150),
        Y = OfficeLength.Points(400),
        FontSize = OfficeLength.Points(48)
    })
    .Save("report-confidential.pdf");
```

`Stamp.Image` and `Stamp.ImageWatermark` place images, and `Stamp.OverlayPage` puts a page from another PDF (such as letterhead) over every page.

### PowerShell

```powershell
Add-OfficePdfStamp -Path .\report.pdf -OutputPath .\report-approved.pdf -Text 'Approved' -Page 1 -X 450 -Y 36
Add-OfficePdfStamp -Path .\report.pdf -OutputPath .\report-letterhead.pdf -OverlayPath .\letterhead.pdf -Behind
```

## Set document metadata

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("report.pdf")
    .UpdateProperties(title: "Service review", author: "Operations")
    .Save("report-tagged.pdf");

Console.WriteLine(PdfDocument.Load("report-tagged.pdf").Inspect().Metadata.Title);
```

### PowerShell

```powershell
Set-OfficePdfMetadata -Path .\report.pdf -OutputPath .\report-tagged.pdf -Title 'Service review' -Author 'Operations'
```
