---
title: "Convert to PDF/A"
description: "Convert existing PDFs to PDF/A-2, PDF/A-3, or PDF/A-4 archival profiles, assess readiness, and create PDF/A documents with OfficeIMO.Pdf, PSWriteOffice, and the officeimo CLI."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Get-OfficePdfCompliance
limits: "OfficeIMO reports what it changed and what blocks conversion; it does not certify conformance on its own. Confirm the saved file with an independent PDF/A validator when compliance is a requirement. PDF/A-4e is not a conversion target."
search_terms:
  - pdf/a
  - archive pdf
  - pdfa conversion
---

PDF/A is the ISO profile for long-term archiving: fonts are embedded, colors are device-independent, and features that depend on external resources are removed. OfficeIMO.Pdf supports the `PdfA2B` (default), `PdfA2U`, `PdfA3B`, `PdfA3U`, `PdfA4`, and `PdfA4F` profiles.

## Convert an existing PDF

### C#

```csharp
using OfficeIMO.Pdf;

PdfAConversionResult result = PdfDocument.Load("contract.pdf")
    .ConvertToPdfA(new PdfAConversionOptions { Profile = PdfComplianceProfile.PdfA3U });

if (result.Succeeded) {
    result.Save("contract-archive.pdf");
} else {
    foreach (var violation in result.Report.Violations) Console.WriteLine(violation);
}
```

### CLI

```sh
officeimo pdf convert-to-pdfa contract.pdf --output contract-archive.pdf --profile 3u
```

## Assess a file before converting

### C#

```csharp
using OfficeIMO.Pdf;

PdfComplianceReadinessReport readiness = PdfDocument.Load("contract.pdf")
    .AssessCompliance(PdfComplianceProfile.PdfA3U);
Console.WriteLine(readiness);
```

### PowerShell

```powershell
Get-OfficePdfCompliance -Path .\contract.pdf -Profile PdfA3U
```

## Create a new PDF as PDF/A

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Archive record")
        .AddParagraph("Retained for ten years.")),
        new PdfCreateOptions().UsePdfA(PdfComplianceProfile.PdfA2B))
    .Save("archive-record.pdf");
```

### PowerShell

```powershell
New-OfficePdf -Path .\archive-record.pdf {
    Set-OfficePdfCompliance -Profile PdfA3B -Groundwork -Language 'en-US'
    PdfHeading 'Archive record'
}
```
