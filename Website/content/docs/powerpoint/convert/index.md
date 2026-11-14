---
title: Convert PowerPoint to PDF, HTML, and ODP
description: "Convert PPTX presentations to PDF slides, notes pages or handouts, to HTML, and to ODP with OfficeIMO.PowerPoint.Pdf, PSWriteOffice and the CLI."
order: 36
packages:
  - OfficeIMO.PowerPoint
  - OfficeIMO.PowerPoint.Pdf
  - OfficeIMO.PowerPoint.Html
  - OfficeIMO.PowerPoint.OpenDocument
cmdlets:
  - Convert-OfficeDocument
  - New-OfficePowerPointPdfOptions
limits: Conversion writes a static view of each slide; animations, transitions, and media playback are not carried into PDF or HTML.
search_terms:
  - pptx to pdf
  - powerpoint to pdf
  - presentation handouts
  - notes pages pdf
  - pptx to odp
  - pptx to html
commands:
  - officeimo convert
verified: 2026-11-14
---

Each output format has its own package, so an application only takes the renderer it ships: `OfficeIMO.PowerPoint.Pdf` for PDF, `OfficeIMO.PowerPoint.Html` for HTML, and `OfficeIMO.PowerPoint.OpenDocument` for ODP. They add `SaveAs...` extension methods to `PowerPointDocument`. Every conversion returns an `OfficeConversionReport` whose `Diagnostics` list what was approximated or left out; treat it as part of the result.

## Convert a presentation to PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
OfficeConversionReport report = presentation.SaveAsPdf("quarterly-review.pdf");

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}
```

Set `FailOnLoss = true` in `PowerPointPdfExportOptions` to throw instead of writing a PDF when the report records loss.

### PowerShell

```powershell
Convert-OfficeDocument -Path .\quarterly-review.pptx -OutputPath .\quarterly-review.pdf
Get-ChildItem .\decks -Filter *.pptx | Convert-OfficeDocument -To Pdf -OutputDirectory .\pdf
```

### CLI

```sh
officeimo convert quarterly-review.pptx quarterly-review.pdf
```

## Print notes pages or handouts

`PageLayout` switches from one slide per page to notes pages (slide plus speaker notes) or handouts with 1, 2, 3, 4, 6, or 9 slides per page.

### C#

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using PowerPointDocument presentation = PowerPointDocument.Load("training.pptx");

presentation.SaveAsPdf("training-notes.pdf", new PowerPointPdfExportOptions {
    PageLayout = PowerPointPdfPageLayout.NotesPages
});

presentation.SaveAsPdf("training-handout.pdf", new PowerPointPdfExportOptions {
    PageLayout = PowerPointPdfPageLayout.Handouts,
    HandoutSlidesPerPage = 6,
    IncludeHiddenSlides = false
});
```

### PowerShell

```powershell
$notes = New-OfficePowerPointPdfOptions -PageLayout NotesPages
Convert-OfficeDocument -Path .\training.pptx -OutputPath .\training-notes.pdf -Options $notes

$handout = New-OfficePowerPointPdfOptions -PageLayout Handouts -HandoutSlidesPerPage 6
Convert-OfficeDocument -Path .\training.pptx -OutputPath .\training-handout.pdf -Options $handout
```

## Convert a presentation to HTML

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
OfficeConversionReport report = presentation.SaveAsHtml("quarterly-review.html");
Console.WriteLine($"Wrote {report.OutputPath}");
```

`presentation.ToHtml()` returns the same markup as a string.

## Convert a presentation to OpenDocument (ODP)

### C#

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;

using PowerPointDocument presentation = PowerPointDocument.Load("quarterly-review.pptx");
presentation.SaveAsOpenDocument("quarterly-review.odp");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\quarterly-review.pptx -OutputPath .\quarterly-review.odp
```

### CLI

```sh
officeimo convert quarterly-review.pptx quarterly-review.odp
officeimo convert proposal.odp proposal.pptx
```
