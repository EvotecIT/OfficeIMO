---
title: Create and Edit PowerPoint Presentations in .NET
description: Create, open, edit, and save PPTX and PowerPoint 97-2003 presentations with OfficeIMO.PowerPoint, PSWriteOffice, and the officeimo CLI, without PowerPoint installed.
order: 30
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - New-OfficePowerPoint
  - Add-OfficePowerPointSlide
  - Open-OfficePowerPoint
  - Update-OfficePowerPointText
search_terms:
  - pptx
  - powerpoint presentation
  - create slides
  - ppt
  - slide deck
verified: 2026-11-14
---

`OfficeIMO.PowerPoint` creates and edits `.pptx` presentations, and opens binary `.ppt`, `.pot`, and `.pps` files into the same object model. `PowerPointDocument` is the presentation, `PowerPointSlide` is one slide, and slides hold text boxes, tables, charts, pictures, and speaker notes. Nothing needs PowerPoint installed, so it runs on servers, in containers, and in CI.

Use this page for the basics. The guides in this section cover slide content, charts and tables, designer decks, corporate templates, and PDF and image export.

## Create a presentation

`SlideSize.GetContentBox` returns the slide area inside a margin, so titles and body text line up across slides.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("status-update.pptx");
PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeLength.Centimeters(1.5));

PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Weekly status", new PowerPointLayoutBox(
    content.Left, content.Top, content.Width, OfficeLength.Centimeters(1.4)));

PowerPointTextBox body = slide.AddTextBox(string.Empty, new PowerPointLayoutBox(
    content.Left, content.Top + OfficeLength.Centimeters(2),
    content.Width, content.Height - OfficeLength.Centimeters(2)));
body.AddBullets(new[] {
    "Deployment completed",
    "Onboarding guide updated",
    "Response times improved week over week"
});

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\status-update.pptx {
    Add-OfficePowerPointSlide -Title 'Weekly status' {
        Add-OfficePowerPointList -Items 'Deployment completed', 'Onboarding guide updated', 'Response times improved week over week'
    }
}
```

## Open a presentation and replace text

`Load` opens `.pptx`, `.pptm`, `.potx`, and the binary `.ppt` family. `ReplaceAll` returns the number of replacements.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("account-review.pptx");
int replaced = presentation.ReplaceAll("{{Customer}}", "Northwind Traders");
Console.WriteLine($"{replaced} replacements on {presentation.Slides.Count} slides");
presentation.Save();
```

### PowerShell

```powershell
Open-OfficePowerPoint -Path .\account-review.pptx {
    Update-OfficePowerPointText '{{Customer}}' 'Northwind Traders' -IncludeNotes
}
```

## Save a PowerPoint 97-2003 file

Saving to `.ppt`, `.pot`, or `.pps` writes the binary format and blocks known loss by default. Run the write preflight first to see what would not survive.

### C#

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;

using PowerPointDocument presentation = PowerPointDocument.Load("legacy-deck.ppt");
presentation.ReplaceAll("Draft", "Approved");

LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();
if (!report.CanWrite) {
    foreach (var finding in report.Findings) Console.WriteLine(finding);
    return;
}
presentation.SaveCopy("approved.ppt");
```

`LegacyPptCapabilityCatalog` lists what each binary feature supports for import, authoring, round-trip, and conversion from PPTX. Tables, charts, and SmartArt convert to static pictures in binary files only when you accept that loss explicitly.

## Turn a list of objects into a deck

`Export-OfficePowerPoint` builds a title slide, a table, and a chart from pipeline objects in one call.

### PowerShell

```powershell
$sales = Import-Csv .\sales.csv
$sales | Export-OfficePowerPoint -Path .\Sales.pptx -Title 'Regional sales' -Property Region, Revenue -Category Region -Value Revenue -Theme Boardroom
```
