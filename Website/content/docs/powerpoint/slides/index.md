---
title: Add Slides, Text, Images, and Notes to PowerPoint
description: Add slides with titles, formatted text, bullets, pictures, and speaker notes, and duplicate, import, reorder, hide, or group slides into sections with OfficeIMO.PowerPoint.
order: 31
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - Add-OfficePowerPointSlide
  - Add-OfficePowerPointTextBox
  - Add-OfficePowerPointList
  - Add-OfficePowerPointImage
  - Set-OfficePowerPointNotes
  - Copy-OfficePowerPointSlide
  - Import-OfficePowerPointSlide
search_terms:
  - powerpoint slide
  - speaker notes
  - pptx bullets
  - slide text box
  - reorder slides
verified: 2026-11-14
---

`PowerPointSlide` is where slide content goes. Text boxes, pictures, tables, and charts are placed by position and size. `OfficeLength` states the unit (`Centimeters`, `Points`, `Inches`, `Emus`), and `PowerPointLayoutBox` bundles a position and size so you can reuse one region. For tables and charts see [PowerPoint charts and tables](/docs/powerpoint/charts-tables/).

## Add a title and a text box

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("results.pptx");
PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeLength.Centimeters(1.5));

PowerPointSlide slide = presentation.AddSlide();
PowerPointTextBox title = slide.AddTitle("Quarterly results", new PowerPointLayoutBox(
    content.Left, content.Top, content.Width, OfficeLength.Centimeters(1.4)));
PowerPointTextStyle.Title.WithColor(OfficeColor.Parse("1F4E79")).Apply(title.Paragraphs[0]);

PowerPointTextBox intro = slide.AddTextBox("Generated from the reporting pipeline.", new PowerPointLayoutBox(
    content.Left, content.Top + OfficeLength.Centimeters(1.8), content.Width, OfficeLength.Centimeters(1)));
intro.ApplyTextStyle(PowerPointTextStyle.Body);

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\results.pptx {
    Add-OfficePowerPointSlide -Title 'Quarterly results' {
        Add-OfficePowerPointTextBox -Text 'Generated from the reporting pipeline.' -X 2cm -Y 4cm -Width 20cm -Height 2cm
    }
}
```

## Format text in a paragraph

`AddText` and `AddFormattedText` return the run they create, so formatting applies to that run only. Paragraph-level settings such as alignment and spacing go in the `AddParagraph` callback.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("formatted.pptx");
PowerPointSlide slide = presentation.AddSlide();

PowerPointTextBox box = slide.AddTextBox(string.Empty, OfficeLength.Centimeters(2), OfficeLength.Centimeters(2),
    OfficeLength.Centimeters(20), OfficeLength.Centimeters(6));
box.SetTextMargins(OfficeLength.Centimeters(0.3), OfficeLength.Centimeters(0.2),
    OfficeLength.Centimeters(0.3), OfficeLength.Centimeters(0.2));
box.TextAutoFit = PowerPointTextAutoFit.Normal;

PowerPointParagraph heading = box.AddParagraph("Highlights", p => {
    p.Alignment = PowerPointTextAlignment.Left;
    p.SpaceAfter = OfficeLength.Points(6);
});
PowerPointTextStyle.Subtitle.Apply(heading);

PowerPointParagraph line = box.AddParagraph();
line.AddText("Ship ");
line.AddFormattedText("faster", bold: true).SetColor(OfficeColor.Parse("C00000"));
line.AddText(" with generated decks.");

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\formatted.pptx {
    Add-OfficePowerPointSlide -Title 'Highlights' {
        Add-OfficePowerPointTextBox -X 2cm -Y 4cm -Width 20cm -Height 4cm {
            Add-OfficePowerPointParagraph {
                Add-OfficePowerPointText 'Ship '
                Add-OfficePowerPointText 'faster' -Bold -Color '#C00000'
                Add-OfficePowerPointText ' with generated decks.'
            }
        }
    }
}
```

## Add bullets and numbered items

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("agenda.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Agenda");

PowerPointTextBox list = slide.AddTextBox(string.Empty, OfficeLength.Centimeters(2), OfficeLength.Centimeters(4),
    OfficeLength.Centimeters(20), OfficeLength.Centimeters(8));
list.AddBullet("Release completed on schedule");
list.AddBullet("Customer guide updated");
list.AddNumberedItem("Review adoption metrics");
list.AddNumberedItem("Prepare the next milestone");
list.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfter: OfficeLength.Points(2));

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\agenda.pptx {
    Add-OfficePowerPointSlide -Title 'Agenda' {
        Add-OfficePowerPointList -Items 'Release completed on schedule', 'Customer guide updated'
        Add-OfficePowerPointList -Numbered -Items 'Review adoption metrics', 'Prepare the next milestone'
    }
}
```

## Add a picture

Give pictures alternative text so screen readers and exported PDFs describe them.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("evidence.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Deployment health");

PowerPointImage picture = slide.AddImage("dashboard.png", OfficeLength.Centimeters(2), OfficeLength.Centimeters(4),
    OfficeLength.Centimeters(16), OfficeLength.Centimeters(9));
picture.AltText = "Dashboard showing all regions healthy";

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\evidence.pptx {
    Add-OfficePowerPointSlide -Title 'Deployment health' {
        Add-OfficePowerPointImage -Path .\dashboard.png -X 2cm -Y 4cm -Width 16cm -Height 9cm -AlternativeText 'Dashboard showing all regions healthy'
    }
}
```

## Add speaker notes

`Notes.Text` sets the notes for a slide, and `Notes.TryGetText` reads them back without creating an empty notes page.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Create("review.pptx");
PowerPointSlide slide = presentation.AddSlide();
slide.AddTitle("Service review");
slide.Notes.Text = "Pause for questions before discussing capacity.";

foreach (PowerPointSlide s in presentation.Slides) {
    if (s.Notes.TryGetText(out string notes)) Console.WriteLine(notes);
}

presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\review.pptx {
    Add-OfficePowerPointSlide -Title 'Service review' {
        Add-OfficePowerPointParagraph 'All regions healthy'
        Set-OfficePowerPointNotes 'Pause for questions before discussing capacity.'
    }
}

$deck = Open-OfficePowerPoint -Path .\review.pptx
Get-OfficePowerPointNotes -Document $deck | Select-Object SlideIndex, Text
Close-OfficePowerPoint -Document $deck
```

## Duplicate, import, reorder, and hide slides

Slide indexes start at 0 in C#. `ImportSlide` copies a slide from another presentation together with its layout and pictures.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument library = PowerPointDocument.Load("slide-library.pptx");
using PowerPointDocument presentation = PowerPointDocument.Load("review.pptx");

PowerPointSlide copy = presentation.DuplicateSlide(0);
copy.Hidden = true;

presentation.ImportSlide(library, 2, insertAt: 1);
presentation.MoveSlide(presentation.Slides.Count - 1, 0);
presentation.AddSection("Appendix", presentation.Slides.Count - 1);

presentation.Save();
```

### PowerShell

PowerShell slide numbers start at 1.

```powershell
$deck = Open-OfficePowerPoint -Path .\review.pptx
Copy-OfficePowerPointSlide -Document $deck -Slide 1 -Position 3
Import-OfficePowerPointSlide -Document $deck -Path .\slide-library.pptx -SourceSlide 3 -Position 2
Add-OfficePowerPointSection -Document $deck -Name Appendix -Slide 5
Save-OfficePowerPoint -Document $deck
Close-OfficePowerPoint -Document $deck
```
