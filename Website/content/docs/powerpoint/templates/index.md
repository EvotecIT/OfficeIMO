---
title: "Create PowerPoint Decks from Templates"
description: "Inspect a corporate PPTX or POTX template, pick layouts and placeholders by name, create decks from it, and reuse its brand colors and fonts."
order: 34
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - New-OfficePowerPoint
  - Get-OfficePowerPointLayout
  - Get-OfficePowerPointPlaceholder
  - Set-OfficePowerPointPlaceholder
  - Set-OfficePowerPointTheme
limits: OfficeIMO copies the template's masters, layouts, and theme as they are; it does not redesign a template or rebuild one from screenshots.
search_terms:
  - potx
  - powerpoint template
  - slide master
  - layout placeholder
  - brand kit
verified: 2026-11-14
---

Use a template when your organization already owns the slide master, layouts, theme, logos, and footer rules. `PowerPointTemplate` reads a `.pptx` or `.potx`, tells you what layouts and placeholders it has, and creates a new presentation that keeps all of that. You then add slides by layout name instead of by index, so the code keeps working when the template gains a layout.

## Inspect a template

`Inspect` returns masters, layouts, placeholders with their roles and bounds, theme colors and fonts, logos, footer content, and a content safe area.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect("Corporate.potx");

foreach (PowerPointTemplateMasterInfo master in inventory.Masters) {
    Console.WriteLine($"{master.Name}: {master.Layouts.Count} layouts");
    foreach (PowerPointTemplateLayoutInfo layout in master.Layouts) {
        Console.WriteLine($"  {layout.Name} ({layout.Placeholders.Count} placeholders)");
    }
}
```

### PowerShell

```powershell
$deck = Open-OfficePowerPoint -Path .\Corporate.potx
Get-OfficePowerPointLayout -Document $deck
Get-OfficePowerPointTheme -Document $deck
Close-OfficePowerPoint -Document $deck
```

## Create a deck from a template

`ResolveLayout` and `ResolvePlaceholder` match by name and role. When two layouts or placeholders match, they throw `PowerPointTemplateResolutionException` with the candidate names, so a pipeline can ask for an exact choice instead of guessing. `SlideRetention` decides whether the template's own sample slides are kept.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect("Corporate.potx");
PowerPointTemplateLayoutInfo summary = inventory.ResolveLayout("Executive Summary");
PowerPointTemplatePlaceholderInfo visual = summary.ResolvePlaceholder(PowerPointTemplatePlaceholderRole.Image);

using PowerPointDocument presentation = PowerPointTemplate.CreateDocument(
    "Corporate.potx",
    "Quarterly Review.pptx",
    new PowerPointTemplateCreationOptions {
        SlideRetention = PowerPointTemplateSlideRetention.None
    });

PowerPointSlide slide = presentation.AddSlide(summary);
if (visual.Bounds is PowerPointLayoutBox bounds) {
    slide.AddImage("evidence.png", bounds);
}
presentation.Save();
```

### PowerShell

```powershell
New-OfficePowerPoint -Path .\QuarterlyReview.pptx -Template .\Corporate.potx {
    Add-OfficePowerPointSlide -TemplateLayout 'Title and Content' {
        Set-OfficePowerPointPlaceholder -Type Title 'Quarterly review'
        Set-OfficePowerPointPlaceholder -Type Body {
            Add-OfficePowerPointParagraph 'All regions are operational.'
        }
    }
}
```

## Map semantic slides to template layouts

A [designer deck plan](/docs/powerpoint/designer/) describes slides by intent (section, process, capability). `PowerPointTemplateLayoutMap` sends each kind to a named template layout, and `CreateDesignBrief` carries the template's colors and fonts into the composition. Keep `ApplyTheme` off to leave the template's theme untouched.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect("Corporate.potx");

var map = new PowerPointTemplateLayoutMap()
    .Map(PowerPointDeckPlanSlideKind.Section, inventory, "Title")
    .Map(PowerPointDeckPlanSlideKind.Process, inventory, "Process");

var plan = new PowerPointDeckPlan()
    .AddSection("Delivery review", "Generated from the corporate template")
    .AddProcess("Implementation", null, new[] {
        new PowerPointProcessStep("Discover", "Collect constraints."),
        new PowerPointProcessStep("Design", "Choose the target model."),
        new PowerPointProcessStep("Roll out", "Deliver in waves.")
    });

using PowerPointDocument presentation = PowerPointTemplate.CreateDocument("Corporate.potx", "Delivery.pptx",
    new PowerPointTemplateCreationOptions { SlideRetention = PowerPointTemplateSlideRetention.None });

PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(
    inventory.CreateDesignBrief("qbr-seed", "quarterly business review"));
composition.TemplateLayouts = map;
composition.ApplyTheme = false;
presentation.Compose(plan, composition);
presentation.Save();
```

## Apply a template's brand to another deck

`ApplyBrandTo` copies the template's color and font tokens into every master of an existing presentation without copying its layouts.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect("Corporate.potx");

using PowerPointDocument presentation = PowerPointDocument.Load("partner-deck.pptx");
inventory.ApplyBrandTo(presentation);
presentation.Save();
```

### PowerShell

```powershell
$deck = Open-OfficePowerPoint -Path .\partner-deck.pptx
Set-OfficePowerPointTheme -Document $deck -Colors @{ Accent1 = '#0F766E' } -HeadingFont 'Segoe UI Semibold' -BodyFont 'Segoe UI' -AllMasters
Save-OfficePowerPoint -Document $deck
Close-OfficePowerPoint -Document $deck
```
