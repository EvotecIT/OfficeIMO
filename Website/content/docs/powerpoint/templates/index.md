---
title: PowerPoint Templates and Brand Kits
description: Inspect corporate PPTX and POTX templates, resolve named layouts and placeholders, and generate editable branded decks.
order: 34
---

Use a template workflow when the organization already owns the slide master, layouts, theme, logos, and footer rules. OfficeIMO copies that package structure and adds editable slides; it does not rebuild a corporate template from screenshots or a parallel JSON theme.

## Inspect before generating

```csharp
PowerPointTemplateInventory inventory =
    PowerPointTemplate.Inspect("Corporate.potx");

foreach (PowerPointTemplateMasterInfo master in inventory.Masters)
{
    Console.WriteLine($"{master.Name}: {master.Layouts.Count} layouts");
}

PowerPointTemplateLayoutInfo summary =
    inventory.ResolveLayout("Executive Summary");

PowerPointTemplatePlaceholderInfo visual =
    summary.ResolvePlaceholder(PowerPointTemplatePlaceholderRole.Image);
```

The inventory includes:

- master and layout names, indexes, and native layout types;
- placeholder names, indexes, inferred roles, default text, and bounds;
- theme colors and major/minor fonts;
- master/layout pictures and likely logo identification;
- distinct footer content, slide size, title area, and a derived content safe area.

Semantic lookup is deliberately strict. If two layouts or placeholders match, `PowerPointTemplateResolutionException` reports a stable code and the candidate names so a pipeline or UI can ask for a precise selection.

## Create from PPTX or POTX

```csharp
using var presentation = PowerPointTemplate.CreatePresentation(
    "Corporate.potx",
    "Quarterly Review.pptx",
    new PowerPointTemplateCreationOptions {
        SlideRetention = PowerPointTemplateSlideRetention.None
    });

PowerPointSlide slide = presentation.AddSlide(summary);
slide.AddPicture("evidence.png", visual.Bounds!.Value);
presentation.Save();
```

`All`, `None`, and `Selected` retention modes are explicit. Selected source slides can be hidden for appendix or reference use. Masters, layouts, themes, and their assets remain in the copied package even when all source slides are removed.

## Render semantic plans into named layouts

```csharp
var map = new PowerPointTemplateLayoutMap()
    .Map(PowerPointDeckPlanSlideKind.Section, inventory, "Title")
    .Map(PowerPointDeckPlanSlideKind.Process, inventory, "Process")
    .Map(PowerPointDeckPlanSlideKind.Capability, inventory, "Executive Summary");

var plan = new PowerPointDeckPlan()
    .AddSection("Delivery review", "Generated from the corporate template")
    .AddProcess("Implementation", null, steps)
    .AddCapability("Operating model", null, capabilities);

PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(
    inventory.CreateDesignBrief("qbr-seed", "quarterly business review"));
composition.TemplateLayouts = map;
composition.ApplyTheme = false;
presentation.Compose(plan, composition);
```

The template remains the package owner. `CreateDesignBrief(...)` imports its brand tokens while `TemplateLayouts` maps semantic intent to named native layouts. Keeping `ApplyTheme` false preserves the copied corporate theme; callers can opt into applying generated theme tokens explicitly.

## Brand reuse without copying the template

```csharp
PowerPointDesignBrief brief = inventory.CreateDesignBrief(
    "campaign-seed",
    "product launch");

inventory.ApplyBrandTo(anotherPresentation);
```

The first call maps theme colors, fonts, name, and footer identity into designer inputs. The second applies the imported color and font tokens to native theme parts across all target masters.
