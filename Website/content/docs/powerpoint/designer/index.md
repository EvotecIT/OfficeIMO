---
title: "Design PowerPoint Decks from a Brief"
description: "Generate designed, editable PowerPoint decks from a brand brief and a deck plan, compare design alternatives, and add slides with composition presets."
order: 33
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - New-OfficePowerPoint
  - Add-OfficePowerPointCard
  - Add-OfficePowerPointStep
  - Add-OfficePowerPointKpi
limits: Design scoring and rhythm checks are advisory; InspectPreflight is the check for clipped text, out-of-bounds shapes, and missing pictures.
search_terms:
  - powerpoint design
  - generate slides automatically
  - deck plan
  - brand colors
  - slide layout variants
verified: 2026-11-14
---

Designer decks sit between placing every shape by hand and filling a fixed template. A `PowerPointDesignBrief` holds the brand color, a stable seed, the deck's purpose, and style preferences. A `PowerPointDeckPlan` holds the story as semantic slides: a section, a process, a case study, cards, metrics. `Compose` renders the plan with a design that fits the content, and every slide stays editable in PowerPoint. The same seed always gives the same result.

Screenshots on this page come from the runnable examples in `OfficeIMO.Examples` (`--powerpoint-design-brief`, `--powerpoint-deck-plan`, `--powerpoint-layout-strategy`).

## Compose a deck from a brief and a plan

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointDesignBrief brief = PowerPointDesignBrief
    .FromBrand("#008C95", "client-demo", "technical rollout proposal")
    .WithIdentity("Client Theme", footerLeft: "CLIENT", footerRight: "Service deck");

PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddSection("Service proposal", "Generated from a semantic plan.")
    .AddProcess("Implementation path", "The selected design handles layout.", new[] {
        new PowerPointProcessStep("Discover", "Collect constraints."),
        new PowerPointProcessStep("Design", "Choose the target model."),
        new PowerPointProcessStep("Roll out", "Deliver in waves.")
    })
    .AddCardGrid("Scope", "What the service covers.", new[] {
        new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot" }),
        new PowerPointCardContent("Care", new[] { "Monitoring", "Reporting" })
    });

using PowerPointDocument presentation = PowerPointDocument.Create("proposal.pptx");
presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
presentation.Compose(plan, PowerPointCompositionOptions.FromBrief(brief));
presentation.Save();
```

### PowerShell

The PowerShell designer cmdlets build the same kinds of slides from a deck style and purpose.

```powershell
New-OfficePowerPoint -Path .\proposal.pptx -Theme Modern -Accent '#008C95' -Purpose 'technical rollout proposal' {
    Add-OfficePowerPointSlide -Title 'Service proposal' -Subtitle 'Generated from PowerShell'
    Add-OfficePowerPointSlide -Title 'Implementation path' {
        Add-OfficePowerPointStep 'Discover' -Text 'Collect constraints.'
        Add-OfficePowerPointStep 'Design' -Text 'Choose the target model.'
        Add-OfficePowerPointStep 'Roll out' -Text 'Deliver in waves.'
    }
}
```

<a href="/images/powerpoint/examples/deck-plan-process.png" data-pf-media><img src="/images/powerpoint/examples/deck-plan-process.png" alt="Process slide rendered from a deck plan" width="1600" height="900" loading="lazy" decoding="async" /></a>

## Compare design alternatives before rendering

`DescribeDeckPlanAlternatives` scores several design directions against the plan without writing slides. Each alternative reports a content-fit score and the reasons behind it, so you can pick one or show the choice to a user.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointDesignBrief brief = PowerPointDesignBrief
    .FromBrand("#008C95", "client-demo", "technical rollout proposal")
    .WithVariety(PowerPointDesignVariety.Exploratory);

PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddSection("Service proposal", "Generated from a semantic plan.")
    .AddProcess("Implementation path", null, new[] {
        new PowerPointProcessStep("Discover", "Collect constraints."),
        new PowerPointProcessStep("Pilot", "Validate with a small group."),
        new PowerPointProcessStep("Roll out", "Deliver in waves.")
    });

var best = brief.DescribeDeckPlanAlternatives(plan, 3)
    .OrderByDescending(alternative => alternative.ContentFitScore)
    .First();
Console.WriteLine($"{best.Design.DirectionName}: {best.ContentFitScore}");

using PowerPointDocument presentation = PowerPointDocument.Create("proposal.pptx");
PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(brief);
composition.SelectBestAlternative = false;
composition.AlternativeIndex = best.Index;
presentation.Compose(plan, composition);
presentation.Save();
```

<a href="/images/powerpoint/examples/design-brief-alternatives.png" data-pf-media><img src="/images/powerpoint/examples/design-brief-alternatives.png" alt="Design brief alternatives" width="1600" height="900" loading="lazy" decoding="async" /></a>

## Steer the look with packs and strategies

A creative direction pack sets recipe, palette style, layout strategy, and ranking preferences in one call: `Boardroom`, `FieldProof`, `EditorialCaseStudy`, `TechnicalMap`, or `QuietAppendix`. `WithLayoutStrategy` steers automatic slide variants: `ContentFirst`, `DesignFirst`, `Compact`, or `VisualFirst`. Override single palette values where needed.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointDesignBrief brief = PowerPointDesignBrief
    .FromBrand("#008C95", "field-proof", "service proposal")
    .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof)
    .WithLayoutStrategy(PowerPointAutoLayoutStrategy.Compact)
    .WithPalette(surfaceColor: "#F6FAFC", panelBorderColor: "#D5E3EA");

foreach (var recommendation in brief.RecommendAlternatives(4)) {
    Console.WriteLine($"{recommendation.Design.Index}: {recommendation.PreferenceScore}");
}
```

`PowerPointDesignBrief.DescribeCreativeDirectionPacks()` lists every pack with its description and settings.

<a href="/images/powerpoint/examples/layout-strategy-comparison.png" data-pf-media><img src="/images/powerpoint/examples/layout-strategy-comparison.png" alt="The same plan rendered with three layout strategies" width="1280" height="720" loading="lazy" decoding="async" /></a>

## Use the semantic story slides

The plan has slide kinds for the parts of a deck that usually need hand-built layouts. Each has two variants, and `Auto` picks one from the seed and the content.

| Slide kind | Variants | Output |
|---|---|---|
| Executive summary | `MetricLead`, `DecisionBrief` | Text, metric strips, and cards |
| Chart story | `ChartHero`, `InsightRail` | Editable chart with caption, source, and alt text |
| Comparison | `SideBySide`, `DecisionMatrix` | Cards or an editable table |
| Screenshot story | `HeroAnnotated`, `SplitNarrative` | Picture with focal crop, caption, and annotations |
| Appendix table | `FullWidth`, `NotesRail` | Editable table that continues on new slides |
| Architecture | `Layered`, `HubSpoke` | Editable shapes, labels, and connectors |
| Closing | `Statement`, `ActionPanel` | Statement or next-action slide |

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#1F4E79", "qbr", "quarterly business review");

PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddSection("Quarterly review", "Operations")
    .AddCaseStudy("Managed workplace rollout",
        new[] {
            new PowerPointCaseStudySection("Client", "A distributed organization."),
            new PowerPointCaseStudySection("Challenge", "Many sites and mixed hardware."),
            new PowerPointCaseStudySection("Solution", "Standard onboarding and monitoring."),
            new PowerPointCaseStudySection("Result", "Faster, predictable delivery.")
        },
        new[] { new PowerPointMetric("18", "sites"), new PowerPointMetric("420", "devices") })
    .AddClosing("Next decision", new PowerPointClosingContent(
        "Approve the pilot for the next two sites.", "Approve the pilot"));

using PowerPointDocument presentation = PowerPointDocument.Create("qbr.pptx");
PowerPointCompositionResult result = presentation.Compose(plan, PowerPointCompositionOptions.FromBrief(brief));

PowerPointDeckRhythmReport rhythm = result.Plan.InspectRhythm(result.Design);
foreach (var finding in rhythm.Findings) Console.WriteLine($"{finding.Code}: {finding.Message}");

presentation.Save();
```

The rhythm report flags repeated slide kinds, dense streaks, long sections, and a missing closing slide with stable codes such as `Rhythm.RepeatedKind` and `Rhythm.MissingClosing`.

## Add a custom slide with composition presets

`AddCustom` gives you a composer that shares the deck's styling. `UsePreset` returns named regions (`Primary`, `Visual`, `Metrics`) so the slide keeps designed spacing without hard-coded coordinates.

### C#

```csharp
using OfficeIMO.PowerPoint;

PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#008C95", "custom", "service proposal");

PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddCustom("Why this approach", composer => {
        composer.AddTitle("Why this approach", "Evidence from the pilot");
        PowerPointCompositionLayout layout = composer.UsePreset(
            PowerPointCompositionPreset.MetricStory, PowerPointCompositionVariant.VisualLead);
        composer.AddCardGrid(new[] {
            new PowerPointCardContent("Speed", new[] { "Devices ready in one day" }),
            new PowerPointCardContent("Control", new[] { "One policy baseline" })
        }, layout.Primary);
        composer.AddVisualFrame(layout.Visual, PowerPointVisualFrameVariant.ProofBoard);
        composer.AddMetricStrip(new[] {
            new PowerPointMetric("98%", "compliant"),
            new PowerPointMetric("4h", "setup time")
        }, layout.Metrics, PowerPointMetricStripVariant.SeparatedTiles);
    }, "why-this-approach");

using PowerPointDocument presentation = PowerPointDocument.Create("custom.pptx");
presentation.Compose(plan, PowerPointCompositionOptions.FromBrief(brief));
presentation.Save();
```

<a href="/images/powerpoint/examples/deck-plan-advisor-summary.png" data-pf-media><img src="/images/powerpoint/examples/deck-plan-advisor-summary.png" alt="Custom slide built with a composition preset" width="1600" height="900" loading="lazy" decoding="async" /></a>

## Choose the right level

| Level | Use it when | Main API |
|---|---|---|
| Raw slide | Exact placement matters | `PowerPointSlide` |
| Composer presets | One slide needs custom structure with designed spacing | `PowerPointCompositionPreset` |
| Semantic slide | The content has a known shape: process, case study, cards | `PowerPointDeckPlan` |
| Deck composition | Render the whole story with one fitting design | `PowerPointDocument.Compose` |
| Design brief | Brand, purpose, and preferences travel together | `PowerPointDesignBrief` |
