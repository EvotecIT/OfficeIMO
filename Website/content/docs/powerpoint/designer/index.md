---
title: Designer Decks
description: Build visually structured PowerPoint decks with design briefs, semantic deck plans, recommendations, and raw composition primitives.
order: 32
---

# Designer Decks

Designer decks sit between fully manual slide placement and fixed templates. You describe brand, purpose, content, and preferences; OfficeIMO.PowerPoint can propose deterministic visual directions, explain why one direction fits, and render editable slides from semantic content.

Use this when a generated deck should feel designed without forcing every caller into the same layout, color mix, or slide rhythm.

## Design brief recommendations

A `PowerPointDesignBrief` packages brand color, stable seed, purpose, identity, palette preferences, and creative preferences. Callers can inspect alternatives before choosing one.

![PowerPoint design brief alternatives](/images/powerpoint/examples/design-brief-alternatives.png)

```csharp
PowerPointDesignBrief brief = PowerPointDesignBrief
    .FromBrand("#008C95", "design-brief-recommendations", "technical rollout proposal")
    .WithIdentity("Client Theme", eyebrow: "OfficeIMO.PowerPoint",
        footerLeft: "OFFICEIMO", footerRight: "Design brief")
    .WithPaletteStyle(PowerPointPaletteStyle.SplitComplementary)
    .WithPalette(secondaryAccentColor: "#6D5BD0", warmAccentColor: "#FFB000")
    .WithLayoutStrategy(PowerPointAutoLayoutStrategy.ContentFirst)
    .WithVariety(PowerPointDesignVariety.Exploratory)
    .WithPreferredMoods(PowerPointDesignMood.Energetic, PowerPointDesignMood.Editorial)
    .WithPreferredVisualStyles(PowerPointVisualStyle.Geometric, PowerPointVisualStyle.Soft);

var recommendations = brief.RecommendAlternatives(4);
var selected = recommendations
    .OrderByDescending(recommendation => recommendation.PreferenceScore)
    .ThenBy(recommendation => recommendation.Design.Index)
    .First();

using var presentation = PowerPointPresentation.Create("design-brief.pptx");
presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

PowerPointDeckComposer deck = presentation.UseDesigner(brief, selected.Design.Index);
```

The recommendation object keeps the choice explainable: score, direction, mood, visual style, fonts, palette, and human-readable reasons are all available before slides are rendered.

`WithLayoutStrategy(...)` lets a brief steer `Auto` slide variants without turning the deck into a fixed template.
Use `ContentFirst` for content-fit defaults, `DesignFirst` for more seeded variation, `Compact` for denser business decks,
or `VisualFirst` when visual proof and hero compositions should win when the content allows it.

![PowerPoint selected direction slide](/images/powerpoint/examples/design-brief-selected.png)

## Variation controls

The same `PowerPointDeckPlan` can be reused with several briefs. Change the layout strategy and palette style to get a
different rhythm without changing slide coordinates or duplicating every slide recipe.

```csharp
PowerPointDeckPlan plan = CreateReusablePlan();

foreach (PowerPointAutoLayoutStrategy strategy in new[] {
    PowerPointAutoLayoutStrategy.ContentFirst,
    PowerPointAutoLayoutStrategy.Compact,
    PowerPointAutoLayoutStrategy.VisualFirst
}) {
    PowerPointDesignBrief variant = PowerPointDesignBrief
        .FromBrand("#008C95", "layout-strategy-comparison", "service proposal")
        .WithPaletteStyle(PowerPointPaletteStyle.SplitComplementary)
        .WithLayoutStrategy(strategy);

    var preview = variant.DescribeDeckPlan(plan);
    PowerPointDeckComposer deck = presentation.UseDesigner(variant);
    deck.AddSlides(plan);
}
```

Run the comparison deck from the examples project:

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-layout-strategy
```

![PowerPoint layout strategy comparison](/images/powerpoint/examples/layout-strategy-comparison.png)

## Semantic deck plan

A `PowerPointDeckPlan` describes the story: section, case study, process, cards, coverage, capability, or custom slides. The plan can be scored across design alternatives before rendering.

```csharp
PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddSection("Service proposal",
        "A semantic plan keeps the story reusable while the design can change.")
    .AddCaseStudy("Managed workplace rollout",
        new[] {
            new PowerPointCaseStudySection("Client", "A distributed organization needed a clear service story."),
            new PowerPointCaseStudySection("Challenge", "Many locations and mixed hardware made delivery difficult."),
            new PowerPointCaseStudySection("Solution", "Standardized onboarding, monitoring, and operating roles."),
            new PowerPointCaseStudySection("Result", "Outcomes, metrics, and visual emphasis stay editable.")
        },
        new[] {
            new PowerPointMetric("18", "sites"),
            new PowerPointMetric("420", "devices")
        })
    .AddProcess("Implementation path",
        "The plan describes content intent; the chosen design handles layout.",
        new[] {
            new PowerPointProcessStep("Discover", "Collect constraints and service expectations."),
            new PowerPointProcessStep("Design", "Choose target architecture and rollout rules."),
            new PowerPointProcessStep("Pilot", "Validate the model with a controlled user group."),
            new PowerPointProcessStep("Roll out", "Deliver in waves with clear reporting."),
            new PowerPointProcessStep("Operate", "Move into repeatable support and optimization.")
        });

var alternatives = brief.DescribeDeckPlanAlternatives(plan, 4);
var selectedPlan = alternatives
    .OrderByDescending(alternative => alternative.ContentFitScore)
    .ThenBy(alternative => alternative.Index)
    .First();

PowerPointDeckComposer deck = presentation.UseDesigner(brief, selectedPlan.Index);
deck.AddSlides(plan);
```

![PowerPoint deck plan process slide](/images/powerpoint/examples/deck-plan-process.png)

When a deck already contains slides, preview through the active composer so fallback seeds line up with the render path:

```csharp
PowerPointDeckComposer deck = presentation.UseDesigner(brief, selectedPlan.Index);
var livePreview = deck.DescribeSlides(plan);
deck.AddSlides(plan);
```

## Raw composition still matters

Semantic plans are not meant to remove control. Use `ComposeSlide` when a slide needs custom structure, then reuse the same design primitives for title, cards, metric strips, callout bands, coverage maps, and visual frames.

```csharp
deck.ComposeSlide(composer => {
    composer.AddTitle("Why this alternative wins", selectedPlan.Design.DirectionName);

    PowerPointLayoutBox[] columns = composer.ContentColumns(2, 0.8, topCm: 3.85);
    composer.AddCardGrid(
        selectedPlan.ContentFitReasons.Take(4)
            .Select((reason, index) => new PowerPointCardContent(
                "Fit signal " + (index + 1), new[] { reason })),
        columns[0]);

    composer.AddMetricStrip(new[] {
        new PowerPointMetric(selectedPlan.ContentFitScore.ToString(), "fit score"),
        new PowerPointMetric(selectedPlan.Slides.Count.ToString(), "planned slides"),
        new PowerPointMetric(selectedPlan.Diagnostics.Count.ToString(), "diagnostics")
    }, columns[1].TakeTopCm(2.2));
}, "advisor-summary");
```

![PowerPoint deck plan advisor summary](/images/powerpoint/examples/deck-plan-advisor-summary.png)

## Runnable examples

The runnable examples generate the same decks used for the screenshots above.

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-design-brief
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-deck-plan
```

Use `--powerpoint` to run the full PowerPoint example set and validate every generated deck.

## Choosing the right level

| Level | Use it when | Main API |
|-------|-------------|----------|
| Raw slide | Exact placement or low-level Open XML behavior matters | `PowerPointSlide` |
| Composer primitives | One slide needs custom structure but should share deck styling | `PowerPointSlideComposer` |
| Semantic slide | The content has a known shape such as process, case study, or card grid | `PowerPointDeckComposer` |
| Deck plan | The caller wants to describe the whole story and choose a fitting design | `PowerPointDeckPlan` |
| Design brief | Brand, purpose, palette, and preferences should travel together | `PowerPointDesignBrief` |
| Auto layout strategy | Auto variants should lean toward content fit, seeded variation, compactness, or visual proof | `PowerPointAutoLayoutStrategy` |
