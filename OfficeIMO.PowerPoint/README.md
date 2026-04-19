# OfficeIMO.PowerPoint — .NET PowerPoint Utilities

OfficeIMO.PowerPoint focuses on creating and editing .pptx presentations with Open XML.

- Targets: netstandard2.0, net472, net8.0, net9.0
- License: MIT
- NuGet: `OfficeIMO.PowerPoint`
- Dependencies: DocumentFormat.OpenXml

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

See `OfficeIMO.Examples` for runnable samples. This README hosts PowerPoint‑specific usage and notes.

## Runnable modern deck sample

To generate the richer validation sample deck with theme colors, theme fonts, background image, transitions, shape effects,
charts, a table, and speaker notes:

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --modern-powerpoint
```

The sample writes `Modern PowerPoint Deck.pptx` to the examples `Documents` output folder and validates the generated Open
XML package before reporting success.

To generate the designer examples used by the website screenshots:

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-design-brief
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-deck-plan
```

The first sample demonstrates explainable design recommendations. The second demonstrates semantic deck-plan scoring
before rendering editable slides.

To run the full PowerPoint examples set without opening PowerPoint and validate every generated deck:

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint
```

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint
```

## Quick sample

```csharp
using OfficeIMO.PowerPoint;

using var ppt = PowerPointPresentation.Create("demo.pptx");
var slide = ppt.AddSlide();
slide.AddTitle("Hello PowerPoint");
var box = slide.AddTextBox("Generated with OfficeIMO.PowerPoint");
box.SetPositionCm(2, 2);
box.SetSizeCm(6, 2);
ppt.Save();
```

## Designer composition sample

For business decks where you want polished placement without manually positioning every text box, use the designer
composition helpers. They create editable PowerPoint shapes, text, optional pictures, and theme-driven layouts.
Use `PowerPointDesignIntent` and layout variants when the same content should not produce the same slide every time.

```csharp
using OfficeIMO.PowerPoint;

using var ppt = PowerPointPresentation.Create("designer-demo.pptx");
ppt.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
var alternatives = PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "client-demo",
    name: "Client Theme", eyebrow: "Client Group",
    footerLeft: "CLIENT", footerRight: "Service deck");
var design = alternatives[1]; // pick the stable Editorial creative direction for this deck
var deck = ppt.UseDesigner(design);

// Or use a recipe when you want a scenario-specific family of distinct directions.
var portfolioAlternatives = PowerPointDesignRecipe.ConsultingPortfolio.CreateAlternativesFromBrand("#008C95", "client-demo",
    name: "Client Theme", footerLeft: "CLIENT", footerRight: "Service deck");
var portfolioDesign = portfolioAlternatives[0]; // Board Story, Field Proof, Quiet Appendix, ...

// Recipes can also be selected from plain-language purpose text.
var recipeChoices = PowerPointDesignRecipe.DescribeBuiltIns(); // names, keywords, directions, fonts, and moods
var purposeMatches = PowerPointDesignRecipe.DescribeMatches("technical rollout proposal");
var recipe = PowerPointDesignRecipe.FindBuiltIn("technical rollout proposal")
    ?? PowerPointDesignRecipe.ConsultingPortfolio;

// For the shortest path, start the deck composer directly from brand and purpose text.
var quickDeck = ppt.UseDesigner("#008C95", "client-demo", "technical rollout proposal",
    name: "Client Theme", footerLeft: "CLIENT", footerRight: "Service deck");

// Use a design brief when brand, purpose, identity, and custom directions should travel together.
var brief = PowerPointDesignBrief
    .FromBrand("#008C95", "client-demo", "technical rollout proposal")
    .WithIdentity("Client Theme", footerLeft: "CLIENT", footerRight: "Service deck")
    .WithPaletteStyle(PowerPointPaletteStyle.SplitComplementary)
    .WithPalette(secondaryAccentColor: "#6D5BD0", warmAccentColor: "#FFB000")
    .WithLayoutStrategy(PowerPointAutoLayoutStrategy.ContentFirst)
    .WithVariety(PowerPointDesignVariety.Exploratory)
    .WithPreferredMoods(PowerPointDesignMood.Energetic)
    .WithPreferredVisualStyles(PowerPointVisualStyle.Geometric);
var choices = brief.DescribeAlternatives(3); // direction, mood, fonts, and palette preview
var recommendations = brief.RecommendAlternatives(3); // preference score and reasons before choosing
var briefDeck = ppt.UseDesigner(brief, alternativeIndex: 1);

// A deck plan lets callers describe the story while the designer chooses the slide compositions.
var plan = new PowerPointDeckPlan()
    .AddSection("Case Study", "Project portfolio", "cover")
    .AddCaseStudy("Example client",
        new[] {
            new PowerPointCaseStudySection("Client", "A concise customer story."),
            new PowerPointCaseStudySection("Challenge", "Many details needed structure."),
            new PowerPointCaseStudySection("Solution", "Separate story, evidence, and outcome."),
            new PowerPointCaseStudySection("Result", "Keep the output editable and readable.")
        },
        seed: "case-study")
    .AddProcess("How we work", "Transparent phases reduce risk",
        new[] {
            new PowerPointProcessStep("Analysis", "Understand the environment and constraints."),
            new PowerPointProcessStep("Discovery", "Review configuration and dependencies."),
            new PowerPointProcessStep("Delivery", "Implement changes in controlled stages.")
        },
        seed: "process")
    .AddCustom("Custom detail", composer => {
        composer.AddTitle("Custom detail", "Use raw composition when a planned slide needs something special.");
        composer.AddMetricStrip(new[] { new PowerPointMetric("2", "modes") },
            composer.ContentArea().TakeTopCm(1.5));
    }, seed: "custom-detail");
var plannedSlides = plan.DescribeSlides(); // kind, title, seed, and content count
var planDiagnostics = plan.ValidateSlides(); // density, clipping, and bounds issues before rendering
var renderPreview = brief.DescribeDeckPlan(plan, alternativeIndex: 1); // variants, layout reasons, fonts, and seeds
var planChoices = brief.DescribeDeckPlanAlternatives(plan, 3); // includes content-fit score and reasons
var recommendedPlan = brief.RecommendDeckPlanAlternative(plan, 3); // strongest content-fit choice first
var recommendedDeck = ppt.UseDesigner(brief, plan, alternativeCount: 3); // choose the recommended design
var livePreview = recommendedDeck.DescribeSlides(plan); // seed preview accounts for slides already composed in this deck
recommendedDeck.AddSlides(plan); // validates errors before rendering and keeps warnings inspectable

// Or supply your own creative directions so decks do not all share the same house style.
var clientDirections = new[] {
    new PowerPointDesignDirection("Board Brief", PowerPointDesignMood.Corporate,
        PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
        showDirectionMotif: false),
    new PowerPointDesignDirection("Field Ops", PowerPointDesignMood.Energetic,
        PowerPointSlideDensity.Compact, PowerPointVisualStyle.Geometric, "Poppins", "Segoe UI")
};
var clientAlternatives = PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "client-demo",
    clientDirections, name: "Client Theme", footerLeft: "CLIENT");
var uniqueBrief = PowerPointDesignBrief.FromBrand("#008C95", "client-demo")
    .WithIdentity("Client Theme", footerLeft: "CLIENT")
    .WithDirections(clientDirections)
    .WithPaletteStyle(PowerPointPaletteStyle.CoolNeutral)
    .WithLayoutStrategy(PowerPointAutoLayoutStrategy.Compact)
    .WithPreferredDensities(PowerPointSlideDensity.Compact);

deck.AddSectionSlide("Case Study", "Project portfolio", "cover",
    options => options.SectionVariant = PowerPointSectionLayoutVariant.EditorialRail);

deck.AddCaseStudySlide("Example client",
    new[] {
        new PowerPointCaseStudySection("Client", "A concise customer story."),
        new PowerPointCaseStudySection("Challenge", "Many details needed structure."),
        new PowerPointCaseStudySection("Solution", "Separate story, evidence, and outcome."),
        new PowerPointCaseStudySection("Result", "Keep the output editable and readable.")
    },
    seed: "case-study",
    configure: options => options.Variant = PowerPointCaseStudyLayoutVariant.EditorialSplit);

deck.AddProcessSlide("How we work", "Transparent phases reduce risk",
    new[] {
        new PowerPointProcessStep("Analysis", "Understand the environment and constraints."),
        new PowerPointProcessStep("Discovery", "Review configuration and dependencies."),
        new PowerPointProcessStep("Delivery", "Implement changes in controlled stages.")
    },
    seed: "process",
    configure: options => options.Variant = PowerPointProcessLayoutVariant.NumberedColumns);

deck.AddCardGridSlide("Scope of services", "Cards choose their own grid.",
    new[] {
        new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot" }),
        new PowerPointCardContent("Maintenance", new[] { "Incidents", "Monitoring" }),
        new PowerPointCardContent("Audits", new[] { "Configuration", "Security" })
    },
    seed: "services",
    configure: options => options.Variant = PowerPointCardGridLayoutVariant.SoftTiles);

deck.AddLogoWallSlide("Proof points", "Reusable logo and certification wall.",
    new[] {
        new PowerPointLogoItem("Lenovo", "Partner"),
        new PowerPointLogoItem("Samsung", "Devices"),
        new PowerPointLogoItem("Epson", "Service")
    },
    seed: "proof",
    configure: options => options.FeatureTitle = "Featured certification");

deck.AddCoverageSlide("Service coverage", "Pins are normalized inside the editable map panel.",
    new[] {
        new PowerPointCoverageLocation("Warszawa", 0.60, 0.48),
        new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
        new PowerPointCoverageLocation("Krakow", 0.58, 0.78)
    },
    seed: "coverage",
    configure: options => {
        options.MapLabel = "Editable locations";
    });

deck.AddCapabilitySlide("Service capability", "Structured text with visual support.",
    new[] {
        new PowerPointCapabilitySection("Warranty service",
            "Nationwide support for distributed environments.",
            new[] { "Computers and notebooks", "Printers and scanners" }),
        new PowerPointCapabilitySection("Extended care",
            "Support beyond standard vendor warranty.",
            new[] { "SLA options", "Continuity monitoring" })
    },
    seed: "capability",
    configure: options => {
        options.VisualKind = PowerPointCapabilityVisualKind.CoverageMap;
        options.VisualLabel = "Service locations";
        options.Locations.Add(new PowerPointCoverageLocation("Warszawa", 0.60, 0.48));
        options.Locations.Add(new PowerPointCoverageLocation("Gdansk", 0.55, 0.18));
        options.Metrics.Add(new PowerPointMetric("8", "locations"));
    });

deck.ComposeSlide(composer => {
    composer.AddTitle("Custom slide", "Use primitives when a recipe is too fixed.");
    var columns = composer.ContentColumns(2);
    composer.AddCardGrid(new[] {
        new PowerPointCardContent("Story", new[] { "Context", "Need" }),
        new PowerPointCardContent("Evidence", new[] { "Metrics", "Visual" })
    }, columns[0]);
    composer.AddCoverageMap(new[] {
        new PowerPointCoverageLocation("North", 0.45, 0.20),
        new PowerPointCoverageLocation("Central", 0.60, 0.48)
    }, columns[1].TakeTopCm(3.0));
    composer.AddCalloutBand("Use composer regions when the slide needs its own structure.",
        columns[1].TakeBottomCm(1.5));
}, "custom");

ppt.Save();
```

Runnable sample:

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --designer-powerpoint
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-layout-strategy
```

The helpers are intentionally not fixed templates. Start with `PowerPointDeckDesign.FromBrand(...)` to define the
deck personality once, including brand color, stable seed, mood, fonts, and chrome. Use a named
`PowerPointDesignDirection` such as `Structured`, `Editorial`, `Quiet`, `Signal`, or `Executive` when you want a
recognizable creative direction without hand-tuning every slide. Use `PowerPointDesignRecipe` values such as
`ConsultingPortfolio`, `ExecutiveBrief`, `TechnicalProposal`, or `TransformationRoadmap` when you want a
scenario-specific family of alternatives instead of one house style repeated across every client deck. The deck design
configures per-slide
`PowerPointDesignIntent` values so repeated content can receive stable but different accents, motifs, and automatic
layout choices. Auto variants use both the design intent and the content shape: dense card grids stay compact, softer
moods get softer cards, long processes stay readable, proof slides emphasize supplied certificate details, many
locations become list-plus-map slides, section-heavy capability slides stack into readable panels, and content-rich
case studies choose stronger structure. Use `PowerPointDeckDesign.CreateAlternativesFromBrand(...)` with either a count,
custom directions, or a recipe when you want stable choices from the same brand before choosing the deck personality.

Use `PowerPointDesignBrief.WithLayoutStrategy(...)` when `Auto` variants should lean toward content fit, seeded design
variety, compact business layouts, or more visual hero/proof compositions without hardcoding every slide variant.
Use `RecommendDeckPlanAlternative(...)` when callers want the library to pick the strongest content-fit alternative
from the same plan while still returning the selected design index and reasons.
Use `presentation.UseDesigner(brief, plan, alternativeCount: ...)` when the caller wants to skip manual ranking and
compose with the recommended alternative directly.
Run the layout strategy sample when you want to compare the same semantic `PowerPointDeckPlan` rendered through
different brief-level palette and layout choices.
Use explicit layout variants when a deck needs a controlled art direction, or use `ComposeDesignerSlide` and
`PowerPointLayoutBox` regions when the slide needs a custom composition while still reusing cards, metrics, process
steps, logo walls, coverage maps, and callout bands.

## Common Tasks by Example

### Title + content
```csharp
var slide = ppt.AddSlide();
slide.AddTitle("Quarterly Review");
slide.AddTextBox("Agenda\n• Intro\n• KPIs\n• Next Steps",
    PowerPointUnits.Cm(1.5), PowerPointUnits.Cm(2.5), PowerPointUnits.Cm(7.5), PowerPointUnits.Cm(3.0));
```

### Bullets (API)
```csharp
var box = slide.AddTextBox("Agenda:");
box.AddBullets(new[] { "Intro", "KPIs", "Next Steps" });
```

### Numbered lists
```csharp
var box = slide.AddTextBox("Plan:");
box.AddNumberedList(new[] { "Discover", "Design", "Deliver" });
```

### Text styles + spacing
```csharp
var box = slide.AddTextBox("Highlights");
box.AddBullets(new[] { "Readable defaults", "Auto spacing", "Consistent styles" });
box.ApplyTextStyle(PowerPointTextStyle.Body.WithColor("1F4E79"));
box.ApplyAutoSpacing(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);
```

### Text box layout (margins + autofit)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var box = slide.AddTextBox("Inset text");
box.SetTextMarginsCm(0.3, 0.2, 0.3, 0.2);
box.SetTextAutoFit(PowerPointTextAutoFit.Normal,
    new PowerPointTextAutoFitOptions(fontScalePercent: 85, lineSpaceReductionPercent: 10));
box.TextVerticalAlignment = A.TextAnchoringTypeValues.Center;
```

### Images
```csharp
slide.AddPicture("logo.png",
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2), PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));

using var logoStream = File.OpenRead("logo.png");
slide.AddPicture(logoStream, ImagePartType.Png,
    PowerPointUnits.Cm(23), PowerPointUnits.Cm(1.2), PowerPointUnits.Cm(5), PowerPointUnits.Cm(2));
```

### Create and edit with streams
```csharp
using var stream = new MemoryStream();

using (var ppt = PowerPointPresentation.Create(stream)) {
    ppt.AddSlide().AddTitle("Created in memory");
}

using (var ppt = PowerPointPresentation.Open(stream, readOnly: false, autoSave: true)) {
    ppt.AddSlide().AddTitle("Updated from the same stream");
}
```

### Background image
```csharp
slide.SetBackgroundImage("hero.png");
```

### Simple shapes
```csharp
slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
    PowerPointUnits.Cm(3), PowerPointUnits.Cm(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");
```

### Lines (dash + arrows)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var line = slide.AddLine(
    PowerPointUnits.Cm(1), PowerPointUnits.Cm(1),
    PowerPointUnits.Cm(8), PowerPointUnits.Cm(1));
line.OutlineColor = "2F5597";
line.OutlineDash = A.PresetLineDashValues.Dash;
line.SetLineEnds(A.LineEndValues.Triangle, A.LineEndValues.Stealth,
    A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);
```

### Shape effects (shadow + glow + blur)
```csharp
var card = slide.AddRectangle(PowerPointUnits.Cm(1), PowerPointUnits.Cm(4),
    PowerPointUnits.Cm(6), PowerPointUnits.Cm(2), "Card");
card.FillColor = "FFFFFF";
card.OutlineColor = "D0D0D0";
card.SetShadow("000000", blurPoints: 6, distancePoints: 4,
    angleDegrees: 270, transparencyPercent: 40);
card.SetGlow("000000", radiusPoints: 3, transparencyPercent: 60);
card.SetSoftEdges(1.5);
card.SetBlur(2, grow: false);
```

### Shape effects (reflection)
```csharp
var logo = slide.AddRectangle(PowerPointUnits.Cm(10), PowerPointUnits.Cm(4),
    PowerPointUnits.Cm(4), PowerPointUnits.Cm(2), "Logo");
logo.SetReflection(blurPoints: 4, distancePoints: 2, directionDegrees: 270,
    fadeDirectionDegrees: 90, startOpacityPercent: 60, endOpacityPercent: 0);
```

### Align + distribute shapes
```csharp
slide.AlignShapes(slide.Shapes, PowerPointShapeAlignment.Left);
slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal);   

slide.AlignShapesToSlideContent(slide.Shapes, PowerPointShapeAlignment.Left,
    marginEmus: PowerPointUnits.Cm(1));
slide.DistributeShapesToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    marginEmus: PowerPointUnits.Cm(1));
slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    PowerPointShapeAlignment.Bottom);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), center: true);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), PowerPointShapeAlignment.Right);      
slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), PowerPointShapeAlignment.Right,
    PowerPointShapeAlignment.Bottom);

slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    new PowerPointShapeSpacingOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ClampSpacingToBounds = true,
        Alignment = PowerPointShapeAlignment.Center,
        CrossAxisAlignment = PowerPointShapeAlignment.Bottom
    });
slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    new PowerPointShapeSpacingOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ScaleToFitBounds = true,
        PreserveAspect = true
    });

slide.DistributeShapesWithSpacingToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.5), marginEmus: PowerPointUnits.Cm(1));

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.3), PowerPointShapeAlignment.Bottom);
slide.StackShapesToSlideContent(slide.Shapes, PowerPointShapeStackDirection.Vertical,
    spacingEmus: PowerPointUnits.Cm(0.3), marginEmus: PowerPointUnits.Cm(1));

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.3), PowerPointShapeStackJustify.Center);

slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,       
    new PowerPointShapeStackOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ClampSpacingToBounds = true,
        Justify = PowerPointShapeStackJustify.Center
    });
slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal,
    new PowerPointShapeStackOptions {
        SpacingEmus = PowerPointUnits.Cm(0.4),
        ScaleToFitBounds = true,
        PreserveAspect = true
    });
```

### Resize shapes
```csharp
slide.ResizeShapes(slide.Shapes, PowerPointShapeSizeDimension.Width, PowerPointShapeSizeReference.Largest);
slide.ResizeShapesCm(slide.Shapes, widthCm: 3.0, heightCm: null);
```

### Shape bounds + movement
```csharp
var shape = slide.AddRectangle(0, 0, PowerPointUnits.Cm(3), PowerPointUnits.Cm(2));

shape.Right = PowerPointUnits.Cm(10);
shape.CenterX = PowerPointUnits.Cm(12);
shape.Bounds = PowerPointLayoutBox.FromCentimeters(2, 2, 4, 2);

shape.MoveBy(PowerPointUnits.Cm(0.5), PowerPointUnits.Cm(0.25));
shape.Resize(PowerPointUnits.Cm(5), PowerPointUnits.Cm(2), PowerPointShapeAnchor.Center);
shape.Scale(1.2, PowerPointShapeAnchor.BottomRight);

var inArea = slide.GetShapesInBounds(
    PowerPointLayoutBox.FromCentimeters(1, 1, 5, 3),
    includePartial: true);
```

### Grid layout
```csharp
slide.ArrangeShapesInGrid(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000),
    columns: 4, rows: 2, gutterX: 200, gutterY: 200,
    flow: PowerPointShapeGridFlow.RowMajor);

slide.ArrangeShapesInGridAuto(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000));

slide.ArrangeShapesInGridAuto(slide.Shapes,
    new PowerPointLayoutBox(0, 0, 8000, 4000),
    new PowerPointShapeGridOptions { MinColumns = 2, MaxColumns = 4, TargetCellAspect = 1.0 });

slide.ArrangeShapesInGridToSlideContent(slide.Shapes, columns: 3, rows: 2,
    marginEmus: PowerPointUnits.Cm(1), gutterX: PowerPointUnits.Cm(0.5));
```

### Fit shapes to bounds
```csharp
slide.FitShapesToBounds(slide.Shapes, new PowerPointLayoutBox(0, 0, 8000, 4500));
slide.FitShapesToSlideContentCm(slide.Shapes, marginCm: 1.0, preserveAspect: true, center: true);
```

### Group + ungroup shapes
```csharp
var group = slide.GroupShapes(slide.Shapes);
slide.UngroupShape(group);

slide.AlignGroupChildren(group, PowerPointShapeAlignment.Left);
slide.DistributeGroupChildrenWithSpacing(group, PowerPointShapeDistribution.Horizontal,
    spacingEmus: PowerPointUnits.Cm(0.4));
slide.StackGroupChildren(group, PowerPointShapeStackDirection.Vertical,
    new PowerPointShapeStackOptions { SpacingEmus = PowerPointUnits.Cm(0.3) });
slide.ArrangeGroupChildrenInGrid(group, columns: 2, rows: 2);

var groupTextBoxes = slide.GetGroupTextBoxes(group);
var groupBoundsCm = slide.GetGroupChildBoundsCm(group);
```

### Arrange + duplicate shapes
```csharp
slide.BringForward(shape);
slide.SendBackward(shape);
slide.BringToFront(shape);
slide.SendToBack(shape);

var copy = slide.DuplicateShapeCm(shape, 0.5, 0.5);
```

### Slide properties
```csharp
ppt.BuiltinDocumentProperties.Title = "Contoso Review";
ppt.ApplicationProperties.Company = "Contoso";
```

### Slide visibility + duplication
```csharp
var duplicate = ppt.DuplicateSlide(0);
duplicate.Hidden = true;
```

### Slide layouts + theme helpers
```csharp
using DocumentFormat.OpenXml.Presentation;

var layouts = ppt.GetSlideLayouts();
var titleLayoutIndex = ppt.GetLayoutIndex(SlideLayoutValues.Title);

var slide = ppt.AddSlide(SlideLayoutValues.TitleOnly);
slide.SetLayout("Title and Content");

ppt.SetThemeColor(PowerPointThemeColor.Accent1, "FF0000");
ppt.SetThemeLatinFonts("Aptos", "Calibri");
ppt.SetThemeFonts(new PowerPointThemeFontSet(
    majorLatin: "Aptos",
    minorLatin: "Calibri",
    majorEastAsian: "MS Mincho",
    minorEastAsian: "Yu Gothic",
    majorComplexScript: "Arial",
    minorComplexScript: "Tahoma"));

// Apply updates across all masters
ppt.SetThemeColorForAllMasters(PowerPointThemeColor.Accent2, "00B0F0");
ppt.SetThemeLatinFontsForAllMasters("Aptos", "Calibri");
ppt.SetThemeNameForAllMasters("Contoso Theme");
```

### Import slide from another deck
```csharp
using var source = PowerPointPresentation.Open("source.pptx");
var imported = ppt.ImportSlide(source, sourceIndex: 0);
imported.AddTextBox("Imported content");
```

### Charts (formatting)
```csharp
using C = DocumentFormat.OpenXml.Drawing.Charts;

var chart = slide.AddChart();
var labelTemplate = new PowerPointChartDataLabelTemplate {
    ShowValue = true,
    ShowCategoryName = true,
    Position = C.DataLabelPositionValues.OutsideEnd,
    NumberFormat = "0.0",
    Separator = " - ",
    TextColor = "1F4E79",
    FillColor = "FFFFFF",
    LineColor = "1F4E79",
    LineWidthPoints = 0.5
};

chart.SetTitle("Sales Trend")
     .SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1.25)
     .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "00B0F0", lineWidthPoints: 0.5)
     .SetTitleTextStyle(fontSizePoints: 18, bold: true, color: "1F4E79", fontName: "Calibri")
     .ClearTitleTextStyle()
     .SetLegend(C.LegendPositionValues.Right)
     .SetLegendTextStyle(fontSizePoints: 9, italic: true, color: "404040", fontName: "Calibri")
     .ClearLegendTextStyle()
     .SetDataLabels(showValue: true)
     .SetDataLabelPosition(C.DataLabelPositionValues.OutsideEnd)
     .SetDataLabelNumberFormat("#,##0.0", sourceLinked: false)
     .SetDataLabelTextStyle(fontSizePoints: 9, color: "1F4E79")
     .SetDataLabelShapeStyle(fillColor: "FFFFFF", lineColor: "1F4E79", lineWidthPoints: 0.5)
     .SetDataLabelLeaderLines(showLeaderLines: true, lineColor: "1F4E79", lineWidthPoints: 0.5)
     .SetDataLabelSeparator(" | ")
     .SetDataLabelTemplate(labelTemplate)
     .SetDataLabelCallouts(enabled: true, lineColor: "1F4E79", lineWidthPoints: 0.5)
     .SetSeriesDataLabelTemplate(0, labelTemplate)
     .SetSeriesDataLabelTextStyle(0, bold: true, color: "C00000")
     .SetSeriesDataLabelSeparator(0, " / ")
     .SetSeriesDataLabels(0, showValue: true, showCategoryName: true, position: C.DataLabelPositionValues.Top, numberFormat: "0.0")
     .SetSeriesDataLabelCallouts(0, enabled: true, lineColor: "C00000", lineWidthPoints: 0.75)
     .SetSeriesDataLabelForPoint(0, 1, showValue: true, showCategoryName: true, position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "0.00")
     .SetSeriesDataLabelTemplateForPoint(0, 1, labelTemplate)
     .SetSeriesDataLabelLeaderLines(0, showLeaderLines: true, lineColor: "C00000", lineWidthPoints: 0.75)
     .SetSeriesDataLabelCalloutsForPoint(0, 1, enabled: true)
     .SetSeriesDataLabelSeparatorForPoint(0, 1, " | ")
     .SetSeriesDataLabelTextStyleForPoint(0, 1, fontSizePoints: 11, bold: true, color: "C00000")
     .SetSeriesDataLabelShapeStyleForPoint(0, 1, fillColor: "FFF2CC", lineColor: "C00000", lineWidthPoints: 0.75)
     .SetSeriesTrendline(0, C.TrendlineValues.Polynomial, order: 2, lineColor: "ED7D31", lineWidthPoints: 1.5)
     .SetCategoryAxisTitle("Quarter")
     .SetCategoryAxisTitleTextStyle(fontSizePoints: 11, bold: true, color: "1F4E79", fontName: "Calibri")
     .ClearCategoryAxisTitleTextStyle()
     .SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "404040", fontName: "Calibri")
     .ClearCategoryAxisLabelTextStyle()
     .SetCategoryAxisLabelRotation(45)
     .SetCategoryAxisTickLabelPosition(C.TickLabelPositionValues.High)
     .SetCategoryAxisGridlines(showMajor: true, lineColor: "D9D9D9", lineWidthPoints: 0.5)
     .ClearCategoryAxisGridlines()
     .SetValueAxisTitle("Revenue")
     .SetValueAxisTitleTextStyle(fontSizePoints: 10, italic: true, color: "C55A11", fontName: "Arial")
     .ClearValueAxisTitleTextStyle()
     .SetValueAxisLabelTextStyle(fontSizePoints: 9, italic: true, color: "595959", fontName: "Arial")
     .ClearValueAxisLabelTextStyle()
     .SetValueAxisTickLabelPosition(C.TickLabelPositionValues.Low)
     .SetValueAxisGridlines(showMajor: true, showMinor: true, lineColor: "C0C0C0", lineWidthPoints: 0.75)
     .ClearValueAxisGridlines()
     .SetValueAxisNumberFormat("#,##0.00")
     .SetCategoryAxisReverseOrder()
     .SetValueAxisScale(minimum: 0, maximum: 100, majorUnit: 20, minorUnit: 10)
     .SetValueAxisCrossing(C.CrossesValues.Maximum)
     .SetValueAxisCrossBetween(C.CrossBetweenValues.Between)
     .SetValueAxisDisplayUnits(C.BuiltInUnitValues.Thousands, "Thousands USD", showLabel: true)
     .SetCategoryAxisCrossing(C.CrossesValues.Minimum)
     .SetSeriesFillColor(0, "4472C4")
     .SetSeriesLineColor("Series 2", "ED7D31", widthPoints: 1)
     .SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 6, fillColor: "FFFFFF", lineColor: "4472C4");

chart.ClearDataLabels();
chart.ClearSeriesDataLabels(0);
chart.ClearSeriesDataLabelForPoint(0, 1);
```

### Pie and doughnut charts
```csharp
var chartData = new PowerPointChartData(
    new[] { "North", "South", "West" },
    new[] { new PowerPointChartSeries("Revenue", new[] { 10d, 20d, 30d }) });

slide.AddPieChart(chartData)
    .SetTitle("Revenue Share")
    .SetDataLabels(showValue: true, showPercent: true);

slide.AddDoughnutChart(chartData, PowerPointUnits.Cm(10), PowerPointUnits.Cm(2),
    PowerPointUnits.Cm(8), PowerPointUnits.Cm(5))
    .SetTitle("Revenue Mix");
```

### Scatter chart axes
```csharp
var scatterData = new PowerPointScatterChartData(new[] {
    new PowerPointScatterChartSeries("Revenue", new[] { 1d, 2d, 3d, 4d }, new[] { 10d, 15d, 12d, 18d })
});

slide.AddScatterChart(scatterData)
    .SetTitle("Revenue Scatter")
    .SetScatterXAxisTitle("Month")
    .SetScatterYAxisTitle("Revenue")
    .SetScatterXAxisTitleTextStyle(fontSizePoints: 11, bold: true, color: "1F4E79", fontName: "Calibri")
    .SetScatterYAxisTitleTextStyle(fontSizePoints: 10, italic: true, color: "C55A11", fontName: "Arial")
    .ClearScatterXAxisTitleTextStyle()
    .ClearScatterYAxisTitleTextStyle()
    .SetScatterXAxisLabelTextStyle(fontSizePoints: 9, bold: true, color: "404040", fontName: "Calibri")
    .SetScatterYAxisLabelTextStyle(fontSizePoints: 10, italic: true, color: "1F4E79", fontName: "Arial")
    .ClearScatterXAxisLabelTextStyle()
    .ClearScatterYAxisLabelTextStyle()
    .SetScatterXAxisLabelRotation(45)
    .SetScatterYAxisLabelRotation(-30)
    .SetScatterXAxisTickLabelPosition(C.TickLabelPositionValues.Low)
    .SetScatterYAxisTickLabelPosition(C.TickLabelPositionValues.High)
    .SetScatterXAxisGridlines(showMajor: true, lineColor: "D9D9D9", lineWidthPoints: 0.5)
    .SetScatterYAxisGridlines(showMajor: true, showMinor: true, lineColor: "C0C0C0", lineWidthPoints: 0.75)
    .ClearScatterXAxisGridlines()
    .ClearScatterYAxisGridlines()
    .SetScatterXAxisNumberFormat("0.0")
    .SetScatterYAxisNumberFormat("#,##0.00")
    .SetScatterXAxisDisplayUnits(C.BuiltInUnitValues.Hundreds, "Hundreds X", showLabel: true)
    .SetScatterYAxisDisplayUnits(1000d, "Thousands Y", showLabel: true)
    .SetScatterXAxisScale(minimum: 1, maximum: 12, majorUnit: 1)
    .SetScatterYAxisScale(minimum: 0, maximum: 20, majorUnit: 5)
    .SetScatterYAxisCrossing(crossesAt: 2d);
```

### Layouts and notes (fluent)
```csharp
using OfficeIMO.PowerPoint.Fluent;

ppt.AsFluent()
   .Slide(masterIndex: 0, layoutIndex: 0, s =>
   {
       s.Title("Fluent Slide");
       s.Bullets("One", "Two", "Three");
       s.Notes("Talking points for the presenter");
   });
```

### Tables (data binding)
```csharp
record SalesRow(string Product, int Q1, int Q2);

var rows = new[] {
    new SalesRow("Alpha", 12, 15),
    new SalesRow("Beta", 9, 11)
};

var columns = new[] {
    PowerPointTableColumn<SalesRow>.Create("Product", r => r.Product).WithWidthCm(4.0),
    PowerPointTableColumn<SalesRow>.Create("Q1", r => r.Q1),
    PowerPointTableColumn<SalesRow>.Create("Q2", r => r.Q2)
};

slide.AddTable(rows, columns, left: PowerPointUnits.Cm(1.5), top: PowerPointUnits.Cm(4),
    width: PowerPointUnits.Cm(20), height: PowerPointUnits.Cm(6));
```

### Tables (merged cells)
```csharp
var table = slide.AddTable(rows: 4, columns: 4, left: PowerPointUnits.Cm(1.5),  
    top: PowerPointUnits.Cm(4), width: PowerPointUnits.Cm(20), height: PowerPointUnits.Cm(6));
table.GetCell(0, 0).Text = "Merged header";
table.MergeCells(0, 0, 0, 3);
```

### Tables (formatting helpers)
```csharp
using A = DocumentFormat.OpenXml.Drawing;

var table = slide.AddTable(rows: 3, columns: 3, left: PowerPointUnits.Cm(1), top: PowerPointUnits.Cm(5),
    width: PowerPointUnits.Cm(18), height: PowerPointUnits.Cm(5));

table.SetRowHeightsEvenly();
table.SetCellPaddingCm(0.2, 0.1, 0.2, 0.1);
table.SetCellAlignment(A.TextAlignmentTypeValues.Center, A.TextAnchoringTypeValues.Center,
    startRow: 0, endRow: 0, startColumn: 0, endColumn: 2);
table.SetCellBorders(TableCellBorders.All, "E0E0E0", widthPoints: 0.5,
    dash: A.PresetLineDashValues.Dash);
table.GetCell(0, 0).SetTextAutoFit(PowerPointTextAutoFit.Normal,
    new PowerPointTextAutoFitOptions(fontScalePercent: 80, lineSpaceReductionPercent: 10));
```

### Tables (style presets)
```csharp
var themed = slide.AddTable(rows: 4, columns: 4,
    styleName: "Medium Style 2 - Accent 1",
    left: PowerPointUnits.Cm(1), top: PowerPointUnits.Cm(1),
    width: PowerPointUnits.Cm(18), height: PowerPointUnits.Cm(5),
    firstRow: true, bandedRows: true);

var headerOnly = slide.AddTable(rows: 3, columns: 3,
    preset: PowerPointTableStylePreset.HeaderOnly,
    left: PowerPointUnits.Cm(1), top: PowerPointUnits.Cm(7),
    width: PowerPointUnits.Cm(18), height: PowerPointUnits.Cm(5));
```

### Guides & grid
```csharp
ppt.SnapToGrid = true;
ppt.SetGridSpacingCm(0.5, 0.5);

ppt.ClearGuides();
ppt.AddGuideCm(PowerPointGuideOrientation.Vertical, 2.0);
ppt.AddColumnGuidesCm(columnCount: 3, marginCm: 1.0, gutterCm: 0.5,
    includeOuterEdges: true);
```

### Placeholders (layout-driven)
```csharp
using DocumentFormat.OpenXml.Presentation;

var slide = ppt.AddSlide(masterIndex: 0, layoutIndex: 1);
var title = slide.GetPlaceholder(PlaceholderValues.Title);
title?.SetTextMarginsCm(0.2, 0.1, 0.2, 0.1);
if (title != null) title.Text = "Layout Placeholder";

var layoutPlaceholders = slide.GetLayoutPlaceholders();
var titleBounds = slide.GetLayoutPlaceholderBounds(PlaceholderValues.Title);
if (titleBounds != null) {
    var aligned = slide.AddTextBox("Aligned to layout");
    titleBounds.Value.ApplyTo(aligned);
}
```

### Replace text
```csharp
ppt.ReplaceText("FY24", "FY25", includeTables: true, includeNotes: true);
```

### Sections
```csharp
ppt.AddSection("Intro", startSlideIndex: 0);
ppt.AddSection("Results", startSlideIndex: 2);
ppt.RenameSection("Results", "Deep Dive");

var sections = ppt.GetSections();
```

## Feature Highlights

- Slides: add, import, duplicate, reorder, hide, edit, and section slides
- Sections: add, rename, and list sections
- Shapes: basic rectangles/ellipses/lines with fill/stroke, line styles, shadows/glow/soft edges/blur/reflection; align/distribute; z-order + duplicate
- Images: add images from file/stream (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/ICO/PCX)   
- Properties: set built‑in and application properties
- Themes & transitions: default theme/table styles + slide transitions
- Guides & grid: snap, spacing, and guide helpers
- Text boxes: margins, auto-fit, vertical alignment
- Tables: styling + merged cells + sizing helpers
- Placeholders: read/update layout placeholders
- Backgrounds: set background images
- Text replacement: find/replace across slides
- Charts: add + format titles/legend/labels/series/markers

## Feature Matrix (scope today)

- 📽️ Slides
  - ✅ Add slides; ✅ import/duplicate/reorder; ✅ hide/show; ✅ sections; ✅ set title; ✅ add text boxes; ✅ basic bullets
- 🖼️ Media & Shapes
  - ✅ Insert images; ✅ basic shapes (rect/ellipse/line) with fill/stroke + line styles + shadows/glow/soft edges/blur/reflection; ✅ align/distribute/arrange
- 🗒️ Notes & Layout
  - ✅ Speaker notes; ✅ guides/grid helpers; ⚠️ basic layout selection
- 📋 Tables
  - ⚠️ Basic styling + merged cells
- 📊 Charts
  - ✅ Add clustered column, pie, and doughnut charts; ✅ title/legend/labels; ✅ axis formatting; ✅ series fill/line/markers
- ✨ Themes/Transitions
  - ✅ Default theme + full table styles; ✅ slide transitions (fade/wipe/push/etc.)

> Roadmap: richer shape/text APIs, layout/master controls, advanced charts — tracked in issues.

## Why OfficeIMO.PowerPoint (today)

- Cross‑platform, pure Open XML — no Office automation
- Simple API surface to add slides, titles, text, bullets, and images without repair prompts
- Fluent helpers available for quick demos and templated decks

## Measurements

Positions and sizes are stored in EMUs (English Metric Units). Use `PowerPointUnits` or the `SetPositionCm`/`SetSizeCm`
helpers to work in centimeters, inches, or points.

### Slide size presets
```csharp
ppt.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
ppt.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen4x3, portrait: true);
ppt.SlideSize.SetSizeCm(25.4, 14.0); // custom size
var ratio = ppt.SlideSize.AspectRatio;
var portrait = ppt.SlideSize.IsPortrait;
```

### Layout helpers
```csharp
var content = ppt.SlideSize.GetContentBoxCm(1.5);
var columns = ppt.SlideSize.GetColumnsCm(2, marginCm: 1.5, gutterCm: 1.0);      
columns[0].ApplyTo(slide.AddTextBox("Left column"));
columns[1].ApplyTo(slide.AddTextBox("Right column"));
```

### Layout boxes as parameters
```csharp
var content = ppt.SlideSize.GetContentBoxCm(1.5);
var columns = content.SplitColumnsCm(2, 1.0);
slide.AddTextBox("Left column", columns[0]);
slide.AddPicture("logo.png", columns[1]);
```

## Dependencies & License

- DocumentFormat.OpenXml: 3.3.x (range [3.3.0, 4.0.0))
- License: MIT

<!-- (No migration notes: these APIs are new additions.) -->
