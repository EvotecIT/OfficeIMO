# OfficeIMO.PowerPoint API migration

This guide covers the intentional breaking cleanup that makes presentation creation, editing, semantic
composition, templates, inspection, and exports follow one ownership model.

## Ownership model

| Capability | Public owner |
| --- | --- |
| Create, load, save, slides, sections, masters, and themes | `PowerPointPresentation` |
| Concrete content and editing | `PowerPointSlide` and concrete `PowerPointShape` types |
| Semantic story intent | `PowerPointDeckPlan` |
| Semantic rendering | `PowerPointPresentation.Compose(...)` |
| Corporate template inspection and copying | `PowerPointTemplate` |
| Cross-format chart data | `OfficeIMO.Drawing.OfficeChartData` |
| Package and quality inspection | `PowerPointPresentation.Inspect(...)` and focused `Inspect...` methods |
| PDF and HTML conversion | `OfficeIMO.PowerPoint.Pdf` and `OfficeIMO.PowerPoint.Html` extensions |

The old designer extensions, public deck composer, PowerPoint-only fluent builders, and PowerPoint-only chart
data contract were removed from the public surface. Their layout and Open XML behavior remains internal to the
single owning workflow.

Custom semantic slides remain available through `PowerPointDeckPlan.AddCustom(...)`. Its callback receives a
`PowerPointSlideCompositionContext`, making the type's role as plan-owned context explicit instead of exposing
another standalone composer entry point.

## Lifecycle

Creation and loading are detached from their destinations. `Save()` writes to the associated path or stream,
while disposal only persists when `SaveOnDispose` is explicitly requested:

```csharp
using PowerPointPresentation presentation = PowerPointPresentation.Load("deck.pptx");
presentation.ReplaceText("Draft", "Approved");
presentation.Save();
```

Read-only intent uses the shared lifecycle contract:

```csharp
using PowerPointPresentation presentation = PowerPointPresentation.Load(
    "deck.pptx", new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
```

Stream behavior uses the same typed options. The caller retains stream ownership, and its position is preserved
when loading:

```csharp
using var stream = new MemoryStream();
using (PowerPointPresentation created = PowerPointPresentation.Create(stream)) {
    created.AddSlide().AddTitle("Stream-backed deck");
    created.Save();
}

using PowerPointPresentation inspected = PowerPointPresentation.Load(stream,
    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
```

| Before | After |
| --- | --- |
| `Open(path)` | `Load(path)` |
| `OpenRead(path)` | `Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Open(stream, readOnly: true, autoSave: false)` | `Load(stream, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })` |
| `Create(stream, autoSave: false)` | `Create(stream)` |
| implicit save on dispose | `PersistenceMode = DocumentPersistenceMode.SaveOnDispose` |

## Concrete editing

Use the presentation, slide, and shape objects directly. The removed `OfficeIMO.PowerPoint.Fluent` wrapper
duplicated the same operations with different names and did not own additional document behavior.

New presentations start with zero slides. Every `AddSlide()` call creates exactly one slide, so setup and
editing code no longer depend on a hidden reusable placeholder slide.

```csharp
using PowerPointPresentation presentation = PowerPointPresentation.Create("deck.pptx");
PowerPointSlide slide = presentation.AddSlide();
PowerPointTextBox title = slide.AddTitle("Quarterly review");
title.FontSize = 30;

PowerPointSlide detail = presentation.AddSlide();
detail.AddTextBoxCm("Editable content", 1.5, 2.0, 20.0, 1.5);
presentation.MoveSlide(1, 0);
presentation.Save();
```

## Semantic composition

Previously, the same semantic slide could be added through `AddDesigner...`, `PowerPointDeckComposer.Add...`,
or `PowerPointDeckPlan.Add...`. Build the plan once and compose it once:

```csharp
PowerPointDeckPlan plan = new PowerPointDeckPlan()
    .AddSection("Delivery plan", "Implementation overview")
    .AddProcess("Rollout", null, new[] {
        new PowerPointProcessStep("Discover", "Confirm requirements."),
        new PowerPointProcessStep("Deliver", "Implement in controlled waves."),
        new PowerPointProcessStep("Operate", "Transfer ownership and evidence.")
    });

PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand(
    "#008C95", "customer-rollout", "technical rollout proposal");

using PowerPointPresentation presentation = PowerPointPresentation.Create("proposal.pptx");
PowerPointCompositionResult result = presentation.Compose(
    plan, PowerPointCompositionOptions.FromBrief(brief));
presentation.Save();
```

`PowerPointCompositionResult` returns the actual continuation-expanded plan, resolved design, rendered editable
slides, and preflight report. Use `PreviewComposition(...)` when the resolved variants are needed before mutation.

| Before | After |
| --- | --- |
| `presentation.UseDesigner(...).AddSlides(plan)` | `presentation.Compose(plan, PowerPointCompositionOptions.FromBrief(brief))` |
| `presentation.AddDesignerProcessSlide(...)` | `plan.AddProcess(...)`, then `presentation.Compose(...)` |
| `deck.AddSlidesWithContinuation(plan)` | `options.ExpandContinuations = true` (the default), then `Compose(...)` |
| `deck.AddSlidesWithReport(plan)` | `PowerPointCompositionResult result = presentation.Compose(...)` |

## Templates

Templates are optional adapters, not the presentation model:

```csharp
PowerPointTemplateInventory inventory = PowerPointTemplate.Inspect("Corporate.potx");
PowerPointTemplateLayoutMap layouts = new PowerPointTemplateLayoutMap()
    .Map(PowerPointDeckPlanSlideKind.Section, inventory, "Title")
    .Map(PowerPointDeckPlanSlideKind.Capability, inventory, "Executive Summary");

using PowerPointPresentation presentation = PowerPointTemplate.CreatePresentation(
    "Corporate.potx", "Proposal.pptx",
    new PowerPointTemplateCreationOptions { SlideRetention = PowerPointTemplateSlideRetention.None });

PowerPointCompositionOptions options = PowerPointCompositionOptions.FromBrief(
    inventory.CreateDesignBrief("proposal", "service proposal"));
options.TemplateLayouts = layouts;
options.ApplyTheme = false;
presentation.Compose(plan, options);
presentation.Save();
```

| Before | After |
| --- | --- |
| `PowerPointPresentation.InspectTemplate(path)` | `PowerPointTemplate.Inspect(path)` |
| `PowerPointPresentation.CreateFromTemplate(...)` | `PowerPointTemplate.CreatePresentation(...)` |
| `presentation.UseTemplateDesigner(...)` | `PowerPointCompositionOptions.TemplateLayouts`, then `Compose(...)` |

## Charts

`OfficeChartData` is now the public chart contract. It is shared by PowerPoint, Excel, Drawing, HTML, PDF, and
image workflows and supports combo kinds and secondary axes without a PowerPoint-only copy.

```csharp
var data = new OfficeChartData(
    new[] { "Q1", "Q2", "Q3" },
    new[] { new OfficeChartSeries("Revenue", new[] { 10d, 14d, 19d }) });

PowerPointChart chart = slide.AddChartCm(
    OfficeChartKind.ColumnClustered, data, 1.5, 3.0, 20.0, 8.0);
chart.UpdateData(data);
```

Replace `PowerPointChartData`, `PowerPointScatterChartData`, their series types, and family-specific add methods
with `OfficeChartData` plus `AddChart`, `AddChartCm`, `AddChartInches`, or `AddChartPoints`.

## Inspection and save

Use `Inspect()` for an end-to-end gate over package validity, visual preflight, and accessibility. Opt into feature,
review, animation, signature, or rendered visual proof when the workflow needs them:

```csharp
PowerPointInspectionReport report = presentation.Inspect(new PowerPointInspectionOptions {
    InspectFeatures = true,
    InspectReviewComments = true,
    InspectAnimations = true
});

if (!report.IsSuccessful) {
    throw new InvalidOperationException("Presentation inspection failed.");
}
presentation.Save();
```

Focused inspection names now use the same verb: `InspectPreflight()`, `InspectAccessibility()`,
`InspectFeatures()`, `InspectReviewComments()`, `InspectAnimations()`, `InspectSignatures()`, and
`InspectVisuals()`.

| Before | After |
| --- | --- |
| `Preflight()` | `InspectPreflight()` |
| `CreateVisualProofReport()` | `InspectVisuals()` |
| `SaveWithPreflight()` | `InspectPreflight()`, apply the desired gate, then `Save()` |

`ValidateDocument()` remains the focused Open XML validation operation shared with OfficeIMO Word and Excel.

## Exports

PDF and HTML remain thin extension packages over the same presentation model. Simple save methods write output;
`TrySaveAsPdf(...)` and `ToPdfDocumentResult(...)` return diagnostics where callers need machine-readable results.
No designer, template, or alternate presentation model is introduced by export.
