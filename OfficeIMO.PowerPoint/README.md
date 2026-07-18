# OfficeIMO.PowerPoint - PowerPoint presentations for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

`OfficeIMO.PowerPoint` creates, reads, edits, and writes `.pptx` presentations and PowerPoint 97-2003
`.ppt`, `.pot`, and `.pps` files. It works without COM automation and without Microsoft PowerPoint installed.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint
```

## Quick start

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.Drawing;

using var presentation = PowerPointPresentation.Create("deck.pptx");
presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

var slide = presentation.AddSlide();
slide.AddTitle("OfficeIMO.PowerPoint");

var body = slide.AddTextBox("Generated without PowerPoint automation.");
body.SetPositionCm(2, 2);
body.SetSizeCm(18, 2);

slide.Transition = SlideTransition.Fade;
presentation.Save();
```

`Create(...)` starts with zero slides, and each `AddSlide()` call creates exactly one. This keeps creation and
editing deterministic; there is no hidden placeholder slide to reuse.

Load operations are detached from their source. Persistence is explicit by default, and read-only intent or
save-on-dispose behavior uses the same shared lifecycle options as Word and Excel:

```csharp
using var edited = PowerPointPresentation.Load("deck.pptx");
edited.ReplaceText("Draft", "Approved");
edited.Save();

using var inspected = PowerPointPresentation.Load("deck.pptx",
    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

using var output = new MemoryStream();
using var streamed = PowerPointPresentation.Create(output);
streamed.AddSlide().AddTitle("Stream-backed deck");
streamed.Save();
```

### PowerPoint 97-2003 binary files

Normal loading detects the OLE document streams, so callers do not need a separate code path for `.ppt`,
`.pot`, or `.pps` input. Supported binary content is projected into the same editable model used for PPTX:

```csharp
using var presentation = PowerPointPresentation.Load("legacy-deck.ppt");

Console.WriteLine(presentation.SourceFormat); // Ppt
presentation.ReplaceText("Draft", "Approved");
presentation.SaveCopy("updated.pptx");
```

Use `LoadLegacyPptWithReport(...)` when an import gate needs parser diagnostics and a conversion-loss result:

```csharp
using OfficeIMO.PowerPoint.LegacyPpt;

using LegacyPptLoadResult result =
    PowerPointPresentation.LoadLegacyPptWithReport("legacy-deck.ppt");

result.EnsureNoImportErrors();
if (result.HasConversionLoss) {
    foreach (var diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

PowerPointPresentation presentation = result.Document;
```

Native binary saving is selected by a `.ppt`, `.pot`, or `.pps` destination, or explicitly through
`ToBytes(PowerPointFileFormat.Ppt)`. Fresh binary authoring and supported PPTX conversion use the same
slides, masters, layouts, placeholders, themes, text, shapes, pictures, notes, comments, interactions,
transitions, animations, properties, macros, embedded objects, and media contracts exposed by the normal
PowerPoint model:

Explicit `.pptx` output is macro-free. VBA imported from a binary presentation remains available in the
editable model and is retained when saving back to `.ppt`, `.pot`, or `.pps`, but it is omitted from
`Save("output.pptx")` and `ToBytes(PowerPointFileFormat.Pptx)` so the package content matches its extension.
Matching macro-enabled Open XML path destinations (`.pptm`, `.potm`, `.ppsm`, and `.ppam`) preserve their
package type and VBA, including encrypted path saves. Associated no-format stream saves preserve the loaded
macro-enabled package type and VBA as well.

```csharp
using OfficeIMO.PowerPoint.LegacyPpt;

using var presentation = PowerPointPresentation.Create("simple-deck.ppt");
var slide = presentation.AddSlide();
slide.AddTitle("Binary PowerPoint");
slide.AddTextBox("Written without PowerPoint automation.");
slide.AddRectangle(600000, 3200000, 1800000, 700000);

LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();
if (report.CanWrite) {
    presentation.Save();
}
```

Imported binary files use a preservation-aware writer: no-op saves can retain the original compound file,
representable edits append compatible records, and unrelated or unknown records and streams remain intact.
Saving blocks on known loss by default. Tables, charts, and SmartArt can be converted to deterministic static
PNG visuals only after the caller accepts the reported loss with
`new PowerPointSaveOptions { LossPolicy = PowerPointConversionLossPolicy.Allow }`. Features with no safe
binary representation are blocked rather than silently omitted.

The versioned capability catalog is the source of truth for import, fresh binary authoring, binary round-trip,
and PPTX-to-binary behavior:

```csharp
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

string json = LegacyPptCapabilityCatalog.ToJson();
string markdown = LegacyPptCapabilityCatalog.ToMarkdown();
LegacyPptCapability pictures = LegacyPptCapabilityCatalog.Get(
    LegacyPptFeature.RasterPictures);
```

Password-to-open encryption works for both Open XML and binary presentations. Binary destinations use RC4
CryptoAPI for PowerPoint 97-2003 compatibility; it is a legacy interoperability mechanism, not modern
cryptography:

```csharp
using var encrypted = PowerPointPresentation.LoadEncrypted(
    "protected.ppt", "open-password");

encrypted.SaveEncrypted("protected-copy.ppt", "new-password",
    new PowerPointSaveOptions {
        LegacyPptEncryptionKeySizeBits = 128,
        LegacyPptEncryptDocumentProperties = true
    });
```

`InspectSignatures()` detects Open XML and legacy binary signature metadata. The safe default blocks a
mutating save of signed content; callers must explicitly choose whether invalidated signature markup should
be removed or preserved.

## What it does

- Creates and edits PowerPoint presentations, slides, slide size, text boxes, pictures, tables, charts, backgrounds, transitions, notes, and metadata.
- Reads and writes `.ppt`, `.pot`, and `.pps` through a dependency-free binary parser, native writer, and preservation-aware incremental writer.
- Keeps generated output as editable PowerPoint content instead of screenshots.
- Reports editable, partially editable, preserved, and unsupported deck features through `InspectFeatures()` before edit-heavy round trips.
- Composes reusable semantic plans through one `presentation.Compose(plan, options)` workflow.
- Provides semantic executive-summary, chart-story, comparison, screenshot-story, appendix-table, architecture, and closing families with two editable variants each.
- Inspects deck rhythm before rendering so repetitive layouts, dense streaks, long sections, and missing closings are visible to automation.
- Authors every `OfficeIMO.Drawing.OfficeChartKind` family from one shared chart contract, including categorical combo charts and secondary value axes.
- Supports encrypted Open XML and RC4 CryptoAPI binary presentation save/load workflows.
- Uses Open XML directly, making it suitable for services, build agents, desktop apps, and automation hosts.

## Runnable samples

```powershell
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --modern-powerpoint
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-design-brief
dotnet run --project OfficeIMO.Examples/OfficeIMO.Examples.csproj -f net10.0 -- --powerpoint-deck-plan
```

## Examples

The quick start creates one simple slide. These examples show the editable deck features that belong in `OfficeIMO.PowerPoint`.

### Title, content, and bullets

```csharp
var slide = presentation.AddSlide();
slide.AddTitle("Quarterly Review");

slide.AddTextBox("Agenda",
    PowerPointUnits.FromCentimeters(1.5), PowerPointUnits.FromCentimeters(2.0),
    PowerPointUnits.FromCentimeters(8.0), PowerPointUnits.FromCentimeters(1.0));

var agenda = slide.AddTextBox("Topics",
    PowerPointUnits.FromCentimeters(1.5), PowerPointUnits.FromCentimeters(3.0),
    PowerPointUnits.FromCentimeters(10.0), PowerPointUnits.FromCentimeters(3.0));
agenda.AddBullets(new[] { "Intro", "KPIs", "Next steps" });
```

### Images and SVGs

```csharp
using OfficeIMO.PowerPoint;
using PowerPointImagePartType = OfficeIMO.PowerPoint.ImagePartType;

slide.AddPicture("logo.png",
    PowerPointUnits.FromCentimeters(23), PowerPointUnits.FromCentimeters(1.2),
    PowerPointUnits.FromCentimeters(5), PowerPointUnits.FromCentimeters(2));

using var logo = File.OpenRead("logo.png");
slide.AddPicture(logo, PowerPointImagePartType.Png,
    PowerPointUnits.FromCentimeters(2), PowerPointUnits.FromCentimeters(2),
    PowerPointUnits.FromCentimeters(5), PowerPointUnits.FromCentimeters(2));

slide.AddPicture("diagram.svg",
    PowerPointUnits.FromCentimeters(2), PowerPointUnits.FromCentimeters(5),
    PowerPointUnits.FromCentimeters(8), PowerPointUnits.FromCentimeters(4));
```

### Shapes and layout

```csharp
slide.AddRectangle(
        PowerPointUnits.FromCentimeters(1), PowerPointUnits.FromCentimeters(1),
        PowerPointUnits.FromCentimeters(3), PowerPointUnits.FromCentimeters(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");

slide.AddLine(
        PowerPointUnits.FromCentimeters(1), PowerPointUnits.FromCentimeters(3),
        PowerPointUnits.FromCentimeters(8), PowerPointUnits.FromCentimeters(3))
    .Stroke("#404040");
```

### Tables from data

```csharp
using OfficeIMO.PowerPoint;

var rows = new[] {
    new SalesRow("Alpha", 12, 15),
    new SalesRow("Beta", 9, 11)
};

var columns = new[] {
    PowerPointTableColumn<SalesRow>.Create("Product", row => row.Product).WithWidthCm(4.0),
    PowerPointTableColumn<SalesRow>.Create("Q1", row => row.Q1),
    PowerPointTableColumn<SalesRow>.Create("Q2", row => row.Q2)
};

slide.AddTable(rows, columns,
    left: PowerPointUnits.FromCentimeters(1.5),
    top: PowerPointUnits.FromCentimeters(4),
    width: PowerPointUnits.FromCentimeters(20),
    height: PowerPointUnits.FromCentimeters(6));

record SalesRow(string Product, int Q1, int Q2);
```

### Charts from data

```csharp
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;

var metrics = new[] {
    new MetricRow("Q1", 120, 32),
    new MetricRow("Q2", 145, 36),
    new MetricRow("Q3", 172, 41),
    new MetricRow("Q4", 190, 44)
};

var chartData = new OfficeChartData(
    metrics.Select(row => row.Quarter),
    new[] {
        new OfficeChartSeries("Revenue", metrics.Select(row => row.Revenue)),
        new OfficeChartSeries("Margin", metrics.Select(row => row.Margin))
    });

var slide = presentation.AddSlide();
slide.AddTitle("Revenue and margin");

slide.AddChartCm(OfficeChartKind.ColumnClustered, chartData,
        leftCm: 1.4, topCm: 3.0, widthCm: 13.2, heightCm: 8.0)
    .SetTitle("Quarterly performance")
    .SetCategoryAxisTitle("Quarter")
    .SetValueAxisTitle("Value")
    .SetLegend(LegendPositionValues.Bottom)
    .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "D9E2F3")
    .SetPlotAreaStyle(fillColor: "F8FAFC", lineColor: "D9E2F3");

record MetricRow(string Quarter, double Revenue, double Margin);
```

### Shared chart families, combo axes, and accessibility

Use `OfficeChartData` when the same categories and series should drive PowerPoint, Excel, Drawing, HTML, PDF,
or image workflows. A series can choose its own chart kind and primary or secondary value axis without a
PowerPoint-only chart model.

```csharp
using OfficeIMO.Drawing;

var sharedData = new OfficeChartData(
    new[] { "Q1", "Q2", "Q3", "Q4" },
    new[] {
        new OfficeChartSeries("Revenue", new[] { 120d, 145d, 172d, 190d },
            xValues: null, color: OfficeColor.Parse("#0B7FAB"), pointColors: null,
            showMarkers: false, renderKind: OfficeChartKind.ColumnClustered),
        new OfficeChartSeries("Margin", new[] { 22d, 26d, 31d, 35d },
            xValues: null, color: OfficeColor.Parse("#E85D04"), pointColors: null,
            showMarkers: true, markerSize: 8, renderKind: OfficeChartKind.Line,
            axisGroup: OfficeChartAxisGroup.Secondary)
    });

PowerPointChart chart = slide.AddChartCm(
    OfficeChartKind.ColumnClustered, sharedData,
    leftCm: 1.5, topCm: 3, widthCm: 22, heightCm: 9,
    accessibility: new PowerPointChartAccessibilityOptions {
        AlternativeText = "Revenue columns with margin line on a secondary axis"
    });

chart.SaveDataSummary("quarterly-chart.txt");
```

The shared authoring overload covers clustered, stacked, and 100% stacked column/bar; line variants; area
variants; scatter; radar; pie; and doughnut. `TryGetOfficeSnapshot()` returns the same dependency-free contract
used by PNG/SVG, HTML, and PDF paths.

```csharp
var mix = new OfficeChartData(
    new[] { "Services", "Licenses", "Support" },
    new[] { new OfficeChartSeries("Share", new[] { 55d, 30d, 15d }) });

slide.AddChartCm(OfficeChartKind.Doughnut, mix,
        leftCm: 15.2, topCm: 3.0, widthCm: 8.0, heightCm: 8.0)
    .SetTitle("Revenue mix")
    .SetLegend(LegendPositionValues.Right);
```

### Table and chart together

```csharp
var segments = new[] {
    new SegmentRow("Enterprise", 18, 22, 29, 35),
    new SegmentRow("SMB", 12, 14, 18, 21),
    new SegmentRow("Public", 9, 11, 12, 16)
};

var dashboard = presentation.AddSlide();
dashboard.AddTitle("Segment dashboard");

dashboard.AddTableCm(segments, new[] {
        PowerPointTableColumn<SegmentRow>.Create("Segment", row => row.Segment).WithWidthCm(4.0),
        PowerPointTableColumn<SegmentRow>.Create("Q1", row => row.Q1),
        PowerPointTableColumn<SegmentRow>.Create("Q2", row => row.Q2),
        PowerPointTableColumn<SegmentRow>.Create("Q3", row => row.Q3),
        PowerPointTableColumn<SegmentRow>.Create("Q4", row => row.Q4)
    },
    leftCm: 1.4, topCm: 3.0, widthCm: 10.0, heightCm: 5.0);

var segmentChart = new OfficeChartData(
    segments.Select(row => row.Segment),
    new[] {
        new OfficeChartSeries("Q1", segments.Select(row => (double)row.Q1)),
        new OfficeChartSeries("Q2", segments.Select(row => (double)row.Q2)),
        new OfficeChartSeries("Q3", segments.Select(row => (double)row.Q3)),
        new OfficeChartSeries("Q4", segments.Select(row => (double)row.Q4))
    });

dashboard.AddChartCm(OfficeChartKind.Line, segmentChart,
        leftCm: 12.2, topCm: 3.0, widthCm: 12.0, heightCm: 6.5)
    .SetTitle("Segment trend")
    .SetLegend(LegendPositionValues.Bottom);

record SegmentRow(string Segment, int Q1, int Q2, int Q3, int Q4);
```

### Slides, notes, and duplication

```csharp
var duplicate = presentation.DuplicateSlide(0);
duplicate.Hidden = true;
duplicate.Notes.Text = "Backup slide for Q&A.";
```

### Feature inspection

Use `InspectFeatures()` before broad edits or automated round trips when input decks may contain package features outside the editable OfficeIMO surface:

```csharp
using var presentation = PowerPointPresentation.Load("incoming.pptx");

PowerPointFeatureReport report = presentation.InspectFeatures();
report.EnsureNoUnsupportedFeatures();

foreach (PowerPointFeatureFinding feature in report.PreservedFeatures) {
    Console.WriteLine($"{feature.Name}: {feature.Count}");
}
```

### Accessibility, review, and animation inspection

Generated and imported decks can be inspected with structured policies before they enter CI or a publishing workflow:

```csharp
using var presentation = PowerPointPresentation.Load("incoming.pptx",
    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

var options = new PowerPointInspectionOptions {
    InspectFeatures = true,
    InspectReviewComments = true,
    InspectAnimations = true,
    Accessibility = PowerPointAccessibilityOptions.ForProfile(PowerPointAccessibilityPolicyProfile.Strict)
};
PowerPointInspectionReport inspection = presentation.Inspect(options);
inspection.Accessibility!.SaveJson("incoming.accessibility.json");
inspection.Accessibility.EnsureCompliant();
Console.WriteLine($"{inspection.ReviewComments!.Comments.Count} comments, " +
                  $"{inspection.Animations!.Nodes.Count} timing nodes");
```

Shapes expose `Title`, `Description`, `Decorative`, `ReadingOrder`, `MoveToReadingOrder(...)`, and language helpers. Designer slides apply accessible defaults. The default report treats missing alternative text, table headers, slide titles, and resolvable contrast failures as errors; the strict profile also requires explicit document and shape metadata.

Saving a signed package is blocked by default because mutation invalidates existing signatures. Choose `RemoveInvalidatedSignatures` or `PreserveSignatureMarkup` explicitly only after inspecting `InspectSignatures()`.

### SmartArt and visual proof

Use the bounded semantic SmartArt workflows when native editable diagram data is more useful than flattened artwork:

```csharp
PowerPointSmartArt process = slide.AddSmartArt(
    PowerPointSmartArtType.BasicProcess,
    new[] { "Discover", "Design", "Deliver" });

PowerPointVisualProofReport proof = presentation.InspectVisuals();
proof.RecordArtifact("deck.pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    File.ReadAllBytes("deck.pptx"));
proof.SaveJson("deck.visual-proof.json");
```

The visual proof report records structural and extraction evidence, accessibility results, shared-snapshot diagnostics, PNG/SVG hashes, caller-supplied conversion artifacts, and perceptual-comparison results. PowerPoint Desktop reference rendering is available through `PowerPointDesktopReferenceRenderer.TryRender(...)` only when the caller explicitly enables it; normal generation and export never use Office automation.

### Designer composition

Use a semantic plan when a deck needs readable business composition without hand-positioning every object.
`PowerPointDeckPlan` owns intent and `PowerPointPresentation.Compose(...)` is the only public rendering operation:

```csharp
using var presentation = PowerPointPresentation.Create("proposal.pptx");

var brief = PowerPointDesignBrief
    .FromBrand("#008C95", "client-demo", "technical rollout proposal")
    .WithIdentity("Client Theme", footerLeft: "CLIENT", footerRight: "Service deck");

var plan = new PowerPointDeckPlan()
    .AddSection("Delivery plan", "Implementation overview")
    .AddProcess("Rollout", "Editable native shapes", new[] {
        new PowerPointProcessStep("Discover", "Confirm scope and dependencies."),
        new PowerPointProcessStep("Deliver", "Implement in controlled waves."),
        new PowerPointProcessStep("Operate", "Hand over evidence and ownership.")
    });

PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(brief);
composition.SelectBestAlternative = false;
composition.AlternativeIndex = 0;
PowerPointCompositionResult result = presentation.Compose(plan, composition);
presentation.Save();
```

Use a semantic plan when the whole story should stay reusable. Story slides render as native charts, tables,
pictures, shapes, and connectors rather than flattened artwork. Oversized appendix tables continue across slides,
and the rhythm report can be evaluated before the deck is created.

```csharp
using OfficeIMO.Drawing;

var chartStory = new PowerPointChartStoryContent(
    OfficeChartKind.ColumnClustered,
    new OfficeChartData(
        new[] { "Q1", "Q2", "Q3", "Q4" },
        new[] { new OfficeChartSeries("Adoption", new[] { 28d, 43d, 61d, 72d }) }),
    new[] { "Adoption improved every quarter." }) {
    Provenance = "Customer success dataset",
    AlternativeText = "Quarterly adoption columns",
    DataSummary = "Adoption rose from 28 to 72 percent."
};

var plan = new PowerPointDeckPlan()
    .AddSection("Quarterly review", "Decision-ready evidence")
    .AddChartStory("Adoption", null, chartStory)
    .AddClosing("Next action", new PowerPointClosingContent(
        "Turn the evidence into action.", "Approve the pilot"));

PowerPointCompositionResult result = presentation.Compose(plan,
    PowerPointCompositionOptions.FromBrief(brief));
PowerPointDeckRhythmReport rhythm = result.Plan.InspectRhythm(result.Design);
```

### Corporate templates and brand import

Inventory a real `.pptx` or `.potx` before generating slides. The inventory exposes masters, named layouts, semantic placeholders, theme tokens, likely logos, footer content, slide size, and layout safe areas. Missing or ambiguous semantic names fail with candidate diagnostics instead of selecting an arbitrary layout.

```csharp
PowerPointTemplateInventory inventory =
    PowerPointTemplate.Inspect("Corporate.potx");

PowerPointTemplateLayoutInfo contentLayout =
    inventory.ResolveLayout("Executive Summary");

var layoutMap = new PowerPointTemplateLayoutMap()
    .Map(PowerPointDeckPlanSlideKind.Section, inventory, "Title")
    .Map(PowerPointDeckPlanSlideKind.Capability, contentLayout);

using var presentation = PowerPointTemplate.CreatePresentation(
    "Corporate.potx",
    "Proposal.pptx",
    new PowerPointTemplateCreationOptions {
        SlideRetention = PowerPointTemplateSlideRetention.None
    });

var plan = new PowerPointDeckPlan()
    .AddSection("Service proposal", "Generated into named corporate layouts.")
    .AddCapability("Operating model", null, new[] {
        new PowerPointCapabilitySection("Governance", "Clear ownership and decisions."),
        new PowerPointCapabilitySection("Delivery", "Repeatable rollout evidence.")
    });

PowerPointDesignBrief brief = inventory.CreateDesignBrief("proposal-seed", "service proposal");
PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(brief);
composition.TemplateLayouts = layoutMap;
composition.ApplyTheme = false; // the copied template remains the native theme owner
presentation.Compose(plan, composition);
presentation.Save();
```

Use `layout.ResolvePlaceholder("Customer Screenshot")` or `ResolvePlaceholder(PowerPointTemplatePlaceholderRole.Image)` when placing a native image, chart, table, or text box into authored placeholder bounds. `CreateDesignBrief(...)` maps the imported brand into designer tokens; `ApplyBrandTo(...)` applies the same colors and fonts to native theme parts in another presentation.

Concrete `PowerPointPresentation`, `PowerPointSlide`, and `PowerPointShape` objects are the editing API. There is
no separate PowerPoint builder vocabulary to learn or keep synchronized.

See [the breaking API migration guide](../Docs/officeimo.powerpoint-api-migration.md) for old-to-new mappings.

## Managed image export

Slides and presentation batches can be exported as PNG, JPEG, TIFF, lossless WebP, or SVG:

```csharp
using OfficeIMO.Drawing;

byte[] jpeg = slide.ToJpeg(new PowerPointImageExportOptions {
    RasterEncoding = new OfficeRasterEncodingOptions {
        Jpeg = new OfficeJpegEncodeOptions { Quality = 92 }
    }
});

presentation.ToImages()
    .ForSlideRange(1, 3)
    .AsWebp()
    .Save("slide-previews");
```

PowerPoint owns slide semantics and scene construction; `OfficeIMO.Drawing` owns the common raster encoders. `SaveAsJpeg`, `SaveAsTiff`, and `SaveAsWebp` remain thin wrappers over the shared export builder.

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Export editable PowerPoint slides to PDF and import PDF tables to PowerPoint. |
| [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) | Render semantic markup documents to PowerPoint. |

## Boundaries

- `OfficeIMO.PowerPoint` owns editable PowerPoint creation and manipulation.
- PDF export belongs in `OfficeIMO.PowerPoint.Pdf` and shared PDF primitives belong in `OfficeIMO.Pdf`.
- Showcase decks and long design examples belong in `OfficeIMO.Examples` or focused docs, not in the package README.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** Open XML SDK for `.pptx` package mechanics. Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Drawing`. The presentation model, composition system, inspection, encryption workflow, and PNG/JPEG/TIFF/WebP/SVG export are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
