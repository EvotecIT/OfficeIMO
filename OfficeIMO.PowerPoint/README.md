# OfficeIMO.PowerPoint - PowerPoint presentations for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.PowerPoint)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.PowerPoint?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.PowerPoint)

`OfficeIMO.PowerPoint` creates, reads, edits, and writes `.pptx` presentations and PowerPoint 97-2003
`.ppt`, `.pot`, and `.pps` files. It works without COM automation and without Microsoft PowerPoint installed.

## Inspection without export payloads

`PowerPointExtractOptions.IncludeMarkdown = false` makes `ExtractMarkdownChunks`
emit plain slide/table/note text without constructing Markdown. Extraction options
are snapshotted when the enumerable is created. `PowerPointLoadOptions.AllowLegacyProjection = false`
rejects binary PPT before secondary payload/model projection, including encrypted
legacy input; ordinary loads keep the default import behavior.

`InspectEmbeddedPayloads(options, cancellationToken)` inventories inert package,
OLE and ActiveX descriptors and provides selected bounded `OpenRead(ordinal)`
streams. `OpenEmbeddedPayload(id, maximumBytes, cancellationToken)` opens one
payload directly. Keep the presentation open while requesting reads; returned
complete read-only streams are caller-owned and survive presentation disposal.
The [Core payload contract](../OfficeIMO.Core/README.md#bounded-embedded-payload-reads)
defines count, member, aggregate and cumulative work limits, unknown lengths and
optional hashing. `InspectVbaProjectMetadata()` inventories the part and signature
relationship without reading binary module content, with
`ModuleInspectionRequested = false`.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.PowerPoint
```

## Quick start

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Create("deck.pptx");
presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

var slide = presentation.AddSlide();
slide.AddTitle("OfficeIMO.PowerPoint");

var body = slide.AddTextBox("Generated without PowerPoint automation.");
body.SetPosition(OfficeLength.Centimeters(2), OfficeLength.Centimeters(2));
body.SetSize(OfficeLength.Centimeters(18), OfficeLength.Centimeters(2));

slide.Transition = PowerPointSlideTransition.Fade;
presentation.Save();
```

`Create(...)` starts with zero slides, and each `AddSlide()` call creates exactly one. This keeps creation and
editing deterministic; there is no hidden placeholder slide to reuse.

### Creation themes and automatic content frames

New presentations use `OfficeTheme.Default`. Choose a theme during creation:

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

var metrics = new[] { new MetricRow("Q1", 120D), new MetricRow("Q2", 145D) };
using var presentation = PowerPointDocument.Create(new PowerPointCreateOptions {
    Theme = OfficeTheme.Classic
});
var slide = presentation.AddSlide();
slide.AddTitle("Quarterly revenue");
var chart = slide.AddChart(OfficeChartType.Column, metrics, nameof(MetricRow.Quarter), nameof(MetricRow.Revenue));
var table = slide.AddTable(metrics, new OfficeTableOptions());
chart.SetTitle("Revenue");
table.GetCell(1, 0).Text = "First quarter";

record MetricRow(string Quarter, double Revenue);
```

The title and omitted chart/table bounds use the theme's measured layout grid.
One body object uses the content frame below the header; two use the engine's
content columns. Adding the title after the body objects updates their frames.
Returned chart and table objects remain attached and editable.

Automatic placement supports one title and two body objects per slide. Use another
slide or explicit geometry for additional content. Supplying any position or size
keeps that object out of automatic reflow; omitted components still use the content
frame. Changing a returned object's bounds also ends its automatic reflow.

A template retains its native theme and layouts when `Theme` is omitted. Supplying
`Theme` explicitly applies that theme to the copied presentation. Loading a file
preserves its style; imported content uses neutral scaffolding. Explicit image
frames keep their supplied bounds.

For a deck from an untrusted source, use the bounded load profile:

```csharp
using var incoming = PowerPointDocument.Load("upload.pptx", PowerPointLoadOptions.UntrustedDefaults);
```

It rejects macros, embedded payloads, ActiveX, and external relationships before parsing. Ordinary load options retain compatibility with decks containing those parts; set `PackageSecurity` explicitly for another policy.

Retain a structural report of the exact input used by the loader:

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Load("review.pptx", new PowerPointLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64L * 1024 * 1024 },
    CaptureSourcePackageSecurity = true
});
Console.WriteLine(presentation.SourcePackageSecurityReport?.ContainerKind);
Console.WriteLine(presentation.SourcePackageSecurityScope);
var loadedSafety = presentation.InspectLoadedContentSafety();
```

The captured package report stays attached to the original source scope when the
model is later edited. Loaded-content safety inspection examines that current model
without reopening, saving or running presentation content. Capture is opt-in;
native import and package limits remain necessary. A macro or signature part is
an inert descriptor, not execution or cryptographic-validation evidence.

Load operations are detached from their source. Persistence is explicit by default, and read-only intent or
save-on-dispose behavior uses the same shared lifecycle options as Word and Excel:

```csharp
using var edited = PowerPointDocument.Load("deck.pptx");
edited.ReplaceText("Draft", "Approved");
edited.Save();

using var inspected = PowerPointDocument.Load("deck.pptx",
    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });

using var output = new MemoryStream();
using var streamed = PowerPointDocument.Create(output);
streamed.AddSlide().AddTitle("Stream-backed deck");
streamed.Save();
```

### PowerPoint 97-2003 binary files

Normal loading detects the OLE document streams, so callers do not need a separate code path for `.ppt`,
`.pot`, or `.pps` input. Supported binary content is projected into the same editable model used for PPTX:

```csharp
using var presentation = PowerPointDocument.Load("legacy-deck.ppt");

Console.WriteLine(presentation.SourceFormat); // Ppt
presentation.ReplaceText("Draft", "Approved");
presentation.SaveCopy("updated.pptx");
```

Use the `Load` overload with a report when an import gate needs parser diagnostics and a conversion-loss result:

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("legacy-deck.ppt", null,
    out OfficeIMO.OfficeConversionReport report);
foreach (OfficeIMO.OfficeDiagnostic diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Native binary saving is selected by a `.ppt`, `.pot`, or `.pps` destination, or explicitly through
`ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt })`. Fresh binary authoring and supported PPTX conversion use the same
slides, masters, layouts, placeholders, themes, text, shapes, pictures, notes, comments, interactions,
transitions, animations, properties, macros, embedded objects, and media contracts exposed by the normal
PowerPoint model:

Explicit `.pptx` output is macro-free. VBA imported from a binary presentation remains available in the
editable model and is retained when saving back to `.ppt`, `.pot`, or `.pps`, but it is omitted from
`Save("output.pptx")` and `ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Pptx })` so the package content matches its extension.
Matching macro-enabled Open XML path destinations (`.pptm`, `.potm`, `.ppsm`, and `.ppam`) preserve their
package type and VBA, including encrypted path saves. Associated no-format stream saves preserve the loaded
macro-enabled package type and VBA as well.

```csharp
using OfficeIMO.PowerPoint.LegacyPpt;

using var presentation = PowerPointDocument.Create("simple-deck.ppt");
var slide = presentation.AddSlide();
slide.AddTitle("Binary PowerPoint");
slide.AddTextBox("Written without PowerPoint automation.");
slide.AddRectangle(OfficeIMO.OfficeLength.Emus(600000), OfficeIMO.OfficeLength.Emus(3200000), OfficeIMO.OfficeLength.Emus(1800000), OfficeIMO.OfficeLength.Emus(700000));

LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();
if (report.CanWrite) {
    presentation.Save();
}
```

Imported binary files use a preservation-aware writer: no-op saves can retain the original compound file,
representable edits append compatible records, and unrelated or unknown records and streams remain intact.
Saving blocks on known loss by default. Tables, charts, and SmartArt can be converted to deterministic static
PNG visuals only after the caller accepts the reported loss with
`new PowerPointSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }`. Features with no safe
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
using var encrypted = PowerPointDocument.Load(
    "protected.ppt", new PowerPointLoadOptions { Password = "open-password" });

encrypted.Save("protected-copy.ppt",
    new PowerPointSaveOptions {
        Password = "new-password",
        LegacyPptEncryptionKeySizeBits = 128,
        LegacyPptEncryptDocumentProperties = true
    });
```

`InspectSignatures()` detects Open XML and legacy binary signature metadata. The safe default blocks a
mutating save of signed content; callers must explicitly choose whether invalidated signature markup should
be removed or preserved. `SignPackageSignature(...)` and `ValidatePackageSignatures(...)` use an explicitly
supplied `IOfficeSecurityProvider` for OPC creation and cryptographic validation without adding a transitive
`OfficeIMO.Security` dependency.

## What it does

- Creates and edits PowerPoint presentations, slides, slide size, text boxes, pictures, tables, charts, backgrounds, transitions, notes, and metadata.
- Reads and writes `.ppt`, `.pot`, and `.pps` through a dependency-free binary parser, native writer, and preservation-aware incremental writer.
- Keeps generated output as editable PowerPoint content instead of screenshots.
- Reports editable, partially editable, preserved, and unsupported deck features through `InspectFeatures()` before edit-heavy round trips.
- Composes reusable semantic plans through one `presentation.Compose(plan, options)` workflow.
- Provides semantic executive-summary, chart-story, comparison, screenshot-story, appendix-table, architecture, and closing families with two editable variants each.
- Inspects deck rhythm before rendering so repetitive layouts, dense streaks, long sections, and missing closings are visible to automation.
- Authors every `OfficeIMO.OfficeChartType` family from one shared chart contract, including categorical combo charts and secondary value axes.
- Authors bounded semantic SmartArt, shared custom geometry, classic and modern threaded comments, custom shows, embedded OLE objects, and opaque VBA projects through the normal presentation model.
- Preserves and reports broader animation timelines while exposing first-party shape, text, and chart animation workflows for supported effects.
- Supports encrypted Open XML and RC4 CryptoAPI binary presentation save/load workflows.
- Uses Open XML directly, making it suitable for services, build agents, desktop apps, and automation hosts.

### Imported charts, SmartArt, timelines, media, and OLE

Imported 3-D bar, line, area, and pie charts, of-pie charts, stock charts, and
surface charts support in-place data updates when their referenced caches and
embedded workbook can be changed consistently. `InspectImportedContent()`
reports the native family and export boundary. Image and PDF export draw these
families as an explicit flat 2-D projection and report what it leaves out
(`chart-projected` in both PDF and images);
producer-specific storage or chart groups that do not have a safe model remain
preservation-only. Series and slices that store no colour or marker take the
colours and markers PowerPoint gives them, from the chart's classic style
number and the theme.

Image and PDF export draw SmartArt from the shapes PowerPoint caches with the
diagram, so producer layouts keep their geometry, colours and text. When that
drawing is missing or no longer matches the diagram text, export falls back to
the supported semantic layouts.

`PowerPointSmartArt.TryGetTopology(...)` exposes text, sibling order, and parent
connections only when the imported layout can be represented safely.
`UpdateTopology(...)` rejects node creation/deletion, cycles, and connections
that would change the meaning of the supported SmartArt family.

Slides can append motion, rotation, and command timeline actions without
rebuilding an imported timing tree. Actions can be edited or removed by native
timing ID. Linked and embedded audio/video expose playback, trim, fade, target,
and payload controls; linked and embedded OLE objects expose their appropriate
URI or package-payload controls. Unmodeled timing and package content remains in
the source package unless a caller explicitly targets a supported node.

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
    OfficeIMO.OfficeLength.Centimeters(1.5), OfficeIMO.OfficeLength.Centimeters(2.0),
    OfficeIMO.OfficeLength.Centimeters(8.0), OfficeIMO.OfficeLength.Centimeters(1.0));

var agenda = slide.AddTextBox("Topics",
    OfficeIMO.OfficeLength.Centimeters(1.5), OfficeIMO.OfficeLength.Centimeters(3.0),
    OfficeIMO.OfficeLength.Centimeters(10.0), OfficeIMO.OfficeLength.Centimeters(3.0));
agenda.AddBullets(new[] { "Intro", "KPIs", "Next steps" });
```

### Text geometry

Font sizes, paragraph spacing, indents and text padding use `OfficeLength`. Units are explicit at the call site; nullable properties remove their local override when set to `null`. Line-spacing multipliers and bullet-size percentages remain dimensionless.

```csharp
var text = slide.AddTextBox("Quarterly review");
text.FontSize = OfficeIMO.OfficeLength.Points(18.5);
text.SetTextMargins(OfficeIMO.OfficeLength.Centimeters(0.3), OfficeIMO.OfficeLength.Points(4),
    OfficeIMO.OfficeLength.Centimeters(0.3), OfficeIMO.OfficeLength.Points(4));
text.ApplyParagraphStyle(new PowerPointParagraphStyle(
    hanging: OfficeIMO.OfficeLength.Points(18),
    spaceAfter: OfficeIMO.OfficeLength.Points(6), lineSpacingMultiplier: 1.15));
text.Paragraphs[0].Runs[0].CharacterSpacing = OfficeIMO.OfficeLength.Points(0.25);
```

DrawingML font sizes retain hundredth-point precision and accept 1–4000 points. Absolute paragraph spacing rounds to hundredths of a point using the paragraph conversion rules; character spacing and font sizes use their existing midpoint-away-from-zero conversion. Insets and indentation retain native EMUs.

### Custom paragraph tab stops

`SetTabStops` replaces the paragraph's explicit stops and `TabStops` reads them back. Positions use `OfficeLength`; alignment can be left, center, right, or decimal. Passing an empty sequence removes local stops while leaving inherited defaults unchanged.

```csharp
var paragraph = slide.AddTextBox("Label\t12.50").Paragraphs[0];
paragraph.SetTabStops(new[] {
    new PowerPointTabStop(global::OfficeIMO.OfficeLength.Points(72), PowerPointTabAlignment.Decimal)
});
```

### Images and SVGs

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

slide.AddImage("logo.png",
    OfficeIMO.OfficeLength.Centimeters(23), OfficeIMO.OfficeLength.Centimeters(1.2),
    OfficeIMO.OfficeLength.Centimeters(5), OfficeIMO.OfficeLength.Centimeters(2));

using var logo = File.OpenRead("logo.png");
slide.AddImage(logo, OfficeImageFormat.Png,
    OfficeIMO.OfficeLength.Centimeters(2), OfficeIMO.OfficeLength.Centimeters(2),
    OfficeIMO.OfficeLength.Centimeters(5), OfficeIMO.OfficeLength.Centimeters(2));

slide.AddImage("diagram.svg",
    OfficeIMO.OfficeLength.Centimeters(2), OfficeIMO.OfficeLength.Centimeters(5),
    OfficeIMO.OfficeLength.Centimeters(8), OfficeIMO.OfficeLength.Centimeters(4));
```

### Shapes and layout

Shape, SmartArt, audio, video, and OLE placement uses `OfficeLength`. Mix units in
one call as needed. For optional bounds, omitting a value or passing `null` keeps
that method's default position or size; an explicit zero remains zero.

```csharp
slide.AddRectangle(
        OfficeIMO.OfficeLength.Centimeters(1), OfficeIMO.OfficeLength.Centimeters(1),
        OfficeIMO.OfficeLength.Centimeters(3), OfficeIMO.OfficeLength.Centimeters(1))
    .Fill("#E7F7FF")
    .Stroke("#007ACC");

slide.AddLine(
        OfficeIMO.OfficeLength.Centimeters(1), OfficeIMO.OfficeLength.Centimeters(3),
        OfficeIMO.OfficeLength.Centimeters(8), OfficeIMO.OfficeLength.Centimeters(3))
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
    PowerPointTableColumn<SalesRow>.Create("Product", row => row.Product).WithWidth(OfficeIMO.OfficeLength.Centimeters(4.0)),
    PowerPointTableColumn<SalesRow>.Create("Q1", row => row.Q1),
    PowerPointTableColumn<SalesRow>.Create("Q2", row => row.Q2)
};

slide.AddTable(rows, columns,
    left: OfficeIMO.OfficeLength.Centimeters(1.5),
    top: OfficeIMO.OfficeLength.Centimeters(4),
    width: OfficeIMO.OfficeLength.Centimeters(20),
    height: OfficeIMO.OfficeLength.Centimeters(6));

record SalesRow(string Product, int Q1, int Q2);
```

Table rows and columns use `OfficeLength`: `table.SetColumnWidths(...)` and `table.SetRowHeights(...)` set a batch, while `table.ColumnItems[index].Width` and `table.RowItems[index].Height` edit attached handles. `GetColumnWidth` and `GetRowHeight` return lengths; convert them with `ToPoints()` or `ToCentimeters()` when a numeric value is needed.

Use `slide.BackgroundColor` for a solid slide background or `slide.SetBackgroundNoFill()` for an explicit no-fill override that suppresses inherited layout or master fills. No-fill survives save/reopen and is returned as `PowerPointSlideBackgroundKind.None` by `GetBackground()`.

Set `table.GetCell(row, column).FillColor` to an RGB hex color to replace the cell's local fill. Set `NoFill = true` to show the slide background through that cell and suppress table-theme shading. Setting `NoFill = false` removes that override; setting `FillColor = null` removes only a local solid fill. Cell indexes are zero-based. Explicit no-fill survives PPTX save/reopen, image export and native PPT conversion.

Use `cell.SetBorders(PowerPointTableCellBorders.Bottom, "FF0000", OfficeLength.Points(3))` to set selected edges while preserving the others. Widths use `OfficeLength` and must fit DrawingML's 0–1584-point range. `cell.HideBorders(...)` explicitly hides selected edges and overrides table styling; `cell.ClearBorders(...)` removes local edge declarations so they can inherit the table style. Setting a solid border after hiding it replaces the no-fill declaration.

### Charts from data

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO;
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
        new OfficeSeries("Revenue", metrics.Select(row => row.Revenue)),
        new OfficeSeries("Margin", metrics.Select(row => row.Margin))
    });

var slide = presentation.AddSlide();
slide.AddTitle("Revenue and margin");

slide.AddChart(OfficeChartType.ColumnClustered, chartData, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = OfficeIMO.OfficeLength.Centimeters(1.4), Top = OfficeIMO.OfficeLength.Centimeters(3.0), Width = OfficeIMO.OfficeLength.Centimeters(13.2), Height = OfficeIMO.OfficeLength.Centimeters(8.0) })
    .SetTitle("Quarterly performance")
    .SetCategoryAxisTitle("Quarter")
    .SetValueAxisTitle("Value")
    .SetLegend(OfficeChartLegendPosition.Bottom)
    .SetChartAreaStyle(fillColor: "FFFFFF", lineColor: "D9E2F3")
    .SetPlotAreaStyle(fillColor: "F8FAFC", lineColor: "D9E2F3");

record MetricRow(string Quarter, double Revenue, double Margin);
```

### Shared chart families, combo axes, and accessibility

Use `OfficeChartData` when the same categories and series should drive PowerPoint, Excel, Drawing, HTML, PDF,
or image workflows. A series can choose its own chart kind and primary or secondary value axis without a
PowerPoint-only chart model.

`OfficeSeries.PointColors` assigns solid colours by point index. Native chart snapshots preserve
these overrides, including sparse and theme-based fills, so saved and reopened pie, doughnut, and column
charts retain their colours in managed image and PDF exports. A null entry inherits the chart's normal
colour selection. The [image export capability matrix](../Docs/officeimo.image-export-capability-matrix.md#powerpoint)
tracks the exact chart appearance boundary.
PDF export uses the native chart style unless `PowerPointPdfExportOptions.ChartStyle` supplies an override.

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO;
using OfficeIMO.Drawing;

var sharedData = new OfficeChartData(
    new[] { "Q1", "Q2", "Q3", "Q4" },
    new[] {
        new OfficeSeries("Revenue", new[] { 120d, 145d, 172d, 190d },
            xValues: null, color: OfficeColor.Parse("#0B7FAB"), pointColors: null,
            showMarkers: false, renderKind: OfficeChartType.ColumnClustered),
        new OfficeSeries("Margin", new[] { 22d, 26d, 31d, 35d },
            xValues: null, color: OfficeColor.Parse("#E85D04"), pointColors: null,
            showMarkers: true, markerSize: 8, renderKind: OfficeChartType.Line,
            axisGroup: OfficeChartAxisGroup.Secondary)
    });

PowerPointChart chart = slide.AddChart(OfficeChartType.ColumnClustered, sharedData, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = OfficeIMO.OfficeLength.Centimeters(1.5), Top = OfficeIMO.OfficeLength.Centimeters(3), Width = OfficeIMO.OfficeLength.Centimeters(22), Height = OfficeIMO.OfficeLength.Centimeters(9), Accessibility = new PowerPointChartAccessibilityOptions {
        AlternativeText = "Revenue columns with margin line on a secondary axis"
    } });

chart.SaveDataSummary("quarterly-chart.txt");
```

Use `chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 40,
majorUnit: 10, numberFormat: "0").WithTitle("Margin"))` to give the secondary value axis an independent linear
scale. Native data updates preserve these settings, and snapshots carry its bounds, tick units,
tick appearance, number format, and title separately from the primary axis. Use the optional
`majorTickMark` and `minorTickMark` settings for independent inside, outside, or crossing ticks.
Static snapshots reject secondary logarithmic or reversed scales, display units, and explicit
numeric crossing values. Native data updates preserve these settings.

The shared authoring overload covers clustered, stacked, and 100% stacked column/bar; line variants; area
variants; scatter; radar; pie; and doughnut. `TryGetOfficeSnapshot()` returns the same dependency-free contract
used by PNG/SVG, HTML, and PDF paths.

Category and scatter readback retains supported marker shape, size and outline,
connecting-line visibility, and stroke width and dash. Group marker settings and
scatter styles apply when a series does not override them. Static projections
reject unqualified dashes, curved lines, distinct marker and line colours, and
unfilled marker treatments. Native data editing remains available. Cached chart
projections are bounded to 100,000 positions across all native layers.
Point overrides retain their separate 100,000-record limit, including stale and
duplicate records. Native series retain their plotting order, including combination
layers. Gradients, custom or compound outlines, and per-point marker overrides reject
static projection. Supported filled-series outlines are inherited by points, with
explicit point outlines taking precedence. Picture markers reject static projection. Radar series inherit
visible markers only from the native marker style unless the series supplies a marker.

Category discovery checks all populated series before generating fallback labels.
The longest available category cache supplies labels, and shorter series retain
their positions with zero padding. Mixed snapshots require matching category
caches and share one budget for native and normalized positions across the plot.
Marker-only plots use the marker fill; a marker fill cannot replace an unresolved
connecting-line colour in a static export.

Data updates retain cached formula-linked chart and axis titles and custom
labels as native rich text, and custom error bars as numeric literals when
replacing the embedded worksheet. Uncached bindings and unqualified workbook-linked
extensions reject the update before changing chart data or workbook bytes.
Externally linked or unresolved workbook references reject shared and legacy
updates. Compatible repeated layers retain separate native formatting when
they share an axis pair; repeated layers of the same family and axis group
with different axis pairs reject shared replacement.
Legacy category-only updates require a single native chart layer; use shared
`OfficeChartData` for supported combinations. For shared scatter and bubble
authoring, each series' `RenderKind` must match the chart kind or be omitted.

```csharp
var mix = new OfficeChartData(
    new[] { "Services", "Licenses", "Support" },
    new[] { new OfficeSeries("Share", new[] { 55d, 30d, 15d }) });

slide.AddChart(OfficeChartType.Doughnut, mix, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = OfficeIMO.OfficeLength.Centimeters(15.2), Top = OfficeIMO.OfficeLength.Centimeters(3.0), Width = OfficeIMO.OfficeLength.Centimeters(8.0), Height = OfficeIMO.OfficeLength.Centimeters(8.0) })
    .SetTitle("Revenue mix")
    .SetLegend(OfficeChartLegendPosition.Right);
```

### Individual chart point styles

Attach `OfficeSeries.WithPointStyles` when authoring shared chart data, or call
`chart.SetDataPointStyle(seriesIndex, pointIndex, style)` to edit an existing native chart.
Styles support solid fill, explicit no-fill, outlines, and seven hatch patterns from
`OfficeIMO.Drawing.OfficeChartHatchPattern`. A null style clears the point's fill and outline
overrides. Values remain unchanged and the chart remains editable in PowerPoint.

Supported point appearances survive save/reopen, shared data updates, snapshots, image and
PDF export, and PowerPoint semantic HTML round trips. Static pie and doughnut legends use
the same colours and styles as the slices. Per-point area styling is retained in native
files but static export reports an approximation and uses the series fill.
See the [shared point-style example](../OfficeIMO.Core/README.md#style-individual-chart-points)
for an outlined unknown-status slice and a hatched slice.

### Pie and doughnut geometry

Use `chart.SetRadialLayout(new OfficeChartRadialLayout(90, 70))` on an existing native
two-dimensional pie or doughnut chart. The first argument rotates the first slice boundary
clockwise from the top (0–360 degrees); the second sets the doughnut hole percentage (10–90).
`chart.RadialLayout` reads the native settings. Save/reopen, data updates, snapshots,
managed image/PDF exports, and semantic HTML round trips preserve them.

### Table and chart together

```csharp
var segments = new[] {
    new SegmentRow("Enterprise", 18, 22, 29, 35),
    new SegmentRow("SMB", 12, 14, 18, 21),
    new SegmentRow("Public", 9, 11, 12, 16)
};

var dashboard = presentation.AddSlide();
dashboard.AddTitle("Segment dashboard");

dashboard.AddTable(segments, new[] {
        PowerPointTableColumn<SegmentRow>.Create("Segment", row => row.Segment).WithWidth(OfficeIMO.OfficeLength.Centimeters(4.0)),
        PowerPointTableColumn<SegmentRow>.Create("Q1", row => row.Q1),
        PowerPointTableColumn<SegmentRow>.Create("Q2", row => row.Q2),
        PowerPointTableColumn<SegmentRow>.Create("Q3", row => row.Q3),
        PowerPointTableColumn<SegmentRow>.Create("Q4", row => row.Q4)
    },
    left: OfficeLength.Centimeters(1.4), top: OfficeLength.Centimeters(3.0),
    width: OfficeLength.Centimeters(10.0), height: OfficeLength.Centimeters(5.0));

var segmentChart = new OfficeChartData(
    segments.Select(row => row.Segment),
    new[] {
        new OfficeSeries("Q1", segments.Select(row => (double)row.Q1)),
        new OfficeSeries("Q2", segments.Select(row => (double)row.Q2)),
        new OfficeSeries("Q3", segments.Select(row => (double)row.Q3)),
        new OfficeSeries("Q4", segments.Select(row => (double)row.Q4))
    });

dashboard.AddChart(OfficeChartType.Line, segmentChart, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = OfficeIMO.OfficeLength.Centimeters(12.2), Top = OfficeIMO.OfficeLength.Centimeters(3.0), Width = OfficeIMO.OfficeLength.Centimeters(12.0), Height = OfficeIMO.OfficeLength.Centimeters(6.5) })
    .SetTitle("Segment trend")
    .SetLegend(OfficeChartLegendPosition.Bottom);

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
using var presentation = PowerPointDocument.Load("incoming.pptx");

PowerPointFeatureReport report = presentation.InspectFeatures();
report.EnsureNoUnsupportedFeatures();

foreach (PowerPointFeatureFinding feature in report.PreservedFeatures) {
    Console.WriteLine($"{feature.Name}: {feature.Count}");
}
```

### Accessibility, review, and animation inspection

Generated and imported decks can be inspected with structured policies before they enter CI or a publishing workflow:

```csharp
using var presentation = PowerPointDocument.Load("incoming.pptx",
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

### Review threads, custom shows, and VBA projects

Classic comments round-trip through supported PPTX and binary PPT workflows. Modern comments remain an Open XML feature and support replies, status changes, reassignment, and removal. Adding a modern comment writes its explicit slide reference; removing the last comment removes that reference while preserving other slide extensions:

```csharp
var reviewer = new PowerPointCommentAuthor("A. Reviewer", "AR", "reviewer@example.test");
PowerPointModernComment comment = presentation.AddModernComment(
    slide, reviewer, "Verify the release numbers.");
comment.AddReply(reviewer, "Verified against the source report.");

presentation.AddCustomShow("Executive review", new[] { titleSlide, summarySlide });
```

`GetVbaProjectBytes()`, `SetVbaProject(byte[], ...)`, and `RemoveVbaProject()` handle
validated, bounded compound-storage content. `ReadVbaProject()` and the typed
`SetVbaProject(OfficeVbaProject, ...)` overload use the
[shared native VBA source model](../OfficeIMO.Core/README.md#native-vba-projects):

```csharp
using OfficeIMO;

var project = OfficeVbaProject.Create("Automation");
project.AddModule("Helpers",
    "Public Function Value() As Long\r\nValue = 42\r\nEnd Function\r\n");
presentation.SetVbaProject(project);
```

Existing projects can be read, edited, and applied through the same model.
Document modules bound to Word or Excel are rejected before the project part changes.
Changed signed projects require explicit
`OfficeVbaWriteOptions.AllowSignatureRemoval`; unchanged bytes retain signatures.
Save macro content only to matching macro-enabled Open XML package formats such
as `.pptm`, `.potm`, or `.ppsm`. Source editing adds no runtime dependency and
does not execute VBA or edit binary `.ppt` carriers.

`OfficeDocument.Inspect(path, options).VbaSignatures` reports legacy, agile, and V3 profile evidence for `.pptm`, `.potm`, `.ppsm`, and `.ppam`. `ValidateVbaSignatures(...)` adds managed content binding plus provider-backed CMS, certificate, revocation, and timestamp evidence. `SignVbaProject(...)` clears and recreates the three Microsoft profiles atomically on every supported platform. A registered Microsoft Office SIP is an optional Windows differential check. Apply VBA signatures before an OPC package signature.

### SmartArt and visual proof

Use the bounded semantic SmartArt workflows when native editable diagram data is more useful than flattened artwork:

```csharp
PowerPointSmartArt process = slide.AddSmartArt(
    PowerPointSmartArtType.BasicRelationship,
    new[] { "Discover", "Design", "Deliver" });

PowerPointVisualProofReport proof = presentation.InspectVisuals();
proof.RecordArtifact("deck.pptx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    File.ReadAllBytes("deck.pptx"));
proof.SaveJson("deck.visual-proof.json");
```

The bounded semantic layouts cover process, hierarchy, cycle, list, matrix, pyramid, and relationship diagrams. `AddCustomGeometry(...)` accepts shared `OfficeShape` path or polygon geometry and writes editable DrawingML `custGeom` content.

The visual proof report records structural and extraction evidence, accessibility results, shared-snapshot diagnostics, PNG/SVG hashes, caller-supplied conversion artifacts, and perceptual-comparison results. PowerPoint Desktop reference rendering is available through `PowerPointDesktopReferenceRenderer.TryRender(...)` only when the caller explicitly enables it; normal generation and export never use Office automation.

`SetGridSpacing` and guide creation use explicit `OfficeLength` values.
`AddGuide` returns its `PowerPointGuideInfo`; `AddColumnGuides` and `AddRowGuides`
return the guide values they created. The guide position is available as `Position`.

### Designer composition

Use a semantic plan when a deck needs readable business composition without hand-positioning every object.
`PowerPointDeckPlan` owns intent and `PowerPointDocument.Compose(...)` is the only public rendering operation:

Inside custom composition callbacks, `AddTitle`, `AddCardGrid`,
`AddCalloutBand` and the other composite Add methods return the shapes they
create. Each shape remains attached and editable through the object model.

```csharp
using var presentation = PowerPointDocument.Create("proposal.pptx");

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
    OfficeChartType.ColumnClustered,
    new OfficeChartData(
        new[] { "Q1", "Q2", "Q3", "Q4" },
        new[] { new OfficeSeries("Adoption", new[] { 28d, 43d, 61d, 72d }) }),
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

Concrete `PowerPointDocument`, `PowerPointSlide`, and `PowerPointShape` objects are the editing API. There is
no separate PowerPoint builder vocabulary to learn or keep synchronized.

The concrete `PowerPointDocument`, `PowerPointSlide`, and `PowerPointShape` types shown here are the supported editing surface. Semantic plans remain optional composition input and do not replace the concrete document model.

## PDF export

Forward PDF conversion lives in `OfficeIMO.PowerPoint.Pdf`:

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

using var presentation = PowerPointDocument.Load("deck.pptx");
presentation.SaveAsPdf("deck.pdf");
using var pdf = presentation.ToPdf(null, out OfficeIMO.OfficeConversionReport report);
byte[] pdfBytes = pdf.ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
```

Installed document fonts are embedded by default. Inspect `report.Diagnostics`
for substitutions, approximations, and omissions. Use `OfficeIMO.Pdf.PowerPoint`
when reconstructing PDF content as editable slides.

## Managed image export

Slides and presentation batches can be exported as PNG, JPEG, TIFF, lossless WebP, or SVG:

```csharp
using OfficeIMO.Drawing;

byte[] jpeg = slide.ToImage(new PowerPointImageExportOptions {
    Format = OfficeImageExportFormat.Jpeg,
    RasterEncoding = new OfficeRasterEncodingOptions {
        Jpeg = new OfficeJpegEncodeOptions { Quality = 92 }
    }
});

OfficeIMO.OfficeConversionReport report = presentation.SaveAsImages("slide-previews",
    new PowerPointImageExportOptions {
        SlideNumbers = new[] { 1, 2, 3 },
        IncludeCharts = true,
        IncludeTables = true,
        MaximumOutputWidth = 1600,
        MaximumOutputHeight = 900,
        Format = OfficeImageExportFormat.Webp
    });
```

`OfficeIMO.Drawing` owns sizing and the common raster encoders. `PowerPointImageExportOptions` controls backgrounds, pictures, auto shapes, SmartArt, text boxes, tables, charts and hidden shapes. Batch options also select slides and include hidden slides. `CreateVisualSnapshot().TextSpans` exposes text geometry.

`ToImage(options)` returns encoded bytes. Use `ToImage(options, out report)` for fidelity diagnostics or `ToImageAsync(options)` for bytes and report together. `SaveAsImage(pathOrStream, options)` and `SaveAsImages(folder, options)` return `OfficeConversionReport`; `FailOnLoss` rejects diagnosed loss before writing. The async save forms support cancellation, and caller-owned streams stay open. `FileConflictPolicy` controls existing output files.

## Content provenance

Add the `OfficeIMO` front-door package for the shared inspection entry point.

`OfficeIMO.OfficeDocument.Inspect("input.pptx", new OfficeIMO.OfficeInspectOptions { IncludeContent = false, Provenance = new OfficeProvenanceOptions() }).Provenance!` reports C2PA and AI-specific IPTC metadata in the presentation and its supported embedded images. `PowerPointDocument.RemoveProvenance("input.pptx", "clean.pptx")` removes the selected carriers. Signed-package mutation is blocked unless `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` is selected explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

When the presentation is already in memory, inspect its encoded package bytes directly. This overload validates the package and inspects supported provenance without accessing the filesystem:

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;

static OfficeProvenanceReport InspectUploadedPowerPoint(byte[] packageBytes) =>
    OfficeIMO.OfficeDocument.Inspect(new MemoryStream(packageBytes), "upload.pptx",
        new OfficeIMO.OfficeInspectOptions { IncludeContent = false, Provenance = new OfficeProvenanceOptions() }).Provenance!;
```

## Concealed-content inspection and cleanup

`OfficeDocument.Inspect(path, new OfficeInspectOptions { ContentSafety = new OfficeContentSafetyOptions() }).ContentSafety` reports hidden slides/shapes/groups, zero or fully off-canvas geometry, tiny/transparent/low-contrast text, notes, classic and modern comments, alternative text, and Unicode evidence. `PowerPointDocument.RemoveSelectedContent(...)` uses the native PPTX/PPTM or legacy PPT writer and reinspects the output. Existing signature mutation is blocked unless an explicit supported policy is selected.

## Adjacent packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Export presentations as PDF. |
| [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) | Render semantic markup documents to PowerPoint. |

## Related packages

- `OfficeIMO.PowerPoint` provides editable PowerPoint creation and manipulation.
- `OfficeIMO.PowerPoint.Pdf` exports PDF; `OfficeIMO.Pdf.PowerPoint` reconstructs PDF content as editable slides.
- See `OfficeIMO.Examples` for complete showcase decks and design examples.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** Open XML SDK for `.pptx` package mechanics. Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core` supplies shared themes, units and chart contracts. PDF and page-image export live in `OfficeIMO.PowerPoint.Pdf`; the presentation model does not depend on the PDF engine.
- **Security:** Open XML, legacy binary, and VBA signature carriers are inspected and signed-presentation mutations fail safely without a cryptographic dependency. OPC and VBA creation and cryptographic validation accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

### Hyperlink text colors

Assigning `PowerPointTextRun.Color` keeps that explicit color on hyperlinks in readers supporting the native text-color extension. New links without an explicit color use the theme hyperlink color. To use a style-inherited text color, create the link and set `run.HyperlinkUsesTextColor = true`; set it to `false` to select the theme hyperlink color. Clearing `run.Color` preserves this policy, so a link using text color then inherits its color from the surrounding style. The extension is supported by Office 2019 and newer; older readers may ignore it.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 48 | 2 | 0 | 10 | 0 | 1 |
| Read | 47 | 0 | 8 | 4 | 0 | 2 |
| Edit | 21 | 0 | 34 | 6 | 0 | 0 |
| Preserve | 19 | 0 | 34 | 4 | 0 | 0 |
| Inspect | 6 | 0 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 5 | 0 | 0 | 0 | 0 | 0 |
| Convert | 72 | 14 | 0 | 26 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
