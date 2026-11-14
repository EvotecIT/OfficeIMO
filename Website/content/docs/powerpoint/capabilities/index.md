---
title: PowerPoint Feature Support in OfficeIMO
description: "Which PowerPoint features OfficeIMO.PowerPoint can author, edit, preserve and render, and how to check a deck or legacy PPT file before publishing."
order: 38
packages:
  - OfficeIMO.PowerPoint
cmdlets:
  - Test-OfficePowerPoint
  - New-OfficePowerPointInspectionOptions
search_terms:
  - powerpoint features
  - pptx support
  - ppt binary support
  - preflight check
  - supported charts
verified: 2026-11-14
---

OfficeIMO.PowerPoint is strongest when the output must stay editable and the workflow can use tested Open XML features. The table below separates native authoring from preservation and preview support so a successful save is not confused with full edit fidelity.

| Capability | Author or edit | Preserve on unrelated edits | PNG/SVG and PDF proof | Notes |
|---|---|---|---|---|
| Slides, ordering, sections, size, metadata | Yes | Yes | Yes | Includes duplicate and cross-deck slide import workflows. |
| Text boxes, rich runs, bullets, links, notes | Yes | Yes | Yes | `InspectPreflight()` measures bounds and reports clipped or unreadably reduced text. |
| Auto shapes, lines, connectors, groups, custom paths and polygons | Yes | Yes | Yes, within reported renderer limits | Shared `OfficeShape` paths and polygons write editable DrawingML custom geometry. Effects may still be approximated in fixed-layout exports. |
| Pictures and SVG assets | Yes | Yes | Yes | Crop and transforms are supported; broken image relationships are preflight errors. |
| Native tables | Yes | Yes | Yes | `AddTableSlides(...)` creates deterministic continuation pages and repeats headers. |
| Semantic story families | Yes | Yes | Yes | Executive, chart, comparison, screenshot, appendix, architecture, and closing stories each have two editable compositions. |
| Deck rhythm inspection | Yes, before rendering | Not applicable | Not applicable | Reports repeated kinds/variants, dense streaks, long sections, weak openings, and missing closings with stable codes. |
| All 17 shared chart kinds | Yes | Yes | Yes | Clustered/stacked/100% column and bar, line and area variants, scatter, bubble, radar, pie, and doughnut use `OfficeChartType`. Bubble charts preserve X, Y, and size values in native chart XML and the embedded workbook. |
| Categorical combo charts and secondary value axes | Yes | Yes | Yes | Per-series `RenderKind` and `OfficeChartAxisGroup` stay in native chart XML with cached values and embedded workbook data. |
| Chart accessibility and data summary | Yes | Yes | Yes | Native alt text can include a deterministic data summary; `SaveDataSummary(...)` writes a UTF-8 sidecar. |
| Other chart XML already present in a deck | Limited edit surface | Yes | Snapshot/export support varies | Inspect and test the concrete family before promising mutation parity. |
| Themes, masters, layouts, placeholders | Inspect, select, copy from `.pptx`/`.potx`, and map by semantic name | Yes | Inherited content is included where supported | Template inventory includes brand tokens, footer content, assets, and safe areas. |
| Transitions | Yes | Yes | Static proof only | Preview exports show the slide state, not animated playback. |
| Semantic SmartArt | Yes: process, hierarchy, cycle, list, matrix, pyramid, relationship | Yes | Yes | Node text stays editable; other imported SmartArt layouts remain preservation/reporting scenarios. |
| Classic and modern comments | Yes | Yes | Metadata proof | Classic comments cover PPTX and supported PPT round trips. Modern Open XML comments support replies, status, reassignment, and removal. |
| Shape, text, and chart animations | Yes for supported effects | Broader timelines are preserved and reported | Static proof only | Image and PDF exports prove the static slide state, not playback timing. |
| Custom shows | Yes | Yes | Structural proof | Slide membership, order, rename, and removal use the presentation API. |
| Audio, video, embedded and linked content | OLE storage has a typed add/replace surface; other media mutation varies | Yes | Reported or represented by a fallback | Do not treat a preview as playback proof. |
| VBA projects | Add, replace, extract, and remove opaque projects | Yes in matching macro-enabled formats | Structural proof only | OfficeIMO validates bounded compound storage but does not generate or interpret VBA source. |
| Custom XML and ActiveX | No general authoring contract | Preserve/report where possible | Not a fidelity promise | Use `InspectFeatures()` before edit-heavy round trips. |
| Digital signatures | Inspect/report only | Save policy applies | Not applicable | Editing a signed package can invalidate its signature. |

## Check a generated deck before publishing

`InspectPreflight` measures text and shape bounds and reports clipped or unreadably small text, shapes outside the slide, and broken picture links. Finding codes are stable, so CI can fail on them; the messages are for people.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("deck.pptx");

PowerPointDeckPreflightReport report = presentation.InspectPreflight();
report.SaveJson("deck.preflight.json");

var options = new PowerPointDeckPreflightOptions {
    FailureSeverity = PowerPointDeckPreflightSeverity.Error
};
presentation.InspectPreflight(options).ThrowIfFindings(options.FailureSeverity);
```

### PowerShell

```powershell
Get-ChildItem .\Publish -Filter *.pptx | Test-OfficePowerPoint
```

## Check an imported deck for features outside the editable model

`InspectFeatures` lists package content such as ActiveX, custom XML, media, and macros that OfficeIMO preserves but does not edit. Run it before an edit-heavy round trip.

### C#

```csharp
using OfficeIMO.PowerPoint;

using PowerPointDocument presentation = PowerPointDocument.Load("imported.pptx");
PowerPointFeatureReport features = presentation.InspectFeatures();
foreach (PowerPointFeatureFinding finding in features.PreservedFeatures.Concat(features.UnsupportedFeatures)) {
    Console.WriteLine($"{finding.Category} / {finding.Name}: {finding.SupportLevel} ({finding.Count})");
}
```

### PowerShell

```powershell
$policy = New-OfficePowerPointInspectionOptions -Features -Visuals
Test-OfficePowerPoint -Path .\imported.pptx -Options $policy
```

## Query PowerPoint 97-2003 support

Binary `.ppt`, `.pot`, and `.pps` support is reported in four directions: import into the editable model, new binary authoring, binary round trip, and conversion from PPTX. Each direction is `Native`, `Preserved`, `Converted`, or `Blocked`. Query the versioned catalog instead of assuming support from a successful open.

### C#

```csharp
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(LegacyPptFeature.RasterPictures);
Console.WriteLine(capability);
File.WriteAllText("ppt-capabilities.json", LegacyPptCapabilityCatalog.ToJson());
```

Saving blocks known loss by default. Imported files keep unrelated binary records and streams; tables, charts, and SmartArt convert to static pictures only when you accept that loss explicitly, and features without a safe mapping stay blocked.
