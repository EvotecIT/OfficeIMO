---
title: PowerPoint Capability Matrix
description: Tested authoring, editing, preservation, rendering, and reporting boundaries for OfficeIMO.PowerPoint.
order: 35
---

OfficeIMO.PowerPoint is strongest when the output must stay editable and the workflow can use tested Open XML features. The table below separates native authoring from preservation and preview support so a successful save is not confused with full edit fidelity.

| Capability | Author or edit | Preserve on unrelated edits | PNG/SVG and PDF proof | Notes |
|---|---|---|---|---|
| Slides, ordering, sections, size, metadata | Yes | Yes | Yes | Includes duplicate and cross-deck slide import workflows. |
| Text boxes, rich runs, bullets, links, notes | Yes | Yes | Yes | `InspectPreflight()` measures bounds and reports clipped or unreadably reduced text. |
| Auto shapes, lines, connectors, groups | Yes | Yes | Yes, within reported renderer limits | Custom geometry and effects may be approximated in fixed-layout exports. |
| Pictures and SVG assets | Yes | Yes | Yes | Crop and transforms are supported; broken image relationships are preflight errors. |
| Native tables | Yes | Yes | Yes | `AddTableSlides(...)` creates deterministic continuation pages and repeats headers. |
| Semantic story families | Yes | Yes | Yes | Executive, chart, comparison, screenshot, appendix, architecture, and closing stories each have two editable compositions. |
| Deck rhythm inspection | Yes, before rendering | Not applicable | Not applicable | Reports repeated kinds/variants, dense streaks, long sections, weak openings, and missing closings with stable codes. |
| All 16 shared chart kinds | Yes | Yes | Yes | Clustered/stacked/100% column and bar, line and area variants, scatter, radar, pie, and doughnut use `OfficeChartKind`. |
| Categorical combo charts and secondary value axes | Yes | Yes | Yes | Per-series `RenderKind` and `OfficeChartAxisGroup` stay in native chart XML with cached values and embedded workbook data. |
| Chart accessibility and data summary | Yes | Yes | Yes | Native alt text can include a deterministic data summary; `SaveDataSummary(...)` writes a UTF-8 sidecar. |
| Other chart XML already present in a deck | Limited edit surface | Yes | Snapshot/export support varies | Inspect and test the concrete family before promising mutation parity. |
| Themes, masters, layouts, placeholders | Inspect, select, copy from `.pptx`/`.potx`, and map by semantic name | Yes | Inherited content is included where supported | Template inventory includes brand tokens, footer content, assets, and safe areas. |
| Transitions | Yes | Yes | Static proof only | Preview exports show the slide state, not animated playback. |
| SmartArt and advanced diagrams | Limited inspection/rendering | Yes | May be approximated or reported | Prefer native shapes for generated diagrams that must be predictably editable. |
| Audio, video, embedded and linked content | Limited mutation | Yes | Reported or represented by a fallback | Do not treat a preview as playback proof. |
| Comments, advanced animation timelines, custom XML, ActiveX, macros | No general authoring contract | Preserve/report where possible | Not a fidelity promise | Use `InspectFeatures()` before edit-heavy round trips. |
| Digital signatures | Inspect/report only | Save policy applies | Not applicable | Editing a signed package can invalidate its signature. |

## Generation checks

```csharp
PowerPointDeckPreflightReport report = presentation.InspectPreflight();
report.SaveJson("deck.preflight.json");

var options = new PowerPointDeckPreflightOptions {
    FailureSeverity = PowerPointDeckPreflightSeverity.Error
};
PowerPointDeckPreflightReport gate = presentation.InspectPreflight(options);
gate.ThrowIfFindings(options.FailureSeverity);
presentation.Save();
```

The preflight report is the same contract returned by designer generation and `OfficeIMO.Markup.PowerPoint`. Stable finding codes are intended for CI policy; human messages are supporting context.

For imported presentations, combine layout preflight with `InspectFeatures()` so the pipeline sees both generated-content risks and package features outside the editable surface.

## PowerPoint 97-2003 capability contract

Binary `.ppt`, `.pot`, and `.pps` support has four independently reported directions: import into the editable
model, new binary authoring, binary round-trip, and PPTX-to-binary conversion. Query the versioned catalog
instead of inferring support from a successful open:

```csharp
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;

LegacyPptCapability capability = LegacyPptCapabilityCatalog.Get(
    LegacyPptFeature.RasterPictures);
string json = LegacyPptCapabilityCatalog.ToJson();
```

Each direction is `Native`, `Preserved`, `Converted`, or `Blocked`. There are no provisional catalog rows.
Saving blocks known loss by default. Imported files preserve unrelated compound-file records and streams;
tables, charts, and SmartArt can convert to static PNG visuals only with explicit loss acceptance; features
without a safe mapping remain blocked.
