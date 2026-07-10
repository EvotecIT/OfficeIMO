---
title: PowerPoint Capability Matrix
description: Tested authoring, editing, preservation, rendering, and reporting boundaries for OfficeIMO.PowerPoint.
order: 35
---

# PowerPoint Capability Matrix

OfficeIMO.PowerPoint is strongest when the output must stay editable and the workflow can use tested Open XML features. The table below separates native authoring from preservation and preview support so a successful save is not confused with full edit fidelity.

| Capability | Author or edit | Preserve on unrelated edits | PNG/SVG and PDF proof | Notes |
|---|---|---|---|---|
| Slides, ordering, sections, size, metadata | Yes | Yes | Yes | Includes duplicate and cross-deck slide import workflows. |
| Text boxes, rich runs, bullets, links, notes | Yes | Yes | Yes | `Preflight()` measures bounds and reports clipped or unreadably reduced text. |
| Auto shapes, lines, connectors, groups | Yes | Yes | Yes, within reported renderer limits | Custom geometry and effects may be approximated in fixed-layout exports. |
| Pictures and SVG assets | Yes | Yes | Yes | Crop and transforms are supported; broken image relationships are preflight errors. |
| Native tables | Yes | Yes | Yes | `AddTableSlides(...)` creates deterministic continuation pages and repeats headers. |
| Column/bar, line, scatter, pie, doughnut charts | Yes | Yes | Yes | Chart data stays in native chart XML and the embedded workbook. |
| Other chart XML already present in a deck | Limited edit surface | Yes | Snapshot/export support varies | Inspect and test the concrete family before promising mutation parity. |
| Themes, masters, layouts, placeholders | Inspect and select; theme tokens can be changed | Yes | Inherited content is included where supported | Template consumption is safer than rebuilding corporate masters in code. |
| Transitions | Yes | Yes | Static proof only | Preview exports show the slide state, not animated playback. |
| SmartArt and advanced diagrams | Limited inspection/rendering | Yes | May be approximated or reported | Prefer native shapes for generated diagrams that must be predictably editable. |
| Audio, video, embedded and linked content | Limited mutation | Yes | Reported or represented by a fallback | Do not treat a preview as playback proof. |
| Comments, advanced animation timelines, custom XML, ActiveX, macros | No general authoring contract | Preserve/report where possible | Not a fidelity promise | Use `InspectFeatures()` before edit-heavy round trips. |
| Digital signatures | Inspect/report only | Save policy applies | Not applicable | Editing a signed package can invalidate its signature. |

## Generation checks

```csharp
PowerPointDeckPreflightReport report = presentation.Preflight();
report.SaveJson("deck.preflight.json");

// Use this when errors should stop the output pipeline before Save().
presentation.SaveWithPreflight(new PowerPointDeckPreflightOptions {
    FailureSeverity = PowerPointDeckPreflightSeverity.Error
});
```

The preflight report is the same contract returned by designer generation and `OfficeIMO.Markup.PowerPoint`. Stable finding codes are intended for CI policy; human messages are supporting context.

For imported presentations, combine layout preflight with `InspectFeatures()` so the pipeline sees both generated-content risks and package features outside the editable surface.
