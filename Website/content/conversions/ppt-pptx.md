---
title: "Convert PPT and PPTX in .NET"
description: "Modernize PowerPoint 97–2003 PPT presentations or create supported PPT output from PPTX with feature-level conversion diagnostics."
meta.eyebrow: "PowerPoint conversion"
meta.source_format: "PPT"
meta.destination_format: "PPTX"
meta.package: "OfficeIMO"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.primary_label: "Get OfficeIMO"
meta.primary_url: "https://www.nuget.org/packages/OfficeIMO"
meta.secondary_label: "Check PowerPoint compatibility"
meta.secondary_url: "/compatibility/#powerpoint"
meta.summary_title: "PPT and PPTX summary"
meta.limit: "Legacy and modern PowerPoint formats differ. Review diagnostics for masters, media, charts, animation, embedded objects, and drawing effects."
meta.related_label: "Explore all conversion guides"
meta.related_url: "/convert/guides/"
meta.howto.name: "Convert PPT to PPTX with OfficeIMO"
meta.howto.description: "Analyze a legacy presentation and create a modern PPTX file without PowerPoint automation."
meta.howto.steps:
  - name: "Install"
    text: "Add the OfficeIMO NuGet package."
  - name: "Analyze"
    text: "Review slide, shape, media, chart, and preservation findings."
  - name: "Convert"
    text: "Call OfficeConverter.Convert for the PPT and PPTX paths."
  - name: "Review"
    text: "Validate the slides important to the presentation workflow."
---

`OfficeIMO.PowerPoint` handles legacy PPT, POT, and PPS files as well as modern PPTX, PPTM, POTX, and PPSX presentations. Applications can use the same conversion surface in services, desktop software, containers, and scheduled jobs without launching Microsoft PowerPoint.

File-to-file conversion uses `OfficeConverter` from the `OfficeIMO` package.

## PPT to PPTX

```csharp
using OfficeIMO;

OfficeConversionReport preview =
    OfficeConverter.Analyze("briefing.ppt", "briefing.pptx");

OfficeConversionReport result =
    OfficeConverter.Convert("briefing.ppt", "briefing.pptx");
```

## PPTX to PPT

```csharp
using OfficeIMO;

OfficeConversionReport result =
    OfficeConverter.Convert("briefing.pptx", "briefing.ppt");
```

Slides contain more than text boxes. Masters, layouts, groups, charts, embedded objects, custom XML, animation, media, and drawing effects can cross format generations differently. The compatibility report makes those decisions visible and can block output when the selected policy does not allow the required fallback.

Use PPT to PPTX for archive modernization and continued editing. Use PPTX to PPT only for a real legacy consumer, and test the destination in that consumer when the deck is business-critical.

See [PowerPoint compatibility](/compatibility/#powerpoint) and the [PowerPoint guide](/docs/powerpoint/) for slide creation, editing, notes, charts, and image export.
