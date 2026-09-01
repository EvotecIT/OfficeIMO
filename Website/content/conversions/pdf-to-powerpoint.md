---
title: "Convert PDF pages to PowerPoint"
description: "Import PDF pages into an editable PPTX projection, retain warnings for omitted page content, and avoid presenting the result as the original slide deck."
meta.workflow_id: "pdf-pptx"
meta.eyebrow: "PDF page import"
meta.source_format: "PDF"
meta.destination_format: "PPTX"
meta.package: "OfficeIMO.PowerPoint.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.PowerPoint.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=convert&route=pdf-pptx"
meta.primary_label: "Convert PDF to PowerPoint in the browser"
meta.secondary_url: "/docs/pdf/conversion/"
meta.secondary_label: "Read the conversion guide"
meta.summary_title: "PDF to PowerPoint summary"
meta.limit: "Editable mode reconstructs supported page objects; visual mode creates page images. Neither mode can recover original themes, animations, notes, charts, groups, or authoring intent."
meta.related_url: "/pdf/"
meta.related_label: "Browse PDF tools and imports"
meta.howto.name: "Import PDF pages into an editable presentation"
meta.howto.description: "Project each supported page into PowerPoint structures and retain loss and omission warnings."
meta.howto.steps:
  - name: "Select"
    text: "Choose a readable PDF or load the built-in sample."
  - name: "Convert"
    text: "Project the PDF pages into an editable PPTX profile."
  - name: "Review"
    text: "Download the presentation and companion report, then inspect every warning."
---

A PDF contains final page presentation, not the PowerPoint theme, slide master, animations, speaker notes, or original editable object graph. `OfficeIMO.PowerPoint.Pdf` creates a new presentation from supported page content and reports what could not be reconstructed.

## Convert from .NET

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;

PdfDocument pdf = PdfDocument.Load("briefing.pdf");
PdfPowerPointConversionResult result = pdf.ToPowerPointPresentationResult();
using var presentation = result.Value;

File.WriteAllBytes("briefing.pptx", presentation.ToBytes());
foreach (PdfConversionWarning warning in result.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The .NET API and browser default to editable-content reconstruction. The browser also exposes three explicit alternatives:

- **Editable content** creates native text boxes, detected tables, safe basic shapes, and separate supported images.
- **Visual pages** creates one page-sized image per slide. Its text, shapes, charts, and tables are not editable.
- **Visual + editable tables** keeps the page image and overlays detected native tables.
- **Tables only** creates native tables and intentionally omits other page content.

The companion report names the selected projection, surfaces renderer capability diagnostics, and reports fidelity as degraded or partial when content is omitted or simplified.

## Use the result as a new deck

Treat the output as an editable starting point, review artifact, or migration aid. Do not claim that it recovers the original presentation. Validate slide size, text, images, reading order, and the features that matter to the intended audience before publishing.
