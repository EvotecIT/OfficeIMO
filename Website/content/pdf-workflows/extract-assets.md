---
title: "Extract assets from a PDF"
description: "Save images, embedded fonts, ICC profiles and page SVGs from a PDF, with HTML and JSON reports for image DPI, spot inks and page boxes."
meta.workflow_id: "extract-assets"
meta.eyebrow: "Understand a PDF"
meta.source_format: "PDF"
meta.destination_format: "ZIP archive with assets and reports"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/browser/extract-pdf-assets/"
meta.primary_label: "Extract PDF assets in the browser"
meta.secondary_url: "/docs/cli/pdf/#extract-pdf-assets"
meta.secondary_label: "Extract assets from the CLI"
meta.summary_title: "Asset bundle"
meta.limit: "The browser accepts one PDF up to 25 MiB and 500 pages. Extracted files are limited to 64 MiB before ZIP compression."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Download a PDF asset bundle"
meta.howto.description: "Extract files and an inventory from a PDF on your device."
meta.howto.steps:
  - name: "Choose a PDF"
    text: "Open a local PDF or try the built-in sample. The browser processes it on your device."
  - name: "Extract assets"
    text: "Keep the default RGBA export for supported soft masks or choose standard image formats under More options."
  - name: "Open the report"
    text: "Download and unpack the ZIP, then open report.html to review files and page details."
---

The bundle includes distinct image and embedded font files, ICC profiles and an SVG for each page. Open `report.html` to review image placement DPI, spot inks, page boxes in millimetres, font embedding and subsets, and extraction diagnostics. Use `report.json` to read the same inventory from a script. Font subsets contain the glyphs embedded in the PDF.

## Extract from .NET

```csharp
using OfficeIMO.Pdf;

using var source = PdfDocument.Load("brochure.pdf");
source.ExtractAssets("brochure-assets");
```

Choose a new output folder. For an archive or other stream-based destination, call `ExtractAssetsTo` with a factory that opens a writable stream for each relative path.
