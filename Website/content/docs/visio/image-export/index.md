---
title: "Export Visio Diagrams to Images and PDF"
description: "Render VSDX diagram pages to SVG, PNG, JPEG, TIFF, WebP or PDF in .NET with OfficeIMO.Visio, without Microsoft Visio installed on the machine."
meta.seo_title: "Convert VSDX to SVG, PNG, JPEG, WebP, and PDF"
order: 56
packages:
  - OfficeIMO.Visio
  - OfficeIMO.Visio.Pdf
cmdlets:
  - Export-OfficeDocumentImage
  - New-OfficeVisioImageOptions
  - Convert-OfficeDocument
commands:
  - officeimo convert
search_terms:
  - vsdx to svg
  - vsdx to png
  - vsdx to pdf
  - visio preview
  - convert visio
verified: 2026-11-14
---

`OfficeIMO.Visio` renders diagram pages to images with `SaveAsImage`, `SaveAsImages`, and `ToImages`, configured by `VisioImageExportOptions`. `OfficeIMO.Visio.Pdf` adds `SaveAsPdf`. Neither starts Microsoft Visio. An image is a rendering, not a substitute for [validating the VSDX](/docs/visio/editing-and-validation/).

## Export a page to SVG or PNG

`Format` picks the encoder. `PageIndex` is zero-based; the first page is the default.

### C#

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Load("pipeline.vsdx");

document.SaveAsImage("pipeline.svg", new VisioImageExportOptions {
    Format = OfficeImageExportFormat.Svg
});

document.SaveAsImage("pipeline.png", new VisioImageExportOptions {
    Format = OfficeImageExportFormat.Png,
    TargetDpi = 144,
    Supersampling = 3
});
```

### PowerShell

```powershell
Export-OfficeDocumentImage -Path .\pipeline.vsdx -Page 1 -Format Svg -OutputPath .\pipeline.svg
Export-OfficeDocumentImage -Path .\pipeline.vsdx -Page 1 -Format Png -Dpi 144 -OutputPath .\pipeline.png
```

### CLI

```sh
officeimo convert pipeline.vsdx pipeline.svg
officeimo convert pipeline.vsdx pipeline.png
```

## Export every page

### C#

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Load("runbook.vsdx");

document.SaveAsImages("runbook-pages", new VisioImageExportOptions {
    Format = OfficeImageExportFormat.Jpeg
});

IReadOnlyList<OfficeImageExportResult> previews = document.ToImages(new VisioImageExportOptions {
    Format = OfficeImageExportFormat.Webp
});

foreach (OfficeImageExportResult preview in previews) {
    Console.WriteLine($"{preview.Width}x{preview.Height}, {preview.EncodedLength} bytes");
}
```

### PowerShell

```powershell
$options = New-OfficeVisioImageOptions -Font Arial -NoConnectorLabels
Export-OfficeDocumentImage -Path .\runbook.vsdx -Format Jpeg -OutputDirectory .\runbook-pages -Options $options
```

## Export to PDF

### C#

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

using VisioDocument document = VisioDocument.Load("architecture.vsdx");
document.SaveAsPdf("architecture.pdf");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\architecture.vsdx -OutputPath .\architecture.pdf
```

### CLI

```sh
officeimo convert architecture.vsdx architecture.pdf
```
