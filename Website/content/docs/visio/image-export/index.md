---
title: "Export Visio diagrams to images"
description: "Render VSDX pages to SVG, PNG, JPEG, TIFF, or lossless WebP previews without Microsoft Visio installed."
meta.seo_title: "Convert VSDX to SVG, PNG, JPEG, and WebP"
order: 56
---

OfficeIMO.Visio renders diagram pages without launching Microsoft Visio. Use image output for pull-request previews, evidence packs, web portals, thumbnails, and automated visual review.

## Export SVG and PNG

```csharp
using OfficeIMO.Visio;

var document = VisioDocument.Create("pipeline.vsdx");
var page = document.AddPage("Pipeline").Size(8, 4);
var build = page.AddProcess(1.5, 2, 1.4, 0.7, "Build");
var ship = page.AddProcess(5.5, 2, 1.4, 0.7, "Ship");

page.AddConnector(
    build,
    ship,
    ConnectorKind.RightAngle,
    VisioSide.Right,
    VisioSide.Left).EndArrow = EndArrow.Arrow;

document.SaveAsSvg(
    "pipeline.svg",
    new VisioSvgSaveOptions {
        PixelsPerInch = 96,
        BackgroundColor = null
    });

document.SaveAsPng(
    "pipeline.png",
    new VisioPngSaveOptions {
        PixelsPerInch = 144,
        Supersampling = 3
    });
```

## Use the shared image builder

```csharp
OfficeImageExportResult webp = document
    .ToImage()
    .AtDpi(144)
    .AsWebp()
    .Save("pipeline.webp");

IReadOnlyList<OfficeImageExportResult> pages = document
    .ToImages()
    .AllPages()
    .AsJpeg()
    .Save("pipeline-pages");
```

The image builder also supports PNG, TIFF, and SVG. Choose DPI and supersampling based on the destination: lower values for navigation thumbnails, higher values for reports and visual inspection.

Image export is a rendering of the diagram, not a replacement for package validation. Run `document.Validate()` or `VisioValidator.Validate(path)` on the VSDX and use the image as additional human-readable evidence.
