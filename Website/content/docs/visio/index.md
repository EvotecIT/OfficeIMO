---
title: "Create, Edit, and Export Visio Diagrams"
description: "OfficeIMO.Visio creates, edits, validates, and exports VSDX diagrams without Microsoft Visio installed."
meta.seo_title: "Create, edit, and export VSDX Visio diagrams in .NET"
order: 52
packages:
  - OfficeIMO.Visio
cmdlets:
  - New-OfficeVisio
  - Add-OfficeVisioRectangle
  - Add-OfficeVisioConnector
commands:
  - officeimo convert
search_terms:
  - vsdx
  - visio diagram
  - create visio file
  - visio without visio
verified: 2026-11-14
---

`OfficeIMO.Visio` creates, edits, validates, and exports `.vsdx` diagrams without COM automation or Microsoft Visio. Use a diagram builder when your data already describes a topology or process, or work directly with pages, shapes, connectors, layers, Shape Data, and containers.

| You want to | Page |
|---|---|
| Generate flowcharts, architecture, network, org chart, sequence, timeline, or graph diagrams | [Generate Visio Diagrams with Builders](/docs/visio/diagram-builders/) |
| Load a VSDX, update shapes, edit Shape Data and layers, or validate output | [Edit and Validate Visio Diagrams](/docs/visio/editing-and-validation/) |
| Place shapes from built-in catalogs or VSSX/VSTX stencil packs | [Use Visio Stencils and Shape Catalogs](/docs/visio/stencils-and-catalogs/) |
| Export SVG, PNG, JPEG, TIFF, WebP, or PDF | [Export Visio Diagrams to Images and PDF](/docs/visio/image-export/) |

## Create a diagram

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;

using VisioDocument document = VisioDocument.Create("service-map.vsdx");
VisioPage page = document.AddPage("Services", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

VisioShape web = page.AddRectangle(OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Web");
VisioShape api = page.AddRectangle(OfficeLength.Inches(5), OfficeLength.Inches(4), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "API");

VisioConnector call = page.AddConnector(web, api, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
call.EndArrow = EndArrow.Triangle;
call.Label = "calls";

document.Save();
```

### PowerShell

```powershell
New-OfficeVisio -Path .\service-map.vsdx -Title 'Service map' {
    Add-OfficeVisioRectangle -Key web -Text 'Web' -X 2in -Y 4in
    Add-OfficeVisioRectangle -Key api -Text 'API' -X 5in -Y 4in
    Add-OfficeVisioConnector -From web -To api -Kind RightAngle -EndArrow Triangle -Label 'calls'
}
```

### CLI

Render the diagram to SVG or PDF:

```sh
officeimo convert service-map.vsdx service-map.svg
officeimo convert service-map.vsdx service-map.pdf
```
