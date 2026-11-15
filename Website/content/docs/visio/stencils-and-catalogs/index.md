---
title: "Use Visio Stencils and Shape Catalogs"
description: "Place shapes from OfficeIMO's built-in catalogs or from VSSX and VSTX stencil packs you supply, and keep the link to each shape's master."
meta.seo_title: "Use VSSX and VSTX Visio stencils from .NET"
order: 55
packages:
  - OfficeIMO.Visio
cmdlets:
  - New-OfficeVisio
  - Import-OfficeVisioStencilCatalog
  - Add-OfficeVisioStencilShape
  - Add-OfficeVisioConnector
  - Find-OfficeVisioStencil
  - Export-OfficeVisioStencilPreviewGallery
limits: "OfficeIMO does not bundle third-party stencil packs or infer their license; license and attribution are metadata you supply."
search_terms:
  - vssx
  - vstx
  - visio stencil
  - visio shapes
  - visio masters
verified: 2026-11-14
---

`VisioStencils` (namespace `OfficeIMO.Visio.Stencils`) has built-in catalogs such as `Flowchart`, `Network`, and `BasicShapes`. `VisioStencilPackageCatalog` loads masters from `.vssx` or `.vstx` packages you supply. A placed shape keeps its link to the master, so the diagram stays editable and can be migrated later.

## Place shapes from a built-in catalog

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

using VisioDocument document = VisioDocument.Create("network.vsdx");
VisioPage page = document.AddPage("Network", OfficeLength.Inches(11), OfficeLength.Inches(8.5));

VisioShape server = page.AddStencilShape(VisioStencils.Network.Get("server"), "server",
    OfficeLength.Inches(2), OfficeLength.Inches(4), "Application server");
VisioShape database = page.AddStencilShape(VisioStencils.Network.Get("database"), "database",
    OfficeLength.Inches(6), OfficeLength.Inches(4), "Orders database");

page.AddConnector(server, database, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);

document.Save();
```

### PowerShell

```powershell
New-OfficeVisio -Path .\network.vsdx -UseMastersByDefault {
    Import-OfficeVisioStencilCatalog -BuiltIn Network -Name Net -Default
    Add-OfficeVisioStencilShape -Catalog Net -Stencil server -Key server -Text 'Application server' -X 2 -Y 4
    Add-OfficeVisioStencilShape -Catalog Net -Stencil database -Key database -Text 'Orders database' -X 6 -Y 4
    Add-OfficeVisioConnector -From server -To database -Kind RightAngle
}
```

## Find a shape in a catalog

### C#

```csharp
using OfficeIMO.Visio.Stencils;

foreach (VisioStencilShape shape in VisioStencils.Flowchart.Shapes.Take(10)) {
    Console.WriteLine($"{shape.Id}: {shape.Name}");
}

if (VisioStencils.Flowchart.TryGet("decision", out VisioStencilShape? decision)) {
    Console.WriteLine(decision!.Name);
}
```

### PowerShell

```powershell
Find-OfficeVisioStencil -BuiltIn Flowchart -Query process -First 5
```

## Load a VSSX or VSTX stencil pack

Use a stencil pack only when your application is allowed to use it. `SourceLicense` and `SourceAttribution` are carried into shapes, catalogs, and manifests; OfficeIMO does not check them. Masters that OfficeIMO cannot author stay marked as unsupported when `IncludeUnsupportedMasters` is set.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;

VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load("MyShapes.vssx", new VisioStencilPackageLoadOptions {
    CatalogName = "My Shapes",
    SourceLicense = "License identifier or notice for this pack",
    SourceAttribution = "Required attribution",
    IncludeUnsupportedMasters = true
});

Console.WriteLine($"{catalog.Shapes.Count} masters in {catalog.Name}");

using VisioDocument document = VisioDocument.Create("custom.vsdx");
VisioPage page = document.AddPage("Custom", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
page.AddStencilShape(catalog.Shapes[0], "first", OfficeLength.Inches(3), OfficeLength.Inches(4), catalog.Shapes[0].Name);
document.Save();
```

On Windows, `VisioStencilPackageCatalog.DiscoverInstalledVisioPackages()` lists the stencil packs of an installed Visio, and `LoadMany` loads several packs into one catalog.

### PowerShell

```powershell
$catalog = Import-OfficeVisioStencilCatalog -Path .\MyShapes.vssx -CatalogName 'My Shapes' -IncludeUnsupportedMasters
Find-OfficeVisioStencil -Catalog $catalog -First 10
```

## Preview a stencil pack

### PowerShell

Write an HTML gallery of every master in a pack:

```powershell
Export-OfficeVisioStencilPreviewGallery -Path .\MyShapes.vssx -OutputDirectory .\StencilGallery -Title 'My shapes'
```

## Keep stencil mappings stable

1. Look up masters by a stable catalog key, not by display order.
2. Record the source stencil and the master you expect.
3. Keep application data separate from shape placement.
4. Validate the generated VSDX after upgrading a stencil pack.
