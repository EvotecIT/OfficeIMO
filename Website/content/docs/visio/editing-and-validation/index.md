---
title: "Edit and validate existing Visio diagrams"
description: "Load VSDX files, select and update shapes, edit Shape Data and layers, query topology, save round trips, and validate generated diagrams."
meta.seo_title: "Edit and validate existing VSDX Visio diagrams in .NET"
order: 54
---

Loaded diagrams use the same page, shape, connector, data, and fluent APIs as newly created files. You can select shapes by ID, text, layer, master, or Shape Data and then apply a consistent update.

## Update matching shapes

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Color = OfficeIMO.Drawing.OfficeColor;

VisioDocument.Load("operations.vsdx")
    .AsFluent()
    .ExistingPage("Operations", page => page
        .ShapesWithData("Owner", "Ops", selection => selection
            .Fill(Color.LightBlue)
            .ShapeData(
                "Reviewed",
                "Yes",
                "Reviewed",
                VisioShapeDataType.Boolean))
        .ShapesContainingText("Legacy", selection => selection
            .Text(shape => shape.Text!.Replace(
                "Legacy",
                "Production",
                StringComparison.Ordinal))))
    .End()
    .Save("operations.updated.vsdx");
```

## Work with layers and Shape Data

```csharp
var document = VisioDocument.Create("architecture.vsdx");
var page = document.AddPage("Architecture");

page.AddLayer("Infrastructure");
page.AddLayer("Annotations").Print = false;

var server = page.AddProcess(2, 5, 2, 1, "API server");
server.SetShapeData(
    "Owner",
    "Platform",
    "Owner",
    VisioShapeDataType.String,
    "Owning support team");

page.AddToLayer("Infrastructure", server);
document.Save();
```

## Validate in memory or from disk

```csharp
IReadOnlyList<string> inMemoryIssues = document.Validate();
IReadOnlyList<string> savedIssues =
    VisioValidator.Validate("architecture.vsdx");
```

Validation checks package and diagram invariants understood by OfficeIMO. A successful validation does not claim that every optional feature in every Visio viewer is identical. When desktop Microsoft Visio is available, the example and test tooling can add an application-level open/save check without making it a runtime requirement.

Use [image export](/docs/visio/image-export/) to add a human-reviewable preview to automated validation.
