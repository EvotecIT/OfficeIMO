# OfficeIMO.Visio

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Visio)](https://www.nuget.org/packages/OfficeIMO.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Visio)

Create, edit, inspect, validate and export Visio diagrams in .NET without Microsoft Visio or COM. `OfficeIMO.Visio` works with `.vsdx`, `.vstx`, `.vssx` and their macro-enabled forms, and imports legacy binary `.vsd` and Visio XML drawings.

## Install

```sh
dotnet add package OfficeIMO.Visio
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Visio/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.Visio;

using var document = VisioDocument.Create("Flow.vsdx");
document.Properties.Title = "Approval flow";
var page = document.AddPage("Page-1");
page.AddTitle("Approval flow");
page.AddRectangle("start", OfficeLength.Inches(1), OfficeLength.Inches(1),
    OfficeLength.Inches(2), OfficeLength.Inches(1), "Start").SetBold();
page.AddDiamond("decision", OfficeLength.Inches(4), OfficeLength.Inches(1.5),
    OfficeLength.Inches(2), OfficeLength.Inches(2), "Approved?");
page.AddEllipse("end", OfficeLength.Inches(7), OfficeLength.Inches(1.5),
    OfficeLength.Inches(2), OfficeLength.Inches(1), "Done");
page.AddConnector("start", "decision", fromSide: VisioSide.Right, toSide: VisioSide.Left)
    .SetKind(ConnectorKind.RightAngle).SetArrowEnd(EndArrow.Triangle);
page.AddConnector("decision", "end", fromSide: VisioSide.Right, toSide: VisioSide.Left)
    .SetKind(ConnectorKind.RightAngle).SetArrowEnd(EndArrow.Triangle).SetLabel("Yes");
document.Save();
```
<!-- /powerforge:code -->

Positions and sizes are `OfficeLength` values, and connectors stay glued to the shapes they join.

## What it does

- **Build diagrams** with pages, shapes, connectors, text, styles, Shape Data, layers, hyperlinks, containers and comments.
- **Diagram builders** for flowcharts, block, dependency, architecture, network, swimlane, org chart, sequence and timeline diagrams.
- **Edit loaded diagrams**: select shapes, query topology, replace stencils, maintain containers and swimlanes, and keep ShapeSheet sections and formulas.
- **Lay out** with resize-to-content and topology-aware relayout for dense imported diagrams.
- **Import legacy files**: binary `.vsd`, `.vss` and `.vst` and Visio XML drawings load into the same model.
- **Render** pages to PNG, JPEG, TIFF, SVG and WebP, and validate generated and loaded diagrams.
- **Before sharing**: inspect and remove hidden data and provenance.
- **Macro-enabled packages** keep their VBA project without running or rewriting it.

The [Visio guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.visio-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| PDF | [OfficeIMO.Visio.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Visio.Pdf/README.md) |

## Limits

`.vsdx` and its template and stencil forms are native. Legacy binary and XML drawings import into the model within a bounded profile, and the renderer draws the shapes and connectors the model understands. The [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md) lists what is supported, preserved or rejected.

## Learn more

- [OfficeIMO.Visio on officeimo.com](https://officeimo.com/products/visio/) and the [Visio docs](https://officeimo.com/docs/visio/).
- [Visio guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.visio-guide.md) and [examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Visio).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD). Command line: `officeimo convert flow.vsdx flow.pdf` with [OfficeIMO.Tool](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Tool/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` and `System.IO.Packaging`.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 1 |
| Edit | 1 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 3 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 0 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |
| Convert | 0 | 3 | 0 | 0 | 0 | 0 |
| Export | 5 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Visio` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
