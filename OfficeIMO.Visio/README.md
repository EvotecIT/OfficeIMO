# OfficeIMO.Visio — .NET Visio Utilities

OfficeIMO.Visio provides helpers for creating and editing .vsdx drawings with Open XML.

- Targets: netstandard2.0, net472, net8.0, net9.0
- License: TBD (not MIT yet)
- NuGet: `OfficeIMO.Visio`
- Dependencies: SixLabors.ImageSharp, System.IO.Packaging (Windows), Microsoft.Bcl.AsyncInterfaces (net472)

## Install

```powershell
dotnet add package OfficeIMO.Visio
```

## Quick sample

```csharp
using OfficeIMO.Visio;

using var vsd = new VisioDocument();
var page = vsd.AddPage("Diagram");
var rect = page.AddShape("Rect", VisioMaster.Rectangle, x: 1, y: 1, width: 3, height: 2);
rect.Text = "Hello Visio";
vsd.SaveAs("diagram.vsdx");
```

See `OfficeIMO.Examples/Visio/*` for more.

## Feature Scope (early)

- Pages: add/remove pages
- Shapes: add basic shapes from masters (rectangle, etc.), set text
- Connectors: basic connectors between shapes
- Themes: minimal/default theme usage

This package is intentionally minimal at this stage and will expand over time.
