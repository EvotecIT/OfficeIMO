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

## Quick sample (fluent)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

var vsd = VisioDocument.Create("diagram.vsdx");
vsd.AsFluent()
   .Info(i => i.Title("Demo").Author("You"))
   .Page("Page-1", p => p
       .Rect("S1", 1, 1, 2, 1, "Start")
       .Diamond("D1", 4, 1.5, 2, 2, "Decision")
       .Ellipse("E1", 7, 1.5, 2, 1, "End")
       .Connect("S1", "D1", c => c.RightAngle().ArrowEnd(EndArrow.Triangle))
       .Connect("D1", "E1", c => c.RightAngle().ArrowEnd(EndArrow.Triangle).Label("Yes")))
   .End();
vsd.Save();
```

See `OfficeIMO.Examples/Visio/*` for more.

## Feature Scope (early)

- 📄 Pages: ✅ add/remove pages
- 🧱 Shapes: ✅ basic shapes from masters (rectangle, etc.), ✅ set text
- 🔗 Connectors: ✅ basic connectors between shapes
- 🎨 Themes: ⚠️ minimal/default theme usage

This package is intentionally minimal at this stage and will expand over time.

## At a glance

- Create/Load/Save .vsdx (OPC packaging)
- Add simple pages, shapes, and connectors
- Fluent builder: `Page(...)`, `Rect(...)`, `Square(...)`, `Ellipse(...)`, `Circle(...)`, `Diamond(...)`, `Triangle(...)`, `Connect(...)`

## Why OfficeIMO.Visio (early)

- Minimal, no‑frills VSDX generation and reading using OPC + LINQ to XML
- Practical starting point for simple diagrams (pages, basic shapes, connectors)
- Designed to evolve as core scenarios are validated
