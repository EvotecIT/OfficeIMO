# TODO — OfficeIMO.Visio (.vsdx), normal + fluent (same project)

> `.vsdx` is an OPC (ZIP+XML) package. There is no Open XML SDK object model for Visio.
> We will author/read parts directly (System.IO.Packaging + LINQ to XML).
> Fluent lives in `OfficeIMO.Visio.Fluent` within the same .csproj.

---

## Goals

- Create, read, and update `.vsdx` without Visio installed (no COM).
- MVP: pages + basic shapes (rectangle/ellipse/text), simple connectors.
- Phase 2: masters (stencils), images, recalculation hints.
- Provide a concise fluent DSL for common diagrams.

---

## Package & Parts (write path)

- `[Content_Types].xml` with overrides for:
  - `/visio/document.xml` (document main),
  - `/visio/pages/pages.xml`,
  - `/visio/pages/page1.xml`,
  - add `/visio/masters/*` only when used,
  - images under `/media/*` when used.
- Relationships:
  - package → `/visio/document.xml` (document rel),
  - document → `/visio/pages/pages.xml` (pages rel),
  - pages → `/visio/pages/pageN.xml` (page rel),
  - page/master → `/media/*` images (image rels when used).

---

## API (Normal)

```csharp
using OfficeIMO.Visio;

using var vsd = VisioDocument.Create("diagram.vsdx");
var page = vsd.AddPage("Page-1", widthInches: 8.5, heightInches: 11);

// Add a rectangle centered at (4, 5.5), size 2x1 inches
var rect = page.AddRectangle(centerX: 4, centerY: 5.5, width: 2, height: 1, text: "Hello Visio");

// Simple connector
var rect2 = page.AddRectangle(7, 5.5, 2, 1, "Next");
page.Connect(rect, rect2, ConnectorKind.Dynamic);

vsd.Save();
````

**Read**

```csharp
using var vsd = VisioDocument.Load("existing.vsdx");
foreach (var p in vsd.Pages) {
    Console.WriteLine($"{p.Name} {p.WidthInches}x{p.HeightInches}");
    foreach (var s in p.Shapes)
        Console.WriteLine($"  {s.Id}: {s.NameU} [{s.PinX},{s.PinY}] {s.Text}");
}
```

* Core model

  * `VisioDocument`: `Pages`, `AddPage(name, widthInches, heightInches)`, `Save/Load`.
  * `VisioPage`: `Name`, `WidthInches`, `HeightInches`, `AddRectangle(...)`, `AddEllipse(...)`, `Connect(from,to,kind)`, `Shapes`, `Connectors`.
  * `VisioShape`: `Id`, `NameU`, `Text`, basic ShapeSheet cells: `PinX`, `PinY`, `Width`, `Height`, `Angle`.
  * `VisioConnector`: `From`, `To`, `Kind`.

---

## Fluent DSL (same project, `OfficeIMO.Visio.Fluent`)

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;

VisioDocument.Create("net.vsdx")
  .AsFluent()
  .Page("Network", p => p
      .Rect("Web", center: (2, 4), size: (2, 1)).Out(out var web)
      .Rect("API", center: (6, 4), size: (2, 1)).Out(out var api)
      .Rect("DB",  center: (10,4), size: (2, 1)).Out(out var db)
      .Connect(web, api)
      .Connect(api, db))
  .End()
  .Save();
```

* Fluent surface

  * `.Page(name, builder)`
  * Shape builders: `.Rect(text, center, size)`, `.Ellipse(text, center, size)`, `.Text(text, at, size?)`
  * `.Connect(from, to, ConnectorKind.Dynamic|Straight|Curved)`
  * `.Style(...)` (later), `.WithData(key,value)` (later)
  * `.RequestRecalcOnOpen()` (optional flag on document)

---

## Implementation Plan

### M1 — OPC scaffolding

* [ ] Create package + `[Content_Types].xml`.
* [ ] Add package→document, document→pages, pages→page1 relationships by **type**.
* [ ] Write `/visio/document.xml` skeleton.
* [ ] Write `/visio/pages/pages.xml` with `<Page>` entry and `Rel/@r:id → page1.xml`.
* [ ] Write `/visio/pages/page1.xml`:

  * Root `PageContents` (2012/main),
  * Single `Shape` rectangle with basic Geometry (`PinX`, `PinY`, `Width`, `Height`) and `Text`.

### M2 — Reader

* [ ] Open existing `.vsdx`, resolve relationships by **type** (no hardcoded paths).
* [ ] Parse `pages.xml` → page list; read each `pageN.xml` → shapes (ID, NameU, Text, key cells).

### M3 — Connectors (basic)

* [ ] Author orthogonal connector between two shapes; minimal reroute support.
* [ ] Add helper `.RequestRecalcOnOpen()` to get Visio to reroute on open when topology changes.

### M4 — Masters & Images (optional)

* [ ] `/visio/masters/masters.xml` + `/visio/masters/masterN.xml` (created only if used).
* [ ] Allow shapes to reference a `Master`; drop multiple instances.
* [ ] Embed images in `/media/*` and reference from page/master.

### M5 — Fluent polish + tests

* [ ] DSL sugar (e.g., `.RoundedRect(...)`, `.Align(...)`).
* [ ] Verify snapshots of key parts, reading/writing round-trip tests.
* [ ] “Opens without repair” smoke test file artifacts.

---

## Acceptance Criteria

* New `.vsdx` with 1 page and a rectangle opens cleanly in Visio.
* Read existing `.vsdx`: list pages and basic shapes with positions and text.
* Connectors render between shapes after open (recalc flag where needed).
* Masters/images are optional and only emitted when used.
