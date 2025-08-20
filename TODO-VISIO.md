# TODO — OfficeIMO.Visio (Normal + Fluent APIs)

> `.vsdx` is an OPC (ZIP+XML) format. There is no official OpenXML SDK model for Visio.
> OfficeIMO.Visio will expose two complementary layers:
>
> - **Standard API** — explicit, object-oriented (`VisioDocument`, `VisioPage`, `VisioShape`, `VisioConnector`).
> - **Fluent API** — chainable DSL (`VisioFluentDocument`, `.Page(...)`, `.Shape(...)`, `.Connector(...)`).
>
> Both live in the same project; namespaces:
> - `OfficeIMO.Visio` — standard API
> - `OfficeIMO.Visio.Fluent` — fluent API

---

## Goals

- Provide **ergonomic authoring** for common scenarios (flowcharts, diagrams).
- Hide **OPC/ShapeSheet internals** behind builders and helpers.
- Ensure **future extensibility**: AWS/Azure/Cisco stencils, validation, auto-layout.
- Offer **testability**: snapshot & validation APIs for CI/CD.

---

## Standard API (Normal)

### Example usage

```csharp
using var doc = VisioDocument.Create("diagram.vsdx");

var page = doc.AddPage("Process", widthInches: 11, heightInches: 8.5);

// Shapes
var start = page.AddShape("Start",
    master: FlowchartMasters.Terminator,
    x: 1.0, y: 6.5, w: 2.2, h: 0.9,
    text: "Start");

var task1 = page.AddShape("Task1",
    master: FlowchartMasters.Process,
    x: 1.0, y: 5.0, w: 2.5, h: 1.0,
    text: "Validate");
task1.Data["Owner"] = "Ops";

// Connector
var conn = page.AddConnector("PathYes",
    fromShape: task1,
    toShape: start,
    kind: ConnectorKind.RightAngle);
conn.Label = "Yes";

doc.Save();
````

### Key classes and methods

* **VisioDocument**

  * `static Create(string path)`, `static Load(string path)`
  * `AddPage(string name, double widthInches, double heightInches)`
  * `Pages { get; }`
  * `Save()`
  * `Validate(...)`, `Snapshot()`, `RequestRecalcOnOpen()`

* **VisioPage**

  * `AddShape(id, master, x, y, w, h, text?)`
  * `AddConnector(id, fromShape, toShape, ConnectorKind)`
  * `Size(w,h)`, `Grid(visible, snap)`
  * `Shapes`, `Connectors`

* **VisioShape**

  * `Id`, `NameU`, `Master`
  * `PinX`, `PinY`, `Width`, `Height`
  * `Text`
  * `Data[string key]` (shape data properties)

* **VisioConnector**

  * `From`, `To`
  * `Kind` (Straight, RightAngle, Curved, Dynamic)
  * `Arrow(EndArrow style)`
  * `Label(string)`

---

## Fluent API

### Example usage

```csharp
VisioDocument.Create("fluent-diagram.vsdx")
  .AsFluent()
  .Info(i => i.Title("Order Flow").Author("OfficeIMO"))

  .Stencil(st => st.Use(BuiltInStencil.Flowchart))

  .Page("Process", p => p
      .Size(11, 8.5).Grid(visible: true, snap: true)

      .Shape("Start", s => s.Flowchart(FlowchartMasters.Terminator)
                            .At(1.0, 6.5).Size(2.2, 0.9).Text("Start"))

      .Shape("Task1", s => s.Flowchart(FlowchartMasters.Process)
                            .At(1.0, 5.0).Size(2.5, 1.0).Text("Validate")
                            .Data("Owner","Ops"))

      .Connector("PathYes", c => c.From("Task1").To("Start").RightAngle().Label("Yes"))

      .Shape("End", s => s.Flowchart(FlowchartMasters.Terminator)
                          .At(1.0, 2.0).Size(2.2, 0.9).Text("End"))

      .AlignHorizontal("Start","Task1","End")
      .DistributeVertical("Start","Task1","End")
      .Theme(BuiltInVisioTheme.Office))

  .End()
  .Save();
```

### Fluent surface

* **VisioFluentDocument**

  * `.Info(Action<VisioInfoBuilder>)`
  * `.Stencil(Action<StencilBuilder>)`
  * `.Page(string name, Action<PageBuilder>)`
  * `.End()`

* **StencilBuilder**

  * `.Use(BuiltInStencil stencil)`
  * `.UseFile(string vssxPath)` (load external stencils, e.g. AWS, Cisco)

* **PageBuilder**

  * `.Size(widthIn, heightIn)`
  * `.Grid(visible, snap)`
  * `.Background(colorHexOrImage)`
  * `.Theme(BuiltInVisioTheme)`
  * `.Shape(id, Action<ShapeBuilder>)`
  * `.Connector(id, Action<ConnectorBuilder>)`
  * `.AlignHorizontal(params string[] ids)`, `.AlignVertical(...)`, `.DistributeHorizontal(...)`, `.DistributeVertical(...)`

* **ShapeBuilder**

  * `.Flowchart(FlowchartMasters m)` / `.Master(BasicMasters m)` / `.Master(string customName)`
  * `.At(x,y)`, `.Size(w,h)`, `.Text(string)`
  * `.Data(string key, string value)`
  * `.Fill(string hex)`, `.Line(string hex, double pt)`
  * `.Rotate(double degrees)`, `.ZOrder(int)`

* **ConnectorBuilder**

  * `.From(string id)`, `.To(string id)`
  * `.Straight()`, `.RightAngle()`, `.Curved()`
  * `.Arrow(EndArrow style)`, `.Label(string)`
  * `.Reroute()`

---

## Helpers & Improvements

### 1. Units & grids

* `Units.Inches(double)`, `Units.Cm(double)`, `Units.Pt(double)` → standardize coordinates.
* `Grid.Snap((x,y), stepInches: 0.125)` → auto snap positions.

### 2. Naming & IDs

* `Names.Safe("Decision 1")` → `"Decision1"`
* `Names.EnsureUnique(page, id)` → avoid duplicates.

### 3. Master registry

* `Masters.Resolve(FlowchartMasters.Process)` → returns master definition.
* `Masters.Resolve(BuiltInStencil.BasicShapes, "Rectangle")`.

### 4. Shape presets

* `ShapePresets.Flow.Process("Validate", w: 2.5, h: 1.0)`.
* `ShapePresets.Flow.Decision("OK?", w: 2.0, h: 1.5)`.

### 5. Connector presets

* `ConnectorPresets.RightAngle().Arrow(EndArrow.Triangle)`.
* `ConnectorPresets.Curved().Dashed()`.

### 6. Validation & testing

* `doc.Validate(v => v.RequireUniqueShapeIds().RequireAllConnectorsResolved().WarnOnOffGrid(0.1))`.
* `doc.Snapshot("Process")` → POCO for test assertions.

---

## Future Features

* **Stencils**

  * Built-in: Flowchart, UML, BPMN, Network, BasicShapes.
  * Vendor: AWS, Azure, Cisco, GCP (`.UseFile("AWS.vssx")`).
  * Treat as `Masters`.

* **Containers & Callouts**

  * `.Container("Region1", c => c.Style(ContainerStyle.AwsRegion).Include("VPC","EC2"))`
  * `.Callout("Note1", cl => cl.Text("Critical!").Target("EC2"))`

* **Themes & Layout**

  * `.Theme(BuiltInVisioTheme.Modern)`
  * `.AlignHorizontal(...)`, `.DistributeVertical(...)`
  * `.AutoLayout(direction: TopToBottom)`

* **Connector types & arrows**

  * Routing: Straight, RightAngle, Curved, Dynamic.
  * Arrowheads: None, Triangle, Stealth, Diamond, Oval, OpenArrow (\~15 built-in).

* **Validation rules**

  * BPMN, UML, network diagrams with constraints.
  * `doc.ValidateWithTemplate("BPMN")`.

---

## Quick Comparison

| Concept        | Standard API                                                             | Fluent API                                                                    |
| -------------- | ------------------------------------------------------------------------ | ----------------------------------------------------------------------------- |
| Add page       | `doc.AddPage("Process", 11, 8.5)`                                        | `.Page("Process", p => p.Size(11,8.5))`                                       |
| Add shape      | `page.AddShape("Task1", FlowchartMasters.Process, 1,5,2.5,1,"Validate")` | `.Shape("Task1", s => s.Flowchart(...).At(1,5).Size(2.5,1).Text("Validate"))` |
| Add connector  | `page.AddConnector("PathYes", from, to, ConnectorKind.RightAngle)`       | `.Connector("PathYes", c => c.From("Task1").To("Decision1").RightAngle())`    |
| Align          | `page.AlignHorizontal(new[]{"A","B","C"})`                               | `.AlignHorizontal("A","B","C")`                                               |
| Document props | `doc.Title="...", doc.Author="..."`                                      | `.Info(i => i.Title("...").Author("..."))`                                    |

---

## Acceptance Criteria

* **Standard API**: explicit OO model usable without fluent.
* **Fluent API**: ergonomic, consistent with Word/Excel/PowerPoint builders.
* **Helpers**: units, naming, masters, presets, validation, snapshot available for both.
* **Future ready**: stencils (AWS/Azure/etc.), containers/callouts, themes, auto-layout, validation rules.

