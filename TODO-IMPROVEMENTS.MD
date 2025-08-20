## Cross‑cutting conventions (apply to all products)

**Entry/exit**

* `AsFluent()` to enter fluent mode.
* `.End()` returns the underlying document/workbook/deck/drawing.

**Style**

* Prefer property‑like verbs over `SetXxx(...)`.
  Use `Title("...")`, `Margins(...)`, `Size(...)`, `Orientation(...)`, `Align(...)`, `Wrap(...)`, etc.
* Keep **1‑based indexing** for user‑facing row/column APIs in Excel & tables.
* Accept both `"rrggbb"` and `"#rrggbb"` for colors; normalize internally.
* Alignment enums you own (map internally):

  ```csharp
  public enum HorizontalAlignment { Left, Center, Right }
  public enum VerticalAlignment   { Top,  Middle, Bottom }
  ```

**Fluent block structure**

* Document/workbook/deck/drawing owns **blocks**:

  * Word: `.Paragraph(...)`, `.List(...)`, `.Table(...)`, `.Image(...)`, `.Section(...)`, `.Headers(...)`, `.Footers(...)`
  * Excel: `.Sheet(...)`
  * PowerPoint: `.Slide(...)`
  * Visio: `.Page(...)` and `.Stencil(...)`
* Inside a block, chain concise, domain terms.

**Save**

* Keep persistence on the non‑fluent object:
  `doc.AsFluent()...End().Save(...)`.

---

# WORD

> Current status: Word is the most advanced. Features include sections, paragraphs, lists, tables, images, hyperlinks, fields, headers/footers, etc.; repo also has examples/tests and converters projects. Fluent work is underway (as in your shared examples). ([GitHub][1])

### Final fluent shape (write)

```csharp
using var doc = WordDocument.Create("report.docx");

doc.AsFluent()
  .Info(i => i
      .Title("Quarterly Review")
      .Author("OfficeIMO")
      .Keywords("q3, revenue")
      .Custom("Reviewed", true))

  .PageSetup(ps => ps
      .Orientation(PageOrientationValues.Portrait)
      .Size(WordPageSize.A4)
      .Margins(WordMargin.Normal)
      .DifferentFirstPage()
      .DifferentOddAndEvenPages())

  .Section(sec => sec
      .New(SectionMarkValues.NextPage)
        .Margins(WordMargin.Narrow).Size(WordPageSize.Legal).PageNumbering(restart: true)
      .New(SectionMarkValues.NextPage)
        .Margins(WordMargin.Wide).Size(WordPageSize.A3).PageNumbering(restart: false))

  .Paragraph(p => p
      .Text("Hello ").Text("World", t => t.BoldOn().Color("#ff0000")).Text("!"))

  .List(l => l.Numbered().StartAt(3)
      .Item("First").Item("Second").Indent().Item("Second.Child"))

  .Table(t => t
      .Columns(3).PreferredWidth(percent: 100)
      .Header("Name", "Role", "Score")
      .Row("Alice", "Dev", 98)
      .Row("Bob",   "Ops", 91)
      .Style(WordTableStyle.TableGrid)
      .Align(HorizontalAlignment.Center))

  .Image(img => img
      .Add("Images/Chart.png").Size(480).Wrap(WrapTextImage.Square).Align(JustificationValues.Center))

  .End()
  .Save(false);
```

**Key builder surfaces**

* **Info**
  `Title(string)`, `Author(string)`, `Subject(string)`, `Keywords(string)`, `Comments(string)`, `Custom(string key, object value)`
* **PageSetup (defaults for first section)**
  `Orientation(...)`, `Size(WordPageSize)`, `Margins(WordMargin)`, `DifferentFirstPage()`, `DifferentOddAndEvenPages()`
* **Section**
  `New(SectionMarkValues breakType = NextPage)` → fluent override chain: `Margins(...)`, `Size(...)`, `Columns(...)`, `PageNumbering(restart: bool)`
* **Paragraph**
  `Text(string, Action<TextStyleDsl>? style = null)`, `Link(url, text?, style?)`, `Break()`, `Tab()`, `InlineImage(path/bytes/stream, widthPx?, heightPx?, alt?)`
* **List**
  `Bulleted(WordListStyle style = WordListStyle.Bulleted)`, `Numbered(WordListStyle style = WordListStyle.Numbered)`, `StartAt(int)`, `Item(string)`, `Indent()`, `Outdent()`, `Level(int)`
* **Table**
  `Columns(int)`, `Create(rows, cols)`, `From2D(object[,])`, `HeaderRow(int index)`, `Header(params object[])`, `Row(params object[])`,
  `InsertRow(int index, params object[])`, `InsertColumn(int index, params object[])`, `DeleteRow(int index)`, `DeleteColumn(int index)`,
  `Cell(int row, int col, Action<CellBuilder>)`, `ForEachCell(Action<int,int,CellBuilder>)`,
  `PreferredWidth(percent?, points?)`, `Style(WordTableStyle)`, `Align(HorizontalAlignment)`
* **Image (block‑level)**
  `Add(path|stream|bytes|url, width?, height?, wrap?, alignment?, description?)` then `Size(...)`, `MaxWidth(...)`, `Wrap(...)`, `Align(...)`, `BehindText(...)`, `Position(...)`, `Alt(title?, desc?)`, `Link(url)`

### Final fluent shape (read)

Simple, discoverable queries without leaking OOXML:

```csharp
doc.AsFluent()
  .ForEachSection((index, s) => { /* read margins, size, numbering */ })
  .ForEachParagraph((index, p) => Console.WriteLine(p.Text))
  .FindText("World", caseSensitive: false, wholeWord: false, hit => { /* location, section, paragraph index */ })
  .Lists(l => { foreach (var list in l) { /* items, level */ } })
  .Tables(t => { foreach (var table in t) { /* rows, cells, styles */ } })
  .Images(imgs => { foreach (var im in imgs) { /* size, wrap, alt */ } })
  .End();
```

> **Renames / deprecations (Word)**
>
> * `Page(...)` → **`PageSetup(...)`**
> * `SetMargins/SetPageSize/SetOrientation` → **`Margins/Size/Orientation`**
> * `AddListBulleted/AddListNumbered` → **`.List(l => l.Bulleted()/Numbered())`** with `.Item(...)`, `.StartAt(...)`
> * `AddImage(...)` → **`.Image(b => b.Add(...).Size(...).Wrap(...).Align(...))`**; keep `InlineImage` on the paragraph builder
> * Replace `Align(WordHorizontalAlignmentValues)` in fluent with **`Align(HorizontalAlignment)`** (map internally)

---

# EXCEL

> Current status (from README): Excel is described as **experimental**; the project folder exists in the repo. The Word area is mature; Excel APIs are being built out (you’ve discussed AutoFit, TableStyle enums). ([GitHub][1])

### Final normal shape

```csharp
using var book = ExcelDocument.Create("data.xlsx");
var sheet = book.AddWorksheet("Data");

sheet.SetCell("A1", "Name");  sheet.SetCell("B1", "Score");
sheet.SetCell("A2", "Alice"); sheet.SetCell("B2", 93);
sheet.SetCell("A3", "Bob");   sheet.SetCell("B3", 88);

sheet.AddTable("A1:B3", hasHeader: true, name: "Scores", style: TableStyle.TableStyleMedium9);
sheet.AutoFitColumns();                    // heuristic, cross-platform
sheet.AutoFitRows();                       // optional
book.Save();
```

**Reading (normal)**

```csharp
using var x = ExcelDocument.Load("data.xlsx");
var ws = x.Worksheets["Data"];
foreach (var row in ws.UsedRange().Rows)
    Console.WriteLine($"{row["A"]} -> {row["B"]}");
```

### Final fluent shape

```csharp
ExcelDocument.Create("fluent.xlsx")
  .AsFluent()
  .Sheet("Data", s => s
      .HeaderRow("Name", "Score")
      .Row("Alice", 93)
      .Row("Bob",   88)
      .Table("Scores", t => t.Style(TableStyle.TableStyleMedium9))
      .Columns(c => c.Col(1, w => w.Width(22)))
      .AutoFit(columns: true, rows: false))
  .Sheet("Summary", s => s
      .Cell("A1", "Average").Cell("B1", "=AVERAGE(Data!B2:B3)").NumberFormat("0.00"))
  .End()
  .Save();
```

**Fluent surface**

* **Workbook**: `.Sheet(name, Action<SheetBuilder>)`
* **Sheet**: `HeaderRow(params object[])`, `Row(params object[])`,
  `Cell("A1", value)`, `Cell(int row, int col, value)`, `Formula("B1", "=SUM(...)")`,
  `Table(name, Action<TableBuilder>? = null)` creates a table from the contiguous used range or explicit range,
  `Columns(Action<ColumnCollectionBuilder>)`, `Rows(...)`, `Range("A1:B10", Action<RangeBuilder>)`,
  `Freeze(topRows, leftCols)`, `AutoFit(bool columns = true, bool rows = true)`
* **Columns**: `Col(int index, Action<ColumnBuilder>)`
* **Column**: `Width(double)`; optional `Hidden(bool)`
* **Range**: `NumberFormat(string)`, `Style(string or built‑in)`, `Set(values2D)`, `Clear()`

> **Renames / deprecations (Excel)**
>
> * Prefer `Cell("A1", value)` (and/or indexer) to `SetCell` in fluent; keep `SetCell` in normal API if already shipped.
> * Unify **1‑based** row/column indexing for fluent `Cell(row, col, ...)`.
> * `AutoFitColumns/AutoFitRows` are fine on **normal** API; fluent uses `.AutoFit(...)` on the sheet builder.

*(Note on AutoFit: keep it **heuristic** and cross‑platform; measuring text exactly is not necessary for a useful default. Offer an opt‑in font‑metrics path behind a capability flag if you want, but default to an inexpensive approximation.)*

---

# POWERPOINT

> Current status: A `OfficeIMO.PowerPoint` project exists; README does not yet outline its features the way Word does. Propose usage parity with Word for consistency. ([GitHub][1])

### Final normal shape

```csharp
using var deck = PowerPointPresentation.Create("deck.pptx");
var slide = deck.AddSlide(layout: BuiltInSlideLayout.TitleAndContent);

slide.AddTitle("Q3 Results");
slide.AddTextBox("Highlights:", left: 50, top: 150, width: 600, height: 40).Bold();
slide.AddBullets(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, left: 70, top: 200);
slide.AddPicture("Images/chart.png", left: 400, top: 120, width: 300);

deck.Save();
```

### Final fluent shape

```csharp
PowerPointPresentation.Create("fluent-deck.pptx")
  .AsFluent()
  .Info(i => i.Title("Quarterly Review").Author("OfficeIMO"))
  .Slide(s => s
      .Layout(BuiltInSlideLayout.TitleAndContent)
      .Title("Q3 Results")
      .TextBox("Highlights:", x: 50, y: 150, w: 600, h: 40, style => style.Bold())
      .Bullets(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, x: 70, y: 200)
      .Image("Images/chart.png", x: 400, y: 120, w: 300)
      .Notes(n => n.Text("Spoken points...")))
  .End()
  .Save();
```

**Fluent surface**

* **Presentation**: `.Info(...)`, `.Slide(Action<SlideBuilder>)`
* **Slide**: `Layout(BuiltInSlideLayout)`, `Title(string)`,
  `TextBox(text, x, y, w, h, style?)`, `Bullets(IEnumerable<string>, x, y, style?)`,
  `Image(path/bytes/stream, x, y, w?, h?)`, `Table(Action<PptTableBuilder>)`,
  `Background(color|image)`, `Transition(PptTransition)` (simple enum for Fade/Wipe, etc.),
  `Notes(Action<PptNotesBuilder>)`
* **Read (minimal)**: `Slides.ForEach(s => { s.Title; s.TextBoxes; s.Images; s.Tables; s.Notes; })`

> **Renames / deprecations (PowerPoint)**
>
> * Keep names parallel to Word where concepts overlap (`Image`, `Table`, `Info`).
> * Use **`Layout(...)`** rather than `SetLayout`; `Title(...)` over `AddTitle(...)` in fluent.

---

# VISIO

> Current status: `OfficeIMO.Visio` project exists; usage is “working but not friendly”. VSDX is **Open Packaging Conventions + XML**, not part of OpenXML SDK, so you own the OPC parts/relationships. Structure your API around **pages, shapes, connectors**, with a consistent fluent DSL. ([GitHub][1], [Microsoft Learn][2])

### Final fluent shape (write)

```csharp
using var vsd = VisioDocument.Create("diagram.vsdx");

vsd.AsFluent()
  .Info(i => i.Title("Order Flow").Author("OfficeIMO"))
  .Stencil(st => st.Use(BuiltInStencil.BasicShapes).Use(BuiltInStencil.Flowchart))  // preload masters

  .Page("Process", p => p
      .Size(widthIn: 11, heightIn: 8.5)   // inches or cm; provide overloads
      .Grid(visible: true, snap: true)

      .Shape("Start", s => s
          .Master(FlowchartMasters.Terminator)
          .At(x: 1.0, y: 6.5).Size(w: 2.0, h: 0.8)
          .Text("Start"))

      .Shape("Task1", s => s
          .Master(FlowchartMasters.Process)
          .At(1.0, 5.0).Size(2.5, 1.0)
          .Text("Validate")
          .Data("Owner", "Ops"))

      .Shape("Decision1", s => s
          .Master(FlowchartMasters.Decision)
          .At(1.2, 3.5).Size(2.0, 1.5)
          .Text("OK?"))

      .Connector("PathYes", c => c
          .From("Task1").To("Decision1").RightAngle().Label("Yes"))

      .Connector("EndPath", c => c
          .From("Decision1").To("End").Curved().Arrow(EndArrow.Triangle))

      .Shape("End", s => s
          .Master(FlowchartMasters.Terminator)
          .At(1.0, 2.0).Size(2.0, 0.8)
          .Text("End"))

      .Theme(BuiltInVisioTheme.Office)
      .AlignHorizontal("Start", "Task1", "Decision1", "End")
      .DistributeVertical("Start", "Task1", "Decision1", "End"))

  .End()
  .Save();
```

**Fluent surface (Visio)**

* **Document**: `.Info(...)`, `.Stencil(Action<StencilBuilder>)`, `.Page(name, Action<PageBuilder>)`
* **StencilBuilder**: `Use(BuiltInStencil)` or `UseFile(path)` to load masters.
* **PageBuilder**:

  * `Size(widthIn, heightIn)`, `Landscape()`, `Portrait()`
  * `Grid(visible, snap)`, `Background(color|image)`
  * `Shape(string id, Action<ShapeBuilder>)`
  * `Connector(string id, Action<ConnectorBuilder>)`
  * `Theme(BuiltInVisioTheme)`
  * `AlignHorizontal(params string[] shapeIds)`, `AlignVertical(...)`, `DistributeHorizontal(...)`, `DistributeVertical(...)`
* **ShapeBuilder**:

  * `Master(BasicMasters|FlowchartMasters|UmlMasters|CustomMaster("Name"))`
  * `At(x, y)`, `Size(w, h)` (units: inches/cm; offer both)
  * `Text(string)`, `Data(string key, string value)`, `Rotate(deg)`, `ZOrder(int)`
  * `Fill(string hex)`, `Line(string hex, double pt = 1)`
* **ConnectorBuilder**:

  * `From(string shapeId)`, `To(string shapeId)`
  * `Straight()`, `RightAngle()`, `Curved()`
  * `Arrow(EndArrow style)`, `Label(string)`
  * `Reroute()` (simple auto routing)
* **Read (minimal)**:

  * `Pages["Process"].Shapes` → enumerate `Id`, `Master`, `Text`, `PinX/PinY`, `Width/Height`
  * `Pages["Process"].Connectors` → `From`, `To`, `Kind`

> **Renames / restructuring (Visio)**
>
> * Replace low‑level part/relationship driven calls with the fluent **Page/Shape/Connector** DSL above.
> * Hide master‑lookup boilerplate behind `Stencil(...).Use(...)` and a small set of **built‑in master enums** (`BasicShapes`, `Flowchart`).
> * Keep all fluent APIs inside the **same project** under `OfficeIMO.Visio.Fluent`.

---

## Quick “Old → New” rename map (Word & Tables especially)

| Old (examples seen)                                              | New                                                                 |
| ---------------------------------------------------------------- | ------------------------------------------------------------------- |
| `Page(...).SetPageSize(...).SetMargins(...).SetOrientation(...)` | `PageSetup(ps => ps.Size(...).Margins(...).Orientation(...))`       |
| `Section(s => { s.AddSection(...).Set... })`                     | `Section(sec => sec.New(...).Margins(...).Size(...))`               |
| `Paragraph(p => p.AddParagraph("...")...)`                       | `Paragraph(p => p.Text("...").Text("...", style => ...))`           |
| `AddListBulleted(..), AddListNumbered(..), AddListItem(..)`      | `.List(l => l.Bulleted().Item("...").Indent().Item("..."))`         |
| `AddImage(path, width, height, wrap, alignment, description)`    | `.Image(b => b.Add(path).Size(...).Wrap(...).Align(...).Alt(...))`  |
| `AddTable(2,3)` mixed with `.Columns(3)`                         | Pick one entry verb: `Columns(n)` or `Create(rows, cols)`; not both |
| `Align(WordHorizontalAlignmentValues.Center)` in fluent          | `Align(HorizontalAlignment.Center)`                                 |

---

## Why this fits the repo today

* Word already exposes the required primitives (paragraphs, tables, lists, images, sections, headers/footers, properties). The proposed fluent surface simply **wraps** those mature features without changing their semantics. ([GitHub][1])
* Excel is **experimental**; fluent can be designed without breaking anything, mirroring Word’s approach, and keeping normal APIs (`SetCell`, `AddWorksheet`, `AddTable`) intact. ([GitHub][1])
* PowerPoint/Visio projects exist; aligning their fluent style with Word keeps learning cost low and enforces **one mental model** across W/E/P/V. ([GitHub][1])
* Visio’s VSDX is **OPC + XML**; a Page/Shape/Connector DSL is the right abstraction and hides part management from consumers. ([Microsoft Learn][2])

---

## What to tell the code‑gen AI (per product)

**Word**

* Introduce/keep: `WordFluentDocument`, `.End()`, `.Info`, `.PageSetup`, `.Section`, `.Paragraph`, `.List`, `.Table`, `.Image`.
* Replace any `SetXxx` with property‑like verbs.
* Avoid `AddParagraph` inside the `Paragraph` builder; paragraphs are added by the document, not the paragraph DSL.
* Provide `InlineImage(...)` on the paragraph builder.

**Excel**

* Keep normal `SetCell`, `AddWorksheet`, `AddTable` as-is.
* Fluent: `.Sheet(name, s => ...)`, with `HeaderRow`, `Row`, `Cell`, `Range`, `Table`, `Columns`, `AutoFit`.
* Use 1‑based indexing consistently.

**PowerPoint**

* Fluent: `.Slide(s => s.Layout(...).Title(...).TextBox(...).Bullets(...).Image(...).Notes(...))`.
* Keep normal API for raw control.

**Visio**

* Fluent: `.Page("Name", p => p.Shape(...).Connector(...).Theme(...).Align...Distribute...)`.
* Add `Stencil(...)` to hide master loading; support a few **built‑in master enums** up front.
