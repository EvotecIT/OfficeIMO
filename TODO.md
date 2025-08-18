# TODO — OfficeIMO: Word / Excel / PowerPoint (normal + fluent APIs)

> Design rule: **one .csproj per product**. Fluent lives inside the same project under the `OfficeIMO.{Product}.Fluent` namespace.
> Optional: a tiny `OfficeIMO.Core` can be added later; this file notes what could move there.

---

## Goals

- Add **OfficeIMO.PowerPoint** with normal + fluent APIs.
- Add **fluent APIs** for **Word** and **Excel** (no breaking changes).
- Keep read/write parity and similar ergonomics across W/E/P.
- Reuse shared helpers; keep tests and docs consistent.

---

## Packaging & Structure

- Projects
  - `OfficeIMO.Word` *(existing)* — add `OfficeIMO.Word.Fluent` namespace.
  - `OfficeIMO.Excel` *(existing)* — add `OfficeIMO.Excel.Fluent` namespace.
  - `OfficeIMO.PowerPoint` *(new)* — includes `OfficeIMO.PowerPoint.Fluent`.
  - `OfficeIMO.Tests`, `OfficeIMO.VerifyTests`, `OfficeIMO.Examples` *(extend)*.
  - Optional later: `OfficeIMO.Core` (see **What could move to Core** below).

- Namespaces
  - Normal: `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`
  - Fluent: `OfficeIMO.Word.Fluent`, `OfficeIMO.Excel.Fluent`, `OfficeIMO.PowerPoint.Fluent`

- Targets
  - Align with existing matrix (`netstandard2.0`, `net472`, modern .NET).

---

## What could move to `OfficeIMO.Core` (optional)

- **Units/Math**: twips↔pt, emu↔pt↔px, Excel column width↔pixels, basic geometry.
- **Colors/Themes**: ARGB helpers, theme palette, tint/shade.
- **OPC/OpenXML helpers**: part creation, rel management, content-type overrides.
- **Feature SPIs**:
  - `ITextMeasurer` (used by Excel AutoFit), default heuristic + optional ImageSharp adapter.
  - `IImageProcessor` (resize/re-encode for Word/PowerPoint), optional.
- **Diagnostics**: shared exceptions, guards, validation helpers.
- **File IO**: async load/save helpers; `IAsyncDisposable` pattern; sync wrappers.

> If Core isn’t created now, keep these blocks behind internal static helpers in each product so extraction is trivial later.

---

## Milestones (vertical slices)

### M1 — Scaffolding & Baseline
- [ ] Create `OfficeIMO.PowerPoint` project.
- [ ] Add fluent namespaces to Word/Excel projects (`*.Fluent`).
- [ ] Extend examples and tests projects for W/E/P.

### M2 — Fluent for Word (additive)
- [ ] `WordFluentDocument` + `WordDocument.AsFluent()`.
- [ ] Builders: `Info`, `Section`, `Page`, `Paragraph`, `Run`, `List`, `Table`, `Image`, `Headers`, `Footers`.
- [ ] Read helpers: `.ForEachParagraph(...)`, `.Find(...)`.
- [ ] Verify snapshots for fluent samples.

### M3 — Excel: normalize + fluent
- [ ] Normal surface audit (`Workbook`, `Worksheet`, `Cell`, `Range`, `Table`, `Style`).
- [ ] `ExcelFluentWorkbook` + `.AsFluent()`; chain sheets/rows/columns/ranges/tables/styling.
- [ ] `AutoFitColumns/AutoFitRows` with options (heuristic by default; precise if measurer available).

### M4 — PowerPoint: core write
- [ ] Normal API: create/open/save `.pptx`, add slides, text boxes, pictures, tables.
- [ ] Fluent API: `AsFluent().Slide(s => ...)`.
- [ ] Slide layouts + masters (basic), slide notes (read/write), basic theme.

### M5 — PowerPoint read + refinements
- [ ] Read model: enumerate slides/shapes/text/images/tables.
- [ ] Simple transitions; per-slide background; per-shape formatting.
- [ ] Charts MVP via embedded Excel part (write first, read later).

### M6 — Docs & stabilization
- [ ] README and samples showing normal vs fluent for W/E/P.
- [ ] API docs (XML comments).
- [ ] Verify tests for generated parts where appropriate.

---

## API Proposals (Normal + Fluent)

> Fluent is a thin, chainable layer; normal API remains the authoritative, complete surface.

### Word — Normal

```csharp
using OfficeIMO.Word;

using (var doc = WordDocument.Create("demo.docx")) {
    doc.Title = "Report";
    var section = doc.AddSection();
    section.PageOrientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Portrait;

    var p = doc.AddParagraph("Hello Word from OfficeIMO"); p.Bold = true;

    var list = doc.AddList(WordListStyle.Bulleted);
    list.AddItem("Item 1"); list.AddItem("Item 2");

    doc.AddHeadersAndFooters();
    doc.Header.Default.AddParagraph("Header text");
    doc.Save();
}
````

**Read**

```csharp
using var doc = WordDocument.Load("demo.docx");
foreach (var s in doc.Sections)
    foreach (var para in s.Paragraphs)
        Console.WriteLine(para.Text);
```

### Word — Fluent

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

using var doc = WordDocument.Create("fluent.docx")
    .AsFluent()
    .Info(i => i.Title("Report").Author("OfficeIMO"))
    .Section(s => s
        .Page(p => p.Portrait().A4())
        .Paragraph(p => p.Text("Hello Word").Bold())
        .List(l => l.Bulleted().Item("Item 1").Item("Item 2")))
    .Headers(h => h.Default(hd => hd.Paragraph("Header text")))
    .End();

doc.Save();
```

---

### Excel — Normal

```csharp
using OfficeIMO.Excel;

using (var x = ExcelDocument.Create("data.xlsx")) {
    var sheet = x.AddWorksheet("Data");
    sheet.SetCell("A1", "Name"); sheet.SetCell("B1", "Score");
    sheet.SetCell("A2", "Alice"); sheet.SetCell("B2", 93);
    sheet.SetCell("A3", "Bob");   sheet.SetCell("B3", 88);
    sheet.AddTable("A1:B3", hasHeader: true, name: "Scores");

    sheet.AutoFitColumns(); // heuristic; precise if measurer is available
    sheet.AutoFitRows();
    x.Save();
}
```

**Read**

```csharp
using var x = ExcelDocument.Load("data.xlsx");
var sheet = x.Worksheets["Data"];
foreach (var row in sheet.UsedRange().Rows)
    Console.WriteLine($"{row["A"]} -> {row["B"]}");
```

### Excel — Fluent

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

using var book = ExcelDocument.Create("fluent.xlsx")
    .AsFluent()
    .Sheet("Data", s => s
        .HeaderRow("Name", "Score")
        .Row("Alice", 93)
        .Row("Bob", 88)
        .Table("Scores")
        .Columns(c => c.Col(1, w => w.Width(22)))
        .AutoFit(precise: true))
    .End();

book.Save();
```

---

### PowerPoint — Normal

```csharp
using OfficeIMO.PowerPoint;

using (var ppt = PowerPointPresentation.Create("deck.pptx")) {
    ppt.Title = "Quarterly Review"; ppt.Author = "OfficeIMO";

    var slide = ppt.AddSlide(layout: BuiltInSlideLayout.TitleAndContent);
    slide.AddTitle("Q3 Results");
    slide.AddTextBox("Highlights:", left: 50, top: 150, width: 600, height: 40).Bold();
    slide.AddBullets(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, left: 70, top: 200);

    ppt.Save();
}
```

**Read**

```csharp
using var ppt = PowerPointPresentation.Load("deck.pptx");
foreach (var slide in ppt.Slides) {
    Console.WriteLine($"Slide #{slide.Number} — {slide.Title?.Text}");
    foreach (var shape in slide.Shapes)
        Console.WriteLine($"  Shape: {shape.Kind} Text: {shape.Text}");
}
```

### PowerPoint — Fluent

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;

using var deck = PowerPointPresentation.Create("fluent-deck.pptx")
    .AsFluent()
    .Info(i => i.Title("Quarterly Review").Author("OfficeIMO"))
    .Slide(s => s
        .Layout(BuiltInSlideLayout.TitleAndContent)
        .Title("Q3 Results")
        .Text(t => t.Box("Highlights:", 50, 150, 600, 40).Bold())
        .Bullets(b => b.Items(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, 70, 200)))
    .End();

deck.Save();
```

---

## Detailed Work Items

### A. Core / Infrastructure (in each project now; movable to Core later)

* Packaging & Parts helpers (create parts, manage rels, content types).
* Units & Geometry (twips/emu/px/cm; alignment).
* Media (image load/resize/re-encode; de-dupe; rel management).
* Themes & Styles (baseline palette & fonts; simple style intent→OOXML mapping).
* Errors & Validation (friendly exceptions annotated with location).
* API Consistency (method names, options).

### B. Word — Fluent Layer (non-breaking)

* `WordFluentDocument` + `.AsFluent()`.
* Builders: `Info`, `Section`, `Page`, `Paragraph`, `Run`, `List`, `Table`, `Image`, `Headers`, `Footers`.
* Read helpers: `.ForEachParagraph`, `.Find`, `.Select`.
* Thin shorthands for existing converters (`.FromHtml(...)`, `.ToPdf(...)`) if desired.
* Verify snapshots for common docs.

### C. Excel — Normalize + Fluent

* Normal audit: `Workbook`, `Worksheet`, `Cell`, `Range`, `Table`, `Style`.
* Fluent: `ExcelFluentWorkbook` + chaining for sheets/rows/columns/ranges/tables.
* Reading helpers: `.UsedRange()`, `.Rows`, `.Columns`, `.Cells(address)`, `.Table(name)`.
* Writing helpers: `Row(...)`, `Column(...)`, `Range("A1:B10").Values(...)`.
* Formulas, number/date formats, freeze panes, AutoFit.
* Named ranges, table creation with header inference.
* Verify snapshots for samples.

### D. PowerPoint — Normal + Fluent

* Model: `PowerPointPresentation`, `PowerPointSlide`, `PPShape`, `PPTextBox`, `PPPicture`, `PPTable`, `PPNotes`.
* Write: create pptx; add slide (with layout), text boxes, bullets, pictures, tables.
* Read: enumerate slides/shapes/text/images/tables; read notes.
* Layout/Master: pick built-in layout; bind placeholders; minimal theme support.
* Transitions: simple (Fade, Wipe).
* Charts: MVP embed (generate embedded xlsx; bind chart).
* Fluent wrapper: `PowerPointFluentPresentation` + `AsFluent()`; builders for `.Slide()`, `.Title()`, `.Text()`, `.Bullets()`, `.Image()`, `.Table()`, `.Notes()`.
* Verify snapshots on key parts; golden files opened in Office.

---

## Acceptance Criteria

* **Word Fluent**: README scenarios reproducible with fluent and produce equivalent output.
* **Excel Fluent**: Create/Load → write/read values → add table → format → AutoFit; parity examples normal vs fluent.
* **PowerPoint Core**: Create deck; at least Title and Title+Content layouts; title, text box, bullets, picture; Save/Load round-trip.
* **PowerPoint Read/Advanced**: Enumerate shapes/text; notes read/write; simple transitions; chart MVP (embedded xlsx).