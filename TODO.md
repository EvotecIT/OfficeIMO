Below is a ready‑to‑commit **`TODO.md`** you can drop into the root of **OfficeIMO**.
It’s condensed, task‑oriented, and proposes side‑by‑side **normal** and **fluent** APIs for **Word**, **Excel**, and **PowerPoint**, covering both **read** and **write** paths while preserving existing APIs.

---

# TODO — OfficeIMO: Word / Excel / PowerPoint (normal + fluent APIs)

> **Context (what exists today)**
>
> * Repository already ships **OfficeIMO.Word** with rich APIs and examples, plus an **experimental OfficeIMO.Excel** component; repo also contains `OfficeIMO.Word.Html`, `OfficeIMO.Word.Markdown`, `OfficeIMO.Word.Pdf`, unit tests and Verify snapshot tests. Target frameworks include `netstandard2.0`, `net472`, and modern .NET (8/9). ([GitHub][1])
> * Intro and usage background for Word are documented on the project blog. ([Evotec][2])

---

## Goals

* Add **OfficeIMO.PowerPoint** with both **normal** and **fluent** APIs.
* Introduce **fluent APIs** for **Word** and **Excel** alongside existing normal APIs (no breaking changes).
* Unify read/write capabilities and developer ergonomics across W/E/P.
* Reuse common infrastructure; keep packaging, CI, tests, docs consistent with the repo.

---

## Architecture & Packaging

* **Projects**

  * `OfficeIMO.Word` *(keep existing API intact)*.
  * `OfficeIMO.Excel` *(stabilize surface; add fluent)*.
  * `OfficeIMO.PowerPoint` *(new)*.
  * `OfficeIMO.Core` *(new, optional)* — shared helpers (parts, relationships, units, media, theming).
  * `OfficeIMO.Tests`, `OfficeIMO.VerifyTests` — extend coverage to Excel/PowerPoint.

* **Namespaces**

  * Normal:

    * `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`
  * Fluent (opt‑in, additive):

    * `OfficeIMO.Word.Fluent`, `OfficeIMO.Excel.Fluent`, `OfficeIMO.PowerPoint.Fluent`

* **Binary compatibility**

  * No breaking changes to public types and members in current `OfficeIMO.Word`.
  * Fluent API is additive: implemented via wrappers + extension methods.

* **Targets**

  * Align with repo matrix: Word & Excel already target `netstandard2.0`/`net472`/`net8.0`/`net9.0`; PowerPoint to follow same. ([GitHub][1])

---

## Milestones (deliver in thin vertical slices)

1. **M1 — Scaffolding & Baseline**

* [ ] Create `OfficeIMO.PowerPoint` project.
* [ ] Add `OfficeIMO.Core` (common: Packaging, Part helpers, Units, Color & Theme helpers, Image loader).
* [ ] Wire CI, code coverage, signing, packaging similar to existing projects. ([GitHub][1])
* [ ] Sample gallery project updates in `OfficeIMO.Examples`.

2. **M2 — Fluent API for Word (additive)**

* [ ] `WordFluentDocument` wrapper + `AsFluent()` extension on `WordDocument`.
* [ ] Fluent builders for `Document → Section → Paragraph → Run → Table/List → Header/Footer`.
* [ ] Read and Write parity for the most common flows (see API proposals below).
* [ ] Verify snapshots for fluent samples.

3. **M3 — Excel: normalize + fluent**

* [ ] Stabilize normal API surface for workbook/worksheet/cell/table.
* [ ] `ExcelFluentWorkbook` + `AsFluent()`; chainable ranges, tables, styling.
* [ ] Read: enumerate sheets, regions, tables, formulas. Write: set values, formats, tables.

4. **M4 — PowerPoint: core write**

* [ ] Normal API: create/open/save `.pptx`, add slides, text boxes, pictures, tables.
* [ ] Fluent API: `Presentation().AddSlide(s => …)` chain.
* [ ] Slide layouts + masters (basic), notes page (read/write text).
* [ ] Theme application (basic colors/fonts).

5. **M5 — PowerPoint read + advanced**

* [ ] Read model: enumerate slides/shapes/text/images/tables.
* [ ] Transitions (basic), per‑slide background, per‑shape formatting.
* [ ] Charts via embedded Excel part (write MVP, read later).

  * *Note:* Consider community libs (e.g., ShapeCrawler) for inspiration only. ([NuGet][3], [GitHub][4])

6. **M6 — Docs, samples, stabilization**

* [ ] Update README and **Docs** with new sections & samples.
* [ ] Breaking‑change check (should be none).
* [ ] API reference (XML docs) generation.

---

## API Proposals (Normal + Fluent)

> The “normal” APIs mirror current Word style (e.g., `WordDocument.Create`, `AddSection`, `Save`). Fluent is optional and layered via wrappers + extension methods.

### Word — Normal (existing style, both read & write)

**Create & write**

```csharp
using OfficeIMO.Word;

var path = "demo.docx";
using (var doc = WordDocument.Create(path)) {
    doc.Title = "Report";
    doc.Creator = "OfficeIMO";
    var section = doc.AddSection();
    section.PageOrientation = DocumentFormat.OpenXml.Wordprocessing.PageOrientationValues.Portrait;

    var p = doc.AddParagraph("Hello Word from OfficeIMO");
    p.Bold = true;

    var list = doc.AddList(WordListStyle.Bulleted);
    list.AddItem("Item 1");
    list.AddItem("Item 2");

    doc.AddHeadersAndFooters();
    doc.Header.Default.AddParagraph("Header text");
    doc.Save();
}
```

**Read**

```csharp
using OfficeIMO.Word;

using var doc = WordDocument.Load("demo.docx");
foreach (var s in doc.Sections) {
    foreach (var para in s.Paragraphs)
        Console.WriteLine(para.Text);
}
var props = doc.CustomDocumentProperties;
```

### Word — Fluent (additive)

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

var path = "fluent.docx";
using var doc = WordDocument.Create(path)
    .AsFluent()
    .Info(i => i.Title("Report").Author("OfficeIMO"))
    .Section(s => s
        .Page(p => p.Portrait().A4())
        .Paragraph(p => p.Text("Hello Word").Bold())
        .List(l => l.Bulleted()
            .Item("Item 1")
            .Item("Item 2")))
    .Headers(h => h.Default(hd => hd.Paragraph("Header text")))
    .End(); // returns underlying WordDocument

doc.Save();
```

---

### Excel — Normal (stabilize)

```csharp
using OfficeIMO.Excel;

var path = "data.xlsx";
using (var x = ExcelDocument.Create(path)) {
    var sheet = x.AddWorksheet("Data");
    sheet.SetCell("A1", "Name");
    sheet.SetCell("B1", "Score");
    sheet.SetCell("A2", "Alice");
    sheet.SetCell("B2", 93);
    sheet.SetCell("A3", "Bob");
    sheet.SetCell("B3", 88);
    sheet.AddTable("A1:B3", hasHeader: true, name: "Scores");
    x.Save();
}
```

**Read**

```csharp
using OfficeIMO.Excel;

using var x = ExcelDocument.Load("data.xlsx");
var sheet = x.Worksheets["Data"];
foreach (var row in sheet.UsedRange().Rows) {
    Console.WriteLine($"{row["A"]} -> {row["B"]}");
}
```

### Excel — Fluent (additive)

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;

var path = "fluent.xlsx";
using var book = ExcelDocument.Create(path)
    .AsFluent()
    .Sheet("Data", s => s
        .HeaderRow("Name", "Score")
        .Row("Alice", 93)
        .Row("Bob", 88)
        .Table("Scores")
        .Columns(c => c.Col(1, w => w.Width(22))))
    .End(); // underlying ExcelDocument

book.Save();
```

---

### PowerPoint — Normal (new)

```csharp
using OfficeIMO.PowerPoint;

var path = "deck.pptx";
using (var ppt = PowerPointPresentation.Create(path)) {
    ppt.Title = "Quarterly Review";
    ppt.Author = "OfficeIMO";

    var slide = ppt.AddSlide(layout: BuiltInSlideLayout.TitleAndContent);
    slide.AddTitle("Q3 Results");
    slide.AddTextBox("Highlights:", left: 50, top: 150, width: 600, height: 40).Bold();
    slide.AddBullets(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, left: 70, top: 200);

    ppt.Save();
}
```

**Read**

```csharp
using OfficeIMO.PowerPoint;

using var ppt = PowerPointPresentation.Load("deck.pptx");
foreach (var slide in ppt.Slides) {
    Console.WriteLine($"Slide #{slide.Number} — {slide.Title?.Text}");
    foreach (var shape in slide.Shapes)
        Console.WriteLine($"  Shape: {shape.Kind} Text: {shape.Text}");
}
```

### PowerPoint — Fluent (new)

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Fluent;

var path = "fluent-deck.pptx";
using var deck = PowerPointPresentation.Create(path)
    .AsFluent()
    .Info(i => i.Title("Quarterly Review").Author("OfficeIMO"))
    .Slide(s => s
        .Layout(BuiltInSlideLayout.TitleAndContent)
        .Title("Q3 Results")
        .Text(t => t.Box("Highlights:", 50, 150, 600, 40).Bold())
        .Bullets(b => b.Items(new [] { "Revenue +12%", "Churn -1.5%", "NPS 62" }, 70, 200)))
    .End(); // underlying PowerPointPresentation

deck.Save();
```

---

## Detailed Work Items

### A. Core / Infrastructure

* [ ] **Packaging & Parts**: shared helpers for creating parts, managing relationships, content types, and URIs (ppt/word/xl).
* [ ] **Units & Geometry**: twips/emu/px/cm conversions; alignment helpers.
* [ ] **Media**: image loader (PNG/JPG/SVG→PNG conversion decision), dedupe, rel management.
* [ ] **Themes & Styles (baseline)**: color palette, font schemes; map simple style intents to Open XML constructs.
* [ ] **Errors & Validation**: friendly exceptions with location (slide/shape/cell/paragraph).
* [ ] **API Consistency**: method names & options consistent across W/E/P.

### B. Word — Fluent Layer (non‑breaking)

* [ ] `WordFluentDocument` wrapper + `AsFluent()` extension on `WordDocument`.
* [ ] Builders: `Info`, `Section`, `Page`, `Paragraph`, `Run`, `List`, `Table`, `Image`, `Headers`, `Footers`.
* [ ] Read fluent: `.ForEachParagraph(...)`, `.Find(text/options)`, `.Select(selector)` helpers.
* [ ] Interop with existing modules: `Word.Html`, `Word.Markdown`, `Word.Pdf` via fluent `.FromHtml()/.ToPdf()` shorthands (thin wrappers). ([GitHub][1])
* [ ] Verify snapshots for common docs.

### C. Excel — Normalize + Fluent

* [ ] Normal API audit: `ExcelDocument`, `Worksheet`, `Cell`, `Range`, `Table`, `Style`.
* [ ] Fluent `ExcelFluentWorkbook` + `AsFluent()`; chain for sheets, rows, columns, ranges, tables, basic formatting.
* [ ] Reading helpers: `.UsedRange()`, `.Rows`, `.Columns`, `.Cells(address)`, `.Table(name)`.
* [ ] Writing helpers: bulk set (`Row(...)`, `Column(...)`, `Range("A1:B10").Values(...)`).
* [ ] Formulas (string entry), number/date formats, freeze panes, auto‑fit.
* [ ] Table creation (header inference), named ranges.
* [ ] Verify snapshots for sample workbooks.

### D. PowerPoint — Normal + Fluent

* [ ] **Model**: `PowerPointPresentation`, `PowerPointSlide`, `PPShape`, `PPTextBox`, `PPPicture`, `PPTable`, `PPNotes`.
* [ ] **Write**: create pptx; slide add (with layout), text boxes, bullets, pictures (media rels), tables.
* [ ] **Read**: enumerate slides; inspect shapes; extract text; list images/tables; read notes.
* [ ] **Layout/Master**: choose a built‑in layout; bind placeholders; minimal theme support.
* [ ] **Transitions**: per‑slide basic transitions (Fade, Wipe).
* [ ] **Charts**: MVP embed (generate minimal embedded Excel part; bind chart).
* [ ] Fluent wrapper `PowerPointFluentPresentation` + `AsFluent()`; builders for `.Slide()`, `.Title()`, `.Text()`, `.Bullets()`, `.Image()`, `.Table()`, `.Notes()`.
* [ ] Verify snapshots on XML parts + golden files opened in Office.

### E. Samples, Docs, CI

* [ ] Extend `OfficeIMO.Examples` with paired **normal vs fluent** examples (W/E/P).
* [ ] README: add PowerPoint section; add fluent intro; update platform matrix. ([GitHub][1])
* [ ] API Docs (XML comments → doc gen).
* [ ] CI pipelines for new projects; package & publish.

---

## Non‑Goals (initially)

* Advanced animations, motion paths, video/audio embedding (PowerPoint) — backlog.
* Complex Excel charting/pivot tables — backlog (focus on tables + basic charts via embedded parts first).
* Full HTML→PPT/Excel conversion.

---

## Design Notes

* **Fluent is opt‑in**: discoverable via `.AsFluent()`; returns a disposable wrapper exposing chainable methods and `.End()` to access underlying doc.
* **Normal API remains the authority** for completeness and low‑level access.
* **Testing**: continue using Verify snapshot tests (already present in repo) to lock desired XML output; add focused unit tests for builders. ([GitHub][1])
* **Inspiration**: Open XML SDK + ecosystem; for PPT ergonomics, review approaches used by community libs (don’t adopt code, just patterns). ([GitHub][4])

---

## Acceptance Criteria (per milestone)

* **M2 (Word Fluent)**: All README Word scenarios reproducible via fluent with parity of output (verified). ([GitHub][1])
* **M3 (Excel Fluent)**: Create/Load workbook; add sheet; write/read values; add table; format; parity examples normal vs fluent.
* **M4 (PPT Core Write)**: Create deck; at least Title, Title+Content layouts; add title, text box, bullets, picture; Save/Load round‑trip.
* **M5 (PPT Read/Advanced)**: Enumerate all shapes/text; notes read/write; apply simple transitions; chart MVP (embedded xlsx).

---

### References

* OfficeIMO repo (structure, features, tests, targets). ([GitHub][1])
* Blog intro to OfficeIMO.Word (usage background). ([Evotec][2])
* Open XML SDK and ecosystem (context for PPT libs, tools). ([GitHub][4])

---

[1]: https://github.com/EvotecIT/OfficeIMO "GitHub - EvotecIT/OfficeIMO: Fast and easy to use cross-platform .NET library that creates or modifies Microsoft Word (DocX) and later also Excel (XLSX) files without installing any software. Library is based on Open XML SDK"
[2]: https://evotec.xyz/officeimo-free-cross-platform-microsoft-word-net-library/?utm_source=chatgpt.com "OfficeIMO - Free cross-platform Microsoft Word .NET library"
[3]: https://www.nuget.org/packages/documentformat.openxml?utm_source=chatgpt.com "DocumentFormat.OpenXml 3.3.0"
[4]: https://github.com/dotnet/Open-XML-SDK?utm_source=chatgpt.com "dotnet/Open-XML-SDK"