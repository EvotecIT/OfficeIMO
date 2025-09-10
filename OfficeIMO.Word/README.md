# OfficeIMO.Word — .NET Word Utilities

OfficeIMO.Word is a cross‑platform .NET library for creating and editing Microsoft Word (.docx) documents on top of Open XML.

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

- Targets: netstandard2.0, net472, net8.0, net9.0
- License: MIT
- NuGet: `OfficeIMO.Word`
- Dependencies: DocumentFormat.OpenXml, SixLabors.ImageSharp

Quick starts and runnable samples live in `OfficeIMO.Examples/Word/*`.

## Install

```powershell
dotnet add package OfficeIMO.Word
```

## Hello, Word

```csharp
using OfficeIMO.Word;

using var doc = WordDocument.Create("example.docx");
var p = doc.AddParagraph("Hello OfficeIMO.Word");
p.SetBold();
doc.Sections[0].Headers.Default.AddParagraph("Header");
doc.Sections[0].Footers.Default.AddParagraph("Page ");
doc.Sections[0].Footers.Default.AddPageNumber();
doc.Save();
```

## Common Tasks by Example

### Paragraphs and runs
```csharp
var p = doc.AddParagraph("Title");
p.SetBold();
p = doc.AddParagraph("Body text");
p.AddText(" with italic").SetItalic();
p.AddText(" and code").SetFontFamily("Consolas");
```

### Tables
```csharp
var t = doc.AddTable(3, 3);
t[1,1].Text = "Header 1"; t[1,2].Text = "Header 2"; t[1,3].Text = "Header 3";
t.HeaderRow = true; t.Style = WordTableStyle.TableGrid;
t.MergeCells(2,1, 2,3); // row 2, col 1..3
```

### Images
```csharp
var imgP = doc.AddParagraph();
imgP.AddImage("logo.png", width: 96, height: 32);
```

### Headers, footers, and page numbers
```csharp
var sec = doc.Sections[0];
sec.Headers.Default.AddParagraph("Report");
var f = sec.Footers.Default;
f.AddParagraph().AddText("Page ");
f.AddPageNumber();
```

### Hyperlinks
```csharp
doc.AddParagraph().AddHyperLink("OpenAI", new Uri("https://openai.com/"));
```

### Fields
```csharp
var para = doc.AddParagraph();
var field = para.AddField(WordFieldType.Date, wordFieldFormat: WordFieldFormat.ShortDate);
```

### Mail Merge (MERGEFIELD)
```csharp
var para2 = doc.AddParagraph();
// Simple MERGEFIELD with a custom format
para2.AddField(WordFieldType.MergeField, customFormat: "CustomerName");
// Advanced builder for complex instructions/switches
var builder = new WordFieldBuilder(WordFieldType.MergeField)
    .CustomFormat("OrderTotal")
    .Format(WordFieldFormat.Numeric);
para2.AddField(builder);
```

### Footnotes / Endnotes
```csharp
var r = doc.AddParagraph("See note").AddText(" ");
r.AddFootNote("This is a footnote.");
```

### Content controls
```csharp
var sdtPara = doc.AddParagraph("Name: ");
var dd = sdtPara.AddDropDownList(new[]{"Alpha","Beta","Gamma"});
```

### Shapes and charts (basic)
```csharp
var shp = doc.AddShape(ShapeTypeValues.Rectangle, 150, 50);
shp.FillColorHex = "#E7FFE7"; shp.StrokeColorHex = "#008000";
var ch = doc.AddChart(ChartType.Bar, 400, 250);
ch.AddSeries("S1", new[]{1,3,2});
ch.AddLegend(LegendPositionValues.Right);
ch.SetXAxisTitle("Categories");
ch.SetYAxisTitle("Values");
```

### Watermarks
```csharp
doc.SetTextWatermark("CONFIDENTIAL", opacity: 0.15);
```

### Protection
```csharp
doc.ProtectDocument(enforce: true, password: "secret");
```

### Table of Contents (TOC)
```csharp
// Add headings (styles must map to heading levels)
doc.AddParagraph("Chapter 1").SetStyle("Heading1");
doc.AddParagraph("Section 1.1").SetStyle("Heading2");
// Insert TOC near the top (field will update on open)
doc.Paragraphs[0].AddField(WordFieldType.TOC);
```

## Converters — HTML / Markdown / PDF

```csharp
// HTML
using OfficeIMO.Word.Html;
var html = WordHtmlConverter.ToHtml(doc);
var doc2 = WordHtmlConverter.FromHtml("<h1>Hi</h1><p>Generated</p>");

// Markdown
using OfficeIMO.Word.Markdown;
var md = WordMarkdownConverter.ToMarkdown(doc);
var doc3 = WordMarkdownConverter.FromMarkdown("# Title\nBody");

// PDF
using OfficeIMO.Word.Pdf;
doc.SaveAsPdf("out.pdf");
```

## Feature Highlights

- Document: create/load/save, clean/repair, compatibility settings, protection
- Sections: margins, size/orientation, columns, headers/footers (first/even/odd)
- Paragraphs/Runs: bold/italic/underline/strike, shading, tabs, breaks, justification
- Tables: create, merge/split, borders/shading, widths, header row repeat, page breaks
- Images: add from file/stream/base64/URL, wrap/layout, crop/opacity/flip/rotate, position
- Hyperlinks: internal/external with tooltip/target
- Fields: add/read/remove/update (DATE, TOC, PAGE, MERGEFIELD, etc.)
- Footnotes/Endnotes: add/read/remove
- Bookmarks/Cross‑references: add/read/remove
- Content controls (SDT): checkbox/date/dropdown/combobox/picture/repeating section
- Shapes/SmartArt: basic shapes with fill/stroke; SmartArt detection
- Charts: pie/bar/line/combo/scatter/area/radar with axes, legends, multiple series
- Styles: paragraph/run styles, borders, shading

> Explore `OfficeIMO.Examples/Word/*` for complete scenarios.

## Detailed Feature Matrix

- 📄 Documents
  - ✅ Create/Load/Save, SaveAs (sync/async); clean & repair
  - ✅ Compatibility settings; document variables; protection (read‑only recommended/final/enforced)
  - ⚠️ Digital signatures (basic scenarios); ✅ macros (add/extract/remove modules)
- 📑 Sections & Page Setup
  - ✅ Orientation, paper size, margins, columns
  - ✅ Headers/footers (default/even/first), page breaks, repeating table header rows, background color
- ✍️ Paragraphs & Runs
  - ✅ Styles (paragraph/run); bold/italic/underline/strike; shading; alignment; indentation; line spacing; tabs/tab stops
  - ✅ Find/replace helpers
- 🧱 Tables
  - ✅ Create/append; built‑in styles (105); borders/shading; widths; merge/split (H/V); nested tables
  - ✅ Row heights and page‑break control; merged‑cell detection
- 🖼️ Images
  - ✅ From file/stream/base64/URL; alt text
  - ✅ Size (px/pt/EMU); wrap/layout; crop; transparency; flip/rotate; position; read/write EMU sizes
- 🔗 Links & Bookmarks
  - ✅ External/internal hyperlinks (tooltip/target); bookmarks; cross‑references
- 🧾 Fields
  - ✅ Add/read/remove/update (DATE, PAGE, NUMPAGES, TOC, MERGEFIELD, …)
  - ✅ Simple and advanced representations; custom formats
- 📝 Notes
  - ✅ Footnotes and endnotes: add/read/remove
- 🧩 Content Controls (SDT)
  - ✅ Checkbox, date picker, dropdown, combobox, picture, repeating section
- 📊 Charts
  - ✅ Pie/Bar/Line/Combo/Scatter/Area/Radar; axes/legends/series; axis titles
  - ⚠️ Formatting depth varies by chart type
- 🔷 Shapes/SmartArt
  - ✅ Basic AutoShapes with fill/stroke; ⚠️ SmartArt detection/limited operations


## Dependencies & Versions

- DocumentFormat.OpenXml: 3.3.x (range [3.3.0, 4.0.0))
- SixLabors.ImageSharp: 2.1.x
- License: MIT

## Converters (adjacent packages)

- HTML: `OfficeIMO.Word.Html` (AngleSharp) — convert to/from HTML
- Markdown: `OfficeIMO.Word.Markdown` (Markdig) — convert to/from Markdown
- PDF: `OfficeIMO.Word.Pdf` (QuestPDF/SkiaSharp) — export to PDF

> Note: Converters are in active development and will be released to NuGet once they meet quality and test coverage goals. Until then, they ship in‑repo for early evaluation.

## Notes on Versioning

- Minor releases may include additive APIs and perf improvements.
- Patch releases fix bugs and compatibility issues without breaking APIs.

## Notes

- Cross‑platform: no COM automation, no Office required.
- Deterministic save order to keep Excel/Word “file repair” dialogs at bay.
- Nullable annotations enabled; APIs strive to be safe and predictable.
