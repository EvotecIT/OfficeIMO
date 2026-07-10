# OfficeIMO.OpenDocument

`OfficeIMO.OpenDocument` creates and edits ODT, ODS, and ODP files directly. It has no NuGet or project dependencies and does not invoke LibreOffice, Microsoft Office, or UNO at runtime.

```powershell
dotnet add package OfficeIMO.OpenDocument
```

## Create documents

Create an ODT document:

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");

OdtTable table = document.AddTable(2, 2, "Results");
table.Cell(0, 0).Text = "Metric";
table.Cell(0, 1).Text = "Value";
table.Cell(1, 0).Text = "Revenue";
table.Cell(1, 1).Text = "42";

document.Save("summary.odt");
```

Create a sparse ODS workbook:

```csharp
using OdsDocument workbook = OdsDocument.Create();
OdsSheet sheet = workbook.AddSheet("Metrics");
sheet.Cell(0, 0).SetString("Name");
sheet.Cell(0, 1).SetString("Value");
sheet.Cell(1, 0).SetString("Revenue");
sheet.Cell(1, 1).SetDecimal(42.5m);

OdsCell total = sheet.Cell(2, 1);
total.Formula = "of:=SUM([.B2:.B2])";
OdsRecalculationReport calculation = workbook.Recalculate();
if (calculation.FailedCells > 0) {
    Console.WriteLine(calculation.Diagnostics[0].Message);
}

workbook.Save("metrics.ods");
```

Create an ODP presentation:

```csharp
using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Summary");
slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 28, 3), "Native ODP");
slide.AddRectangle(OdfRect.FromCentimeters(2, 5, 8, 3)).FillColor = OdfColor.Parse("#D1E9FF");
slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");
presentation.Save("summary.odp");
```

Convert explicitly between OpenDocument and OfficeIMO Word, Excel, or PowerPoint models by installing the corresponding adapter package. Every conversion returns an `OdfConversionReport` that identifies mapped, approximated, skipped, and unsupported features.

```powershell
dotnet add package OfficeIMO.Word.OpenDocument
dotnet add package OfficeIMO.Excel.OpenDocument
dotnet add package OfficeIMO.PowerPoint.OpenDocument
```

## Edit without flattening the package

Typed objects remain backed by the source XML. A targeted edit rewrites its owning XML part while untouched package entries keep their original bytes.

```csharp
using OdtDocument document = OdtDocument.Open("input.odt");
document.Paragraphs[0].Text = "Updated text";
document.Save("output.odt", new OdfSaveOptions {
    CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
});

IReadOnlyList<string> rewritten = document.LastSaveReport!.RewrittenEntries;
IReadOnlyList<string> lossy = document.LastSaveReport.LossyEntries;
```

New documents use ODF 1.4. Set `OdfCompatibilityProfile.Odf13` when the output needs the ODF 1.3 schema and compatibility profile.

## Supported editing surface

| Area | Current support |
| --- | --- |
| Package | Bounded ZIP/XML loading, direct reading of seekable package streams, manifest updates, deterministic output, metadata, atomic path saves, flat XML projection with loss reporting, unknown-entry preservation |
| ODT | Paragraphs, headings, runs, whitespace controls, styles, lists, tables and spans, links, bookmarks, sections, page layout, headers/footers, page breaks, images, paragraph insertion/deletion tracking |
| ODS | Sparse repeated rows/cells, typed values, OpenFormula text and cached values, bounded formula evaluation/recalculation, styles and data formats, merges, row/column sizing and visibility, sheet order, named ranges, links, validation, print ranges |
| ODP | Slide order and visibility, page size, masters/layouts, text and lists, rectangles, ellipses, lines, groups, transforms, images and crop, tables, speaker notes, backgrounds, transitions, basic shape animations |
| Inspection | Annotations, tracked changes, extension namespaces, scripts, event listeners, external links, embedded objects, formulas, validations, transitions, animations, encryption, and signatures |

Unknown XML, vendor extensions, scripts, embedded content, and unsupported drawing features are preserved when their owning part is not replaced. The library never executes scripts, macros, event listeners, embedded objects, or external links. Formula evaluation is a bounded, side-effect-free parser for the documented local subset; it does not execute active content or fetch data.

`OdfCapabilityCatalog.Advanced` provides stable capability IDs and distinguishes editable subsets, preserved content, inspection, and detected-but-unsupported features.

## Explicit boundaries

- Formula evaluation covers arithmetic, comparisons, concatenation, cell/range references, and common aggregate/math functions. External data, volatile functions, matrix formulas, and the complete OpenFormula language are not included.
- Tracked-change editing covers paragraph insertions and deletions. Arbitrary inline merges and conflict resolution remain preservation-oriented.
- Animation editing covers basic shape-attribute effects and fade-in timing. Advanced timing trees are preserved when untouched.
- Encrypted packages are detected and rejected before editing.
- Changed signed packages fail by default because saving would invalidate signatures. An explicit save option can remove invalidated signature entries.
- Signature creation and cryptographic validation, pivot-table editing, and complete chart editing are outside the current surface.
- Flat XML variants (`.fodt`, `.fods`, `.fodp`) can be opened and written, including embedded raster images. Exotic embedded objects and package-only features may not project losslessly.
- `OdsSheet.Merge` rejects merges above its default 100,000-cell materialization limit. Use the overload with an explicit lower limit when processing untrusted dimensions.
- Unknown package entries and extension XML are always preserved by package editing. Explicit format conversion and flat XML projection report content they cannot carry through `OdfConversionReport` and `OdfSaveReport.LossyEntries`.

The package targets `netstandard2.0`, `net8.0`, and `net10.0`, plus `net472` on Windows. CI checks generated ODF 1.3 and 1.4 XML against pinned OASIS Relax NG schemas, then opens and resaves the generated packages with the runner's reported LibreOffice version.
