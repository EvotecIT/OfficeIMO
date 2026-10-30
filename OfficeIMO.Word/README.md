# OfficeIMO.Word - Word documents for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

`OfficeIMO.Word` is the main Word document package in the OfficeIMO family. It creates, edits, inspects, converts, and saves `.docx` files, and can import and write supported legacy `.doc` files, without COM automation and without Microsoft Office installed.

## Inspection without image copies

`Inspect(includeImageBytes: false, cancellationToken)` keeps the
native semantic document and image descriptors while leaving their `Bytes` null.
The ordinary parameterless snapshot still includes embedded image bytes.
`WordLoadOptions.AllowLegacyProjection = false` rejects binary DOC before its
picture/model projection; the default continues to support that import.

`InspectEmbeddedPayloads(options, cancellationToken)` inventories inert package,
OLE and ActiveX descriptors and provides selected bounded `OpenRead(ordinal)`
streams. `OpenEmbeddedPayload(id, maximumBytes, cancellationToken)` is the direct
single-payload convenience API. Streams are complete, read-only and caller-owned;
they survive disposal of the Word document. Keep the document open while
requesting a read. The [Core payload contract](../OfficeIMO.Core/README.md#bounded-embedded-payload-reads)
defines count, member, aggregate and cumulative work limits, unknown lengths and
optional hashing. `InspectVbaProjectMetadata()` inventories the part and signature
relationship without reading binary module content; `ModuleInspectionRequested`
distinguishes this from full VBA metadata inspection.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Word
```

Upgrading an existing application? The regular create, load, edit, and save workflow keeps the same shape. See the [migration guide](../MIGRATION.md) for package, enum, and type replacements.

Document authoring, reading, signature inspection, and safe signed-package handling do not require the security
package. Install it only when the application creates or cryptographically validates OPC or VBA signatures:

```powershell
dotnet add package OfficeIMO.Security
```

## Quick start

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");

document.AddHeading("Quarterly report", level: 1);
document.AddParagraph("Created with OfficeIMO.Word.");

var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";
table.Rows[1].Cells[0].Paragraphs[0].Text = "Documents";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Generated";
table.RepeatHeaderRowAtTheTopOfEachPage = true;
table.Style = WordTableStyle.TableGrid;

document.Save();
```

`AddHeadersAndFooters()` returns the attached default parts for the document or a section:

```csharp
var (header, footer) = document.AddHeadersAndFooters();
header.AddParagraph("Quarterly report");
footer.AddParagraph("Internal use");
```

`AddPageNumbering`, `AddFootnoteProperties` and `AddEndnoteProperties` return
editable settings for the document's first section. The section methods return
settings for that section:

```csharp
var numbering = document.AddPageNumbering(startNumber: 1);
numbering.NumberingFormat = WordNumberFormat.UpperRoman;
var notes = document.AddFootnoteProperties(startNumber: 1);
notes.NumberingRestart = WordNoteNumberRestart.EachSection;
```

`AddMacro(path, bytes or stream)` returns the attached `WordMacroProject`.
Use its `ToBytes`, `Save`, `Modules` and `Remove` members while the document is
open. Replacing or removing the project invalidates earlier project handles.
Reading and attaching a project never executes its code.

## Editable list levels

`AddListLevel` returns the attached `WordListLevel`. Its indents use physical lengths; numbering and bullet changes stay on that level.

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create();
WordList list = document.AddCustomList();
WordListLevel level = list.AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Arial");
level.IndentationLeft = OfficeLength.Centimeters(1.25);
level.IndentationHanging = OfficeLength.Points(18);
list.AddItem("Review the deployment");
document.Save("custom-list.docx");
```

## Printed reports

`WordPaperStyle` applies shared typography, page chrome, covers, metric strips and quiet tables. Configure its colors, fonts and table density once, then use the same style across reports:

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("assessment.docx");
var paper = new WordPaperStyle { TableFontSize = global::OfficeIMO.OfficeLength.Points(9) };
var cover = paper.AddCover(document, "ASSESSMENT", "Domain assessment", "Example fixture");
var chrome = paper.AddPageChrome(document, "Domain assessment", "Example reports", "example.test");
var table = document.AddTable(2, 2);
paper.SetCellText(table.Rows[0].Cells[0], "Hostname");
paper.SetCellText(table.Rows[0].Cells[1], "Status");
paper.SetCellText(table.Rows[1].Cells[0],
    WordPaperText.WithIdentifierBreaks("mail.example.test", '.', '-'));
paper.SetCellText(table.Rows[1].Cells[1], "Pass");
paper.ApplyTable(table);
paper.FitTableColumnsToHeaders(table, 70, 30);
paper.ApplyTypography(document);
document.Save();
```

`WordPaperText.WithIdentifierBreaks(value)` adds invisible breaks at Latin case/digit boundaries; the separator overload above keeps hostname labels intact. Use `RemoveOptionalBreaks` when exporting plain text. For count charts, `chart.SetPaperStyle().SetCountAxis()` applies paper typography and a whole-number value axis.

Run width and tracking are readable and writable on `WordParagraph`:

```csharp
var text = document.AddParagraph("Condensed text");
text.CharacterScale = 75; // Percentage, from 1 through 600.
text.Spacing = 10;        // Twips: 10 adds half a point per glyph.
```

Set `CharacterScale` to 100 for an explicit normal-width override, or to `null`
to remove the direct setting and restore style inheritance. DOCX and supported
native DOC round trips retain run width and tracking in body text, tables, links,
headers, footers and notes. Native DOC also preserves character width in paragraph
styles and document defaults, including explicit normal-width overrides.

`WordHyperLink.Tooltip` and `TargetFrame` retain hyperlink descriptions and target
windows through DOCX and supported native DOC conversions. External links and
internal bookmark links preserve these settings in body text, tables, headers,
footers and notes. Adjacent links can retain different descriptions or targets.
Native DOC saving rejects line breaks in tooltip or target-window metadata;
these values remain supported in DOCX.

When a Word file comes from an untrusted source, pass the bounded load profile before parsing it:

```csharp
using var incoming = WordDocument.Load("upload.docx", WordLoadOptions.UntrustedDefaults);
```

This profile rejects macros, embedded payloads, ActiveX, and external relationships. Ordinary load options retain compatibility with documents containing those parts; `PackageSecurity` can be set explicitly for a different policy.

Capture structural observations from the exact loader input when an inspection
workflow needs to retain original-package evidence:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("review.docx", new WordLoadOptions {
    AccessMode = DocumentAccessMode.ReadOnly,
    PackageSecurity = new OfficePackageSecurityOptions { MaxPackageBytes = 64L * 1024 * 1024 },
    CaptureSourcePackageSecurity = true
});
Console.WriteLine(document.SourcePackageSecurityReport?.ContainerKind);
Console.WriteLine(document.SourcePackageSecurityScope);
var loadedSafety = document.InspectLoadedContentSafety();
```

The captured report remains source evidence after model edits; it is not a report
of a later save. `SourcePackageSecurityScope` identifies the source container,
including an encrypted outer container when that load route supplies it. Loaded
content-safety inspection examines the current model without reopening or saving
the file. Package findings describe bounded structure and policy, not signature
validity or a guarantee that content is safe. Capture is opt-in and does not replace
the load policy or native import limits.

Create and edit content through the document object model:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");

document.AddParagraph("Quarterly report").SetStyle(WordParagraphStyles.Heading1);
document.AddParagraph("Created with OfficeIMO.Word.");
document.AddCallout(OfficeCalloutKind.Info, "Review", "Check the quarterly totals.");
document.AddCodeBlock("return total;", "csharp");

document.Save();
```

## Shapes and lines

Shape dimensions, positions and line weights use `OfficeLength`. `AddShape`
creates a VML shape; its `Width`, `Height`, nullable `Left`, `Top` and
`StrokeWeight` properties use the same lengths. `AddShapeDrawing` creates a
DrawingML shape; `TryGetLayoutSnapshot` reads its persisted frame.

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("shapes.docx");
var shape = document.AddShape(WordShapeType.RoundedRectangle,
    OfficeLength.Centimeters(4), OfficeLength.Points(36),
    fillColor: OfficeColor.Parse("#DDEEFF"));
shape.Left = OfficeLength.Points(18);
shape.Top = OfficeLength.Points(24);
shape.StrokeWeight = OfficeLength.Points(1);

document.AddParagraph().AddShapeDrawing(WordShapeType.Ellipse,
    OfficeLength.Inches(1), OfficeLength.Centimeters(1),
    OfficeLength.Points(36), OfficeLength.Points(72));
document.AddParagraph().AddLine(OfficeLength.Zero, OfficeLength.Zero,
    OfficeLength.Centimeters(4), OfficeLength.Zero, OfficeColor.Blue);
document.Save();
```

## Paragraph formatting and inheritance

`AddPageNumber(includeTotalPages: true)` uses its source run's typography for both page fields and the separator. Font family, size and colour can be set before or after adding the fields; unrelated runs keep their own formatting.

For paper tables, call `WordPaperStyle.FitTableColumnsToHeaders(table, relativeWidths)` after populating and styling the header. It fits the requested column ratios within the table/container width while reserving each header's longest word plus padding. Hyphenated labels stay intact. It requires rectangular, unmerged rows and throws if the header minima cannot fit; call it before merging note rows and again after changing page geometry or header typography.

`WordTableOfContent.TitleParagraph` exposes the existing contents heading for ordinary paragraph styling and pagination. For example, set its `Style` to `WordParagraphStyles.Heading2` and `PageBreakBefore` to `true` to begin a dedicated contents page. Refreshing entries retains the authored heading, and PDF conversion honors its paragraph formatting.

Use `LineSpacing = WordLineSpacing.Multiple(1.5)` for 1.5 lines, `WordLineSpacing.Exact(OfficeLength.Points(18))` for an exact height, or `WordLineSpacing.AtLeast(OfficeLength.Points(18))` for a minimum height. Assigning `null` removes both direct line attributes while preserving paragraph spacing before and after. A native numeric line value with an omitted rule uses automatic line multiples. Rule-only and unsupported native declarations remain untouched when read; the typed getter returns `null` for these declarations. Effective style inheritance remains the native layout and conversion engine's responsibility. Native DOC saving carries document-default spacing into root paragraph styles without changing the source styles.

Paragraph pagination controls support explicit on, off, and inherited values. Use the nullable `*Override` properties to disable a setting enabled by a paragraph style, or set them to `null` to remove direct formatting:

```csharp
var paragraph = document.AddParagraph("Continue on this page");
paragraph.PageBreakBeforeOverride = false;
paragraph.KeepWithNextOverride = true;
paragraph.KeepLinesTogetherOverride = null;
paragraph.AvoidWidowAndOrphanOverride = true;
paragraph.ContextualSpacing = true;
paragraph.OutlineLevel = 9; // Body text; heading levels 1-9 use values 0-8.
```

`WordParagraphStyleDefinition` exposes the same controls without the `Override` suffix. `ContextualSpacing`, `SuppressLineNumbers`, `SuppressAutoHyphens`, and `MirrorIndents` are nullable on both paragraphs and style definitions. These values survive DOCX and supported native DOC saves, including explicit false values. Existing Boolean pagination properties keep their previous behavior; use the nullable properties when style inheritance matters.

`OutlineLevel` returns directly authored formatting. Word fixes the effective outline level of paragraphs using built-in Heading1–Heading9 styles to the corresponding heading level, even if a different direct value is stored.

PDF conversion honors an explicit `PageBreakBeforeOverride = false` even when the paragraph's style starts paragraphs on a new page. Storing line-number suppression, hyphenation suppression, mirrored indentation, or outline levels does not establish PDF rendering support for those features; see the [Word PDF conversion contract](../OfficeIMO.Word.Pdf/README.md) and [native DOC limits](../Docs/officeimo.word.legacy-doc-compatibility.md).

`document.CompatibilitySettings.SplitPageBreakAndParagraphMark` controls the blank line after a manual page break that ends a paragraph. In older Word layout, enabling it moves the paragraph mark and its line height onto the next page. DOCX retains both explicit values; Word 2013 compatibility mode ignores the stored option during layout. Native DOC import uses the legacy enabled behavior and retains that layout when saved as DOCX. Native DOC saving rejects an explicitly disabled value; save as DOCX to retain it.

`document.CompatibilitySettings.DoNotUseIndentAsNumberingTabStop` controls whether a list's paragraph indent acts as an implicit numbering tab stop. In Word 2003, 2007 and 2010 compatibility modes, enabling it lets an explicit or default numbering tab determine the first line while continuation lines retain the paragraph indent. An explicit tab before the continuation indent is honored in every compatibility mode. Ordinary numbered headers and footers use the same paragraph layout as body content. For stories without tables or SECTIONPAGES fields, text boxes or indents outside either available text frame retain the bounded text layout; exact numbering-tab geometry in those stories remains a fidelity limitation. Tables and SECTIONPAGES fields select the shared flow engine and retain its documented limits. Word 2013 compatibility mode ignores the stored flag during PDF layout. DOCX retains explicit true and false values. Native DOC imports project their effective legacy behavior as enabled; native DOC saving rejects an explicitly disabled value. Save as DOCX when that stored value must be retained.

## Page sizes and orientation

Set a section's paper preset and orientation through `PageSettings`:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

var page = document.Sections[0].PageSettings;
page.PageSize = WordPageSize.Tabloid;
page.Orientation = OfficePageOrientation.Landscape;
WordPageSizeDefinition? definition = WordPageSizes.GetDefinition(WordPageSize.Tabloid);
```

Presets include Letter, Legal, Statement, Executive, A3–A6, JIS B4/B5, Tabloid, C sheet, and number 9, number 10, DL, C5, C4, B5 and Monarch envelopes. `WordPageSize.B5` retains its established JIS dimensions of 182 × 257 mm; `EnvelopeB5` measures 176 × 250 mm. The shared `OfficePageSizes` catalog owns physical dimensions.

`Width` and `Height` expose custom dimensions as nullable `OfficeLength` values. Changing `Orientation` swaps those dimensions. The preset getter recognizes matching dimensions when a producer omits the optional printer code, with a one-twip tolerance for unit rounding. Native DOC retains physical dimensions and orientation; its imported page settings do not carry the DOCX printer code. PDF conversion preserves stored width and height, including a wide custom page without an orientation flag, and supports explicit export orientation overrides.

Set `section.Margins.Gutter = OfficeLength.Points(18)` to reserve binding space. `document.Settings.GutterAtTop` places that space above the body; otherwise `section.RtlGutter` selects the right edge and the default is the left edge. `document.Settings.MirrorMargins` stores the document's facing-page margin setting. These settings survive DOCX and supported native DOC saves. Present on/off XML elements without a `val` attribute remain enabled. The [PDF conversion contract](../OfficeIMO.Word.Pdf/README.md) describes rendering support separately.

Page dimensions, section margins, paragraph indentation and before/after spacing, cell padding, and `WordTableRow.Height`/`MinimumHeight` accept `OfficeLength`. DOCX stores these values as whole twips, so fractional values round when assigned. A null indentation, spacing or cell-padding value removes the direct override.

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("physical-layout.docx");
var section = document.Sections[0];
section.PageSettings.Width = OfficeLength.Inches(8.5);
section.PageSettings.Height = OfficeLength.Inches(11);
section.Margins.Left = OfficeLength.Centimeters(2.5);
var paragraph = document.AddParagraph("A measured paragraph");
paragraph.IndentationFirstLine = OfficeLength.Inches(0.5);
paragraph.LineSpacingAfter = OfficeLength.Points(8);
var row = document.AddTable(1, 1).Rows[0];
row.MinimumHeight = OfficeLength.Points(20);
row.Cells[0].MarginLeft = OfficeLength.Millimeters(2);
document.Save();
```

`document.AddSection(WordSectionBreakType.OddPage)` returns the new section. Its `BreakType` property gets or changes how that section starts relative to the preceding section; the preceding section retains its own start type. `AddSection()` starts on the next page and continues page numbering. Set a new section's numbering restart explicitly when needed. All five start types survive DOCX and supported native DOC saves.

## Section columns

Set `ColumnCount` and nullable `ColumnsSpace` for equal-width columns, for example `section.ColumnsSpace = OfficeLength.Points(36)`. Setting it to `null` omits the stored gap; layout uses a 36-point default. The gap rounds to whole twips at storage (36.03 points becomes 721 twips, or 36.05 points). The property retains its signed 32-bit twip storage range; native DOC saving accepts equal-column gaps only from zero through 65535 twips.

Set `ColumnDefinitions` to author or inspect individual widths and following gaps as `OfficeLength` values:

```csharp
var section = document.Sections[0];
section.ColumnDefinitions = new[] {
    new WordSectionColumn(OfficeLength.Points(100), OfficeLength.Points(20)),
    new WordSectionColumn(OfficeLength.Points(300), OfficeLength.Zero)
};
OfficeLength firstWidth = section.ColumnDefinitions[0].Width;
```

The property takes a snapshot, rounds lengths to whole twips and synchronizes the column count. Widths must round to a positive 32-bit twip value; gaps must be non-negative and round within the same integer range. Replace the definitions to change an explicit layout's count; an empty list restores equal widths while retaining the count and default spacing.

An omitted `SpaceAfter` has an effective gap of zero for unequal columns. `ColumnsSpace` applies to equal-width columns. DOCX retains the omitted individual value; native DOC writes its effective zero explicitly.

DOCX preserves these settings. Native DOC preserves indexed widths and individual gaps, with up to 44 columns, widths from 718 through 32767 twips and gaps from zero through 32767 twips. Saving a layout outside those native limits fails before creating output. Invalid or incomplete native indexed records produce an import diagnostic. PDF column flow is described separately in the [conversion contract](../OfficeIMO.Word.Pdf/README.md).

## Hidden text

`WordParagraph.Hidden` controls the current run, including hyperlink and inline content-control runs. Set it to `true` to hide the text, `false` to override a hidden style, or `null` to remove the direct setting and inherit. The getter reports the directly authored value. Hidden text remains in the document; PDF export omits it according to the effective formatting. DOCX preserves the direct setting, and native DOC preserves supported hidden formatting.

```csharp
var run = document.AddParagraph("Internal reference");
run.Hidden = true;
run.Hidden = false; // Explicitly visible, even under a hidden style.
run.Hidden = null;  // Restore inheritance.
```

Native DOC saving supports field display runs with one effective formatting set. Equivalent runs can use different direct settings, such as an omitted hidden setting and an explicit visible setting. A field whose display runs have different effective formatting, including visibility inherited from a paragraph or table style, raises `NotSupportedException`.

## Page and paragraph borders

`WordBorders` and `WordParagraphBorders` expose each side's `Size` and `Space` as nullable `OfficeLength` values. Border widths round to eighths of a point; spacing rounds to whole points. Midpoint ties round away from zero. Null removes the attribute, while zero remains explicit. Negative or overflowing lengths fail before changing the document.

```csharp
paragraph.Borders.SetBorder(WordParagraphBorderType.Left, WordBorderStyle.Single,
    OfficeColor.Red, OfficeLength.Points(1.25), shadow: false);
paragraph.Borders.LeftSpace = OfficeLength.Points(3);
paragraph.Borders.LeftSpace = null; // Remove the direct spacing attribute.
```

## Paragraph tab stops

Use `AddTabStop` to configure a paragraph's explicit tab positions with `OfficeLength`. Signed positions round to the nearest twip (1/20 point) when written. `ClearTabStops()` removes those local stops without changing paragraph spacing, alignment, or inherited defaults.

```csharp
var paragraph = document.AddParagraph("Label\t12.50");
document.Settings.DecimalSymbol = ".";
paragraph.ClearTabStops().AddTabStop(OfficeLength.Inches(1), WordTabAlignment.Decimal);
```

`Settings.DecimalSymbol` stores the DOCX document's decimal separator. It accepts one noncontrol Unicode character; assigning `null` removes the explicit setting. OfficeIMO PDF export uses that separator for body and table decimal tabs, so grouping punctuation in values such as `1,234.56` does not become the alignment anchor. Other applications also apply their locale and language handling when rendering decimal tabs.

## Comments on table cells

`AddCellComments` adds root comments across tables in one batch. The cell's direct paragraphs form the anchor range; cell values remain unchanged.

```csharp
using var document = WordDocument.Create();
var table = document.AddTable(2, 2);
table.Rows[1].Cells[0].Paragraphs[0].Text = "Review this value";

var comments = document.AddCellComments(new[] {
    new WordCellComment(table.Rows[1].Cells[0], "Reviewer", "R",
        "Check the source figure.", DateTime.UtcNow)
});
comments[0].AddReply("Editor", "E", "Checked.");
comments[0].MarkResolved();
document.Save("review.docx");
```

Targets must be attached cells in this document's body. Empty cells, empty comment text and empty author names are supported. A null creation date leaves the timestamp unspecified. Line feeds, tabs and whitespace are preserved; carriage returns, U+2028 and invalid XML characters are rejected instead of silently normalized. `WordCellComment.CanPreserveText` checks the plain-text contract before importing content. Invalid input and cancellation before batch application do not add comments. Cancellation is checked during enumeration and preparation; once application starts, the batch finishes without cancellation checks between anchors.

Appending comments to older DOCX files assigns missing paragraph identities to existing roots before adding modern metadata. Existing replies and resolved states retain their own comment associations.

## What it does

- Creates, loads, edits, saves, and appends `.docx` documents.
- Opens supported Word 97-2003 `.doc` files through the normal `WordDocument.Load(...)` path and projects them into the regular OfficeIMO Word model.
- Writes native `.doc` files for the currently supported simple-document subset, with preflight checks that block unsupported content before saving.
- Converts supported `.doc` and `.docx` files with `OfficeIMO.OfficeConverter.Convert(...)`, using the same import diagnostics and save preflight as normal load/save workflows.
- Works with paragraphs, runs, styles, sections, headers, footers, page numbers, tables, images, hyperlinks, bookmarks, fields, footnotes, endnotes, content controls, charts, shapes, and document protection.
- Applies optional shared package-security policy before parsing Open XML or compound DOC files, and preflights read, edit, template, render, and save capabilities.
- Inspects and manages VBA and embedded package/OLE/ActiveX payload bytes without executing active content.
- Creates and validates cross-platform OPC XML package signatures and managed VBA legacy, agile, and V3 signatures through an explicitly supplied `IOfficeSecurityProvider`; structural inspection remains provider-free.
- Exports document page ranges as managed PNG, JPEG, TIFF, WebP or SVG pages through `ExportImages(...)` and `SaveAsImages(...)`.
- Chart series Add methods return `WordChartSeries`; set its `Name` or call `Remove()` to edit that exact series. `AddLegend` returns an editable `WordChartLegend`. `AddChartAxisX` returns category settings for subsequent series.
- Keeps Office automation out of the runtime path, making it suitable for services, scheduled jobs, CI, desktop apps, and automation hosts.
- Uses one object model for paragraphs, tables, lists, charts, images, and content blocks.
- Uses `OfficeIMO.Drawing` for shared colors, image metadata, page rendering, and the reusable math expression tree.

Native `.doc` lists retain supported numbering definitions, marker formatting, abstract starts and per-instance restarts through the normal load/save API. Missing or malformed native list definitions appear in `LegacyDocUnsupportedFeatures` as `Numbering` losses and block normal saving. The [DOC compatibility contract](../Docs/officeimo.word.legacy-doc-compatibility.md) describes supported levels and numbering limits.

The List Paragraph style retains contextual spacing and its supported paragraph formatting in native DOC, including when the style is used only in a header or footer.

Advanced drawing, structured comparison, field evaluation, and evidence boundaries are documented in [Word advanced editing and evidence contracts](../Docs/officeimo.word-advanced-contracts.md). The contracts distinguish persisted drawing geometry from desktop Word layout, detected relocation from native move revisions, and supported legacy DOC writing from arbitrary DOC authoring.

For untrusted files, capability preflight, binary DOC/XLS/XLSB loss policies,
macro and embedded-payload handling, and the executable compatibility corpus,
see the [Word and Excel interoperability guide](../Docs/officeimo.word-excel-interoperability.md).

## Examples

The quick start shows the smallest useful document. These examples show the kinds of document work that belong in `OfficeIMO.Word` itself.

### Paragraphs and runs

```csharp
var paragraph = document.AddParagraph("Status: ");
paragraph.AddText("Approved").Bold = true;
paragraph.AddText(" on ");
paragraph.AddText(DateTime.Today.ToString("yyyy-MM-dd")).Italic = true;
```

`KerningMinimumFontSize` controls the minimum font size at which a run uses kerning.
Set it to `null` to inherit the threshold or `0` to disable kerning. Values from `0` to
`1638` points round to the nearest half point and persist in DOCX and supported native DOC content.
The property reads the run's explicit threshold; an inherited value reads as `null`.

```csharp
var title = document.AddParagraph("AV typography");
title.FontFamily = "Arial";
title.FontSize = global::OfficeIMO.OfficeLength.Points(24);
title.KerningMinimumFontSize = global::OfficeIMO.OfficeLength.Points(8);
```

### Footnotes and endnotes

Add notes to body paragraphs, including paragraphs in table cells:

```csharp
document.AddParagraph("A statement with supporting detail.")
    .AddFootNote("The supporting detail.");

var cell = document.AddTable(1, 1).Rows[0].Cells[0];
cell.Paragraphs[0].Text = "A statement in a table.";
cell.Paragraphs[0].AddEndNote("Additional background.");
```

The document and section `FootNotes` and `EndNotes` collections include references in nested tables and follow their body source order. Each collected note retains its own reference, including when several references share a run. Edit its `Paragraphs` or call `Remove()` to change the selected note while preserving sibling references and surrounding text. `AddFootNote` rejects header and footer paragraphs before changing the document.

### Tables with structure

`table.Rows[0].AddCell()` appends a cell and returns the attached `WordTableCell` for editing.

```csharp
var table = document.AddTable(3, 3);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Status";
table.RepeatHeaderRowAtTheTopOfEachPage = true;
table.Style = WordTableStyle.TableGrid;

table.Rows[1].Cells[0].Paragraphs[0].Text = "Documents";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Operations";
table.Rows[1].Cells[2].Paragraphs[0].Text = "Ready";

table.MergeCells(rowIndex: 2, columnIndex: 0, rowSpan: 1, colSpan: 3);
table.Rows[2].Cells[0].Paragraphs[0].Text = "Generated by OfficeIMO.Word";
```

Use `table.Rows[0].MinimumHeight = OfficeLength.Points(20)` to set a minimum height that grows with cell content. `Height` sets an exact physical height; setting either property replaces the other constraint. Setting `MinimumHeight` to null removes a minimum constraint while preserving an exact height.

Table-level shading uses `table.StyleDetails.ShadingFillColorHex`, `ShadingColorHex` and `ShadingPattern`; the color properties also accept `OfficeColor` values. These settings read and write direct table formatting, including imported tables without a named style. For example, set `ShadingFillColorHex = "80C0FF"` for the fill and `ShadingPattern = WordShadingPattern.Clear`. With positive cell spacing, the table fill colors the gaps between cells. Set the fill to `"auto"` to clear inherited table fill, or set `ShadingFillColor` or `ShadingPattern` to null to remove direct shading and restore style inheritance. Clearing `ShadingColor` retains the fill and pattern. Cell shading remains independent through `table.Rows[row].Cells[column]`; table gap formatting does not replace cell-style colors.

DOCX retains table-level shading. Native DOC saving accepts clear shading patterns and Word 97–2003 palette colors; it rejects visible table gap shading when cell spacing is positive. Zero spacing or an automatic gap fill addresses the spacing restriction; unsupported patterns and colors still require changes before saving as DOC. Cell fills retain their own direct, conditional and ordinary style precedence, including automatic fills that clear a lower-priority color.

### Headers and footers

```csharp
document.HeaderDefaultOrCreate.AddParagraph("Internal report");
document.FooterDefaultOrCreate.AddParagraph()
    .AddText("Page ")
    .AddPageNumber();
```

### Images

```csharp
var paragraph = document.AddParagraph();
var image = paragraph.AddImage("logo.png", width: OfficeLength.Pixels(160), height: OfficeLength.Pixels(64));
```

Crop an embedded picture using lengths measured on the original image:

```csharp
image.CropLeft = OfficeLength.Centimeters(0.5);
image.CropRight = OfficeLength.Centimeters(0.25);
```

The physical crop properties use the source image's pixel dimensions and horizontal or vertical resolution. They keep the displayed image scale and resize its frame, so the same crop selects the same source area after resizing. Floating pictures also adjust their anchor to keep the original picture in place, including rotation and mirrors. Assign `null` to restore an edge. `SetCrop(left, top, right, bottom)` applies four `OfficeLength?` values together, validating both axes before changing either.

The `CropLeftRatio`, `CropTopRatio`, `CropRightRatio` and `CropBottomRatio` properties select fractions of the original image: `0.1` means 10% of the original dimension. They leave the displayed frame unchanged. Negative physical crops expand the DOCX frame with empty margins; Word-to-PDF currently reports `NativeImageCropApproximated` for these source outsets. Physical cropping requires a readable embedded source and a floating position that can be resolved without pagination; use the ratio properties when only the source selection should change.

Use `image.Clone(destinationParagraph)` to copy an image to another paragraph,
header, footer, or document. The destination owns its image relationships and fresh DrawingML
identifiers. Cloning preserves VML shapes and their referenced definitions, and remaps
embedded SVG extension resources along with the primary image.

Managed page previews and PDF export prefer an embedded SVG, including Word's
`asvg:svgBlip` extension. If the SVG is unavailable or uses unsupported features,
the renderer uses the image's bitmap fallback when one is present. SVG support
uses the shared managed drawing engine; images with neither a supported vector
nor a usable fallback remain visible in conversion diagnostics.

Optimize embedded images before saving a separate copy:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Word;

using var document = WordDocument.Load("input.docx");
var options = new WordImageOptimizationOptions {
    Mode = OfficeImageOptimizationMode.DownsampleAndRecompress,
    TargetDpi = 144,
    JpegQuality = 85
};
WordImageOptimizationReport report = document.OptimizeImages(options);
document.Save("optimized.docx");
Console.WriteLine($"{report.OptimizedCount} images changed; {report.BytesSaved} media bytes saved.");
```

`AnalyzeImageOptimization(options)` evaluates the same candidates without changing media; it also works on a read-only document. Reports contain one item per unique media part, original/final dimensions and encoded sizes, reference counts, candidate metadata evidence, and preservation reasons. `RequiredStagedBytes` gives the combined original/candidate bytes needed to apply the proposed replacements; preserved parts do not consume that budget. Measure whole-file savings after saving: ZIP compression and document normalization can change that result.

`Downsample` is the default mode. `Recompress` retains pixel dimensions and explicitly re-encodes JPEGs at the selected quality; `DownsampleAndRecompress` combines both. PNG, JPEG, single-page TIFF, and static WebP retain their media type. BMP and static GIF candidates use PNG carriers while keeping drawing relationship IDs. Larger candidates remain unchanged by default. Animated/multi-page images, unsupported vectors, unreferenced media, and unsafe placement geometry are preserved. Linked images are counted without fetching them.

The inventory visits all package stories and XML media references, including body paragraphs, tables, headers, footers, notes, comments, DrawingML, and VML. Shared media uses the greatest pixel demand across its placements; crop edges increase that demand. Cropping remains editable and hidden crop pixels remain in the image. Tile fills, unresolved group transforms, relative sizing, and other unknown placements block downsampling for the shared part; explicit same-size recompression can still apply.

Candidates are staged before mutation, with cooperative cancellation and rollback. Default limits are 32 MiB per encoded image, 128 MiB of staged originals/candidates, and 10,000 package parts. Metadata loss, including deliberate `Strip` or selective removal, blocks replacement unless `AllowMetadataLoss` is explicit. Signed packages block mutation unless `SignedDocumentPolicy` permits invalidation; saving that package also requires the corresponding save policy. `CompressionQuality` records Word's compression-state hint; use `OptimizeImages` to change the encoded bytes.

For supported legacy DOC inline pictures, optimization uses the same projected media API and native save preflight. The [legacy compatibility guide](../Docs/officeimo.word.legacy-doc-compatibility.md) describes the supported subset. The [workflow runner](../OfficeIMO.Workflows/README.md#optimize-embedded-word-images) provides source-preserving file and batch publication.

### Hyperlinks and bookmarks

```csharp
document.AddParagraph("Jump target").AddBookmark("target-section");
document.AddParagraph()
    .AddHyperLink("Open project site", new Uri("https://github.com/EvotecIT/OfficeIMO"));
document.AddParagraph()
    .AddHyperLink("Jump inside document", "target-section", addStyle: true);
```

### Fields and table of contents

```csharp
document.AddParagraph("Chapter 1").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Section 1.1").Style = WordParagraphStyles.Heading2;
WordField contents = document.Paragraphs[0].AddField(WordFieldType.TOC);
contents.UpdateField = true;
contents.Paragraph.Style = WordParagraphStyles.Normal;
```

### Plain DOCX templates

Use ordinary `{{Name}}` placeholders when a Word-authored layout should bind directly to an application model. Scalar placeholders retain the formatting of the first template run, while repeated and conditional marker paragraphs can surround paragraphs, lists, or tables.

```csharp
using var document = WordDocument.Load("invoice-template.docx");

var values = new Dictionary<string, object?> {
    ["Customer"] = new Dictionary<string, object?> { ["Name"] = "Northwind Traders" },
    ["Lines"] = new object[] {
        new Dictionary<string, object?> { ["Description"] = "Assessment", ["Amount"] = 1200m },
        new Dictionary<string, object?> { ["Description"] = "Implementation", ["Amount"] = 3400m }
    },
    ["Portal"] = new WordTemplateHyperlink("Open invoice", new Uri("https://example.com/invoices/42"))
};

WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();
document.Save("invoice-42.docx");
```

Inside the DOCX, use `{{Customer.Name}}` for values and put block markers on their own paragraphs:

```text
{{#each Lines}}
{{Description}} — {{Amount}}
{{/each Lines}}
```

The dictionary overload is trimming and NativeAOT safe. A POCO overload is available for convenience and is annotated because it reflects over public properties. See the [template guide](https://officeimo.com/docs/word/templates/) for conditions, nested blocks, images, diagnostics, and the executable proof workflow.

### Mail merge fields

```csharp
var merge = document.AddParagraph();
merge.AddText("Customer: ");
merge.AddField(new WordFieldBuilder(WordFieldType.MergeField)
    .AddInstruction("CustomerName"));

var totalField = new WordFieldBuilder(WordFieldType.MergeField)
    .AddInstruction("OrderTotal")
    .SetFormat(WordFieldFormat.Arabic);
merge.AddField(totalField);
```

For a strict merge, use the structured result rather than assuming every field was bound:

```csharp
WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
    document,
    new Dictionary<string, string> {
        ["CustomerName"] = "Ada Lovelace",
        ["OrderTotal"] = "1234.5"
    });

report.EnsureComplete();
```

The report distinguishes merged fields, missing values, and unsupported formatting. `ExecuteBatchWithReport(...)` retains the same evidence for every output record.

### OPC package signatures

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
WordDocument.SignPackage("report.docx", security, "CERTIFICATE-THUMBPRINT");

using WordDocument signed = WordDocument.Load("report.docx");
WordSignatureValidationReport validation = signed.ValidateSignatures(
    security,
    new WordSignatureValidationOptions());
```

OPC package signing does not sign VBA code. `WordSigningCapabilities.Package` and
`WordSigningCapabilities.MacroProject` report the two surfaces independently. `InspectSignatures()` remains available
without `OfficeIMO.Security`; it projects the shared bounded OPC inspector and does not claim digest, signature, or
certificate trust. The cross-host `OfficeDocument.Inspect(path, options).PackageSignatures`, `ValidatePackageSignatures(...)`,
`SignPackageSignature(...)`, and `TrySignPackageSignature(...)` APIs expose the same result types used by Excel,
PowerPoint, and Visio. The established `ValidateSignatures(...)` API additionally retains Word-specific timestamp and
diagnostic evidence.

### VBA macro-project signatures

Signature parts in saved `.docm` and `.dotm` files can be inspected on every
supported platform without executing VBA:

```csharp
WordMacroProjectSignatureInfo signatures =
    OfficeIMO.OfficeDocument.Inspect("automation.docm", new OfficeIMO.OfficeInspectOptions {
        IncludeContent = false, WordMacroSignatures = new WordMacroProjectSignatureInspectionOptions()
    }).WordMacroSignatures!;
```

Managed VBA signing and content-binding validation work on every supported
platform. The workflow blocks existing OPC package signatures by default,
clears existing VBA signatures, creates and verifies the legacy, agile, and V3
profiles, proves that `vbaProject.bin` and the source package did not change
concurrently, and atomically replaces the package only after final validation.
When both signature kinds are needed, sign the VBA project first and the OPC
package last:

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
var options = new OfficeVbaSigningOptions();
options.CmsVerification.CertificateValidation.RevocationMode =
    X509RevocationMode.Online;

OfficeVbaSigningResult signing = WordDocument.SignVbaProject(
    "automation.docm",
    security,
    signingCertificate,
    options);

OfficeVbaSignatureValidationResult validation =
    WordDocument.ValidateVbaSignatures("automation.docm", security, options);
```

The caller supplies the certificate and `IOfficeSecurityProvider`; OfficeIMO
does not discover or persist private keys. Validation combines managed MS-OVBA
content binding with CMS signature, caller-controlled certificate-chain,
revocation, and RFC 3161 timestamp policy. Set
`ValidateWithWindowsSipWhenAvailable` only when a registered Microsoft Office
SIP should provide an additional differential check. Microsoft Office, SignTool,
and `offclearsig.exe` are not runtime dependencies. OfficeIMO does not execute VBA.

### VBA source editing

`ReadVbaProject()` returns a detached [native VBA model](../OfficeIMO.Core/README.md#native-vba-projects),
or `null` when the document has no project. Apply edits explicitly:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("automation.docm");
OfficeVbaProject project = document.ReadVbaProject()!;
project.SetModuleSource("Helpers",
    "Public Function Value() As Long\r\nValue = 73\r\nEnd Function\r\n");
document.SetVbaProject(project);
document.Save("updated.docm");
```

To create a project without a seed file, use `OfficeVbaProject.Create()`, add
standard/class modules, and pass it to `SetVbaProject`. Word supplies the required
`ThisDocument` identity and supplementary VBA data when absent. Existing
supplementary data and opaque designer content are preserved. Document classes
from another Office host are rejected. Source editing adds no runtime dependency.

Changes to a signed VBA project require
`new OfficeVbaWriteOptions { AllowSignatureRemoval = true }`; unchanged projects
retain signatures. Re-sign after editing. `Macros` reports logical directory
names for valid projects, and `RemoveMacro` uses the shared native editor for
standard/class modules. Host document/designer removal and protected-project
mutation are rejected.

### Content controls

```csharp
document.FillContentControlValues(new Dictionary<string, object?> {
    ["Name"] = "Ada Lovelace",
    ["Approved"] = true,
    ["DueDate"] = DateTime.Today
});

Dictionary<string, object?> values = document.ExtractContentControlValues();
document.ValidateContentControlValues(values).EnsureValid();
```

Form filling respects content-control locks. A supplied value for a control with
`contentLocked` or `sdtContentLocked` causes `FillContentControlValues` to throw
`InvalidOperationException` before applying any form values. Validation reports
`WordContentControlFormIssueKind.LockedControl` for the same target. Omit locked
fields from the map when filling the editable fields of a form.

An `sdtLocked` control remains fillable because that lock only prevents deleting
the control. Replacing a picture or repeating section also rejects locked nested
controls that would be removed. Direct control setters remain authoring APIs for
intentional document edits.

### Legacy DOC files

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;

using WordDocument document = WordDocument.Load("legacy-input.doc");
document.Save("converted-output.docx");

LegacyDocWriteAssessment assessment = document.AssessLegacyDocWrite();
if (assessment.IsSupported) {
    document.Save("legacy-output.doc");
}

OfficeIMO.OfficeConverter.Convert("legacy-input.doc", "converted-output.docx");
OfficeIMO.OfficeConverter.Convert("openxml-input.docx", "legacy-output.doc");

using WordDocument imported = WordDocument.Load("legacy-input.doc", new WordLoadOptions { FailOnLoss = true },
    out OfficeIMO.OfficeConversionReport importReport);
imported.Save("converted-output.docx");
foreach (OfficeIMO.OfficeDiagnostic diagnostic in importReport.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Legacy `.doc` support is first-party and dependency-free at runtime. The reader
projects supported paragraphs, formatting, styles, tables, sections, headers,
footers, bookmarks, fields, notes, comments, revisions, and inline pictures into
the normal `WordDocument` model. Native writing preflights the destination subset;
supported inline pictures retain inset crops and can use `OptimizeImages` before
saving. Default fonts and Normal style formatting are materialized in the native
stylesheet.

Floating drawings, unsupported image effects and transforms, embedded objects,
and other unprojected legacy content are diagnosed or blocked. File optimization
workflows refuse to publish incomplete legacy projections. Use
`OfficeConverter.Convert` from the `OfficeIMO` package for file-to-file conversion;
set `OfficeConversionOptions.FailOnLoss = true` to reject reported fidelity loss.
When saving an inspected document directly, set `WordSaveOptions.AllowLoss = true`
only after reviewing that loss.
See [DOC and DOCX compatibility](../Docs/officeimo.word.legacy-doc-compatibility.md)
for the current capability matrix and safety contract. Use the
[migration guide](../MIGRATION.md#legacy-doc-and-xls-api-changes) for canonical API replacements.

### Import additional legacy word-processing formats

The `OfficeIMO.Word` package also contains an explicit, read-only importer for selected WordPerfect, WordStar, Ami Pro, Lotus Word Pro, Microsoft Works/Write, and Word for DOS sources. No additional package is required, and these formats are only processed when the application calls `LegacyWordImporter` or explicitly registers the corresponding Reader handler.

```csharp
using OfficeIMO;
using OfficeIMO.Word.Legacy;

using LegacyWordImportResult imported = LegacyWordImporter.Import("archive.wpd");
Console.WriteLine(imported.Report.Quality);
foreach (LegacyWordSectionContent section in imported.Content.Sections) {
    foreach (LegacyWordBlockContent block in section.Blocks) {
        if (block is LegacyWordParagraphContent paragraph) Console.WriteLine(paragraph.Text);
        if (block is LegacyWordTableContent table) Console.WriteLine($"Table: {table.Rows.Count} rows");
    }
}
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Value.Save("archive.docx");
```

The importer never saves back to these source formats, executes macros or embedded code, activates embedded objects, or resolves external links. Each result identifies structured or salvage recovery and reports feature-level loss. The source-oriented `Content.Sections` retains ordered paragraphs and tables, page geometry, and running stories beside the projected `WordDocument`. Formatted runs expose anchored notes through `NoteIndex` and qualified graphics through `Image`; existing Word converter packages can export the document to ODT, HTML, Markdown, or PDF.

#### Profile coverage

| Family/profile | Quality | Recovered today | Explicit boundary |
| --- | --- | --- | --- |
| WordStar 3-7 character streams | Structured | hard and soft returns, paragraphs, common inline formatting, page breaks, selected dot commands, bounded notes/comments, paragraph-style names, and inert graphics references | printer/font/color/style-library sequences and unrecognized dot commands are reported; text-marker lists are identified as inferred |
| Ami Pro SAM 4 | Structured | style definitions, paragraphs, basic character styles, fonts, RGB color, alignment, spacing, page-break and keep properties, and source style names | the current structured profile is ASCII-only; code pages, frames, equations, images, tables, and additional inline tags remain open |
| Weak WordStar or non-SAM4 Ami Pro input with an explicit hint | Salvage | bounded text and paragraphs | a hint selects the family but does not upgrade weak input to structured quality |
| WordPerfect 5.x records | Structured | prefix directories and fonts, paragraph and character records, basic formatting, rectangular tables and widths, anchored footnotes/endnotes, headers/footers, page breaks and explicit geometry | merged/nested tables and tables inside running stories are rejected; source style libraries, fields, advanced cell formatting, and graphics are reported as loss |
| WordPerfect 6.x records | Structured | typed prefix packets and fonts, paragraphs, basic formatting and RGB color, rectangular tables and widths, anchored notes, headers/footers, page breaks and explicit geometry; basic cached WPG1 vectors as inline PNG images | merged/nested tables are rejected; style libraries, lists, fields, other WPG profiles, and floating box layout are reported as loss; WPG1 lines, polylines, polygons, rectangles and unrotated ellipses use solid fills/strokes and an explicit palette or the first sixteen default colors |
| Weak or other WordPerfect profiles | Salvage | bounded document-area text, paragraphs, offsets, and active-content marker inventory | weak input does not gain structured quality from its extension or a format hint |
| Lotus Word Pro LWP | Salvage | bounded text plus compound-content safety inventory | document zones, styles, notes, tables, graphics, and layout are not yet reconstructed |
| Microsoft Works word 2-8 | Salvage | bounded text and paragraphs plus compound-content safety inventory where applicable | formatting, fields, notes, tables, images, and layout are not yet reconstructed |
| Microsoft Write WRI | Salvage | bounded text and paragraph runs | formatting runs, objects, headers, footers, and layout are not yet reconstructed |
| Microsoft Word for DOS 4-6 | Salvage | bounded text and paragraphs | formatting, annotations, objects, and layout are not yet reconstructed |

`Structured` means the input passed the documented profile grammar, not that conversion is lossless. Inspect `Report.Findings`, or call `imported.RequireNoLoss()` when salvage recovery, inert content, or any known approximation must fail the workflow. Detection combines stable signatures and validated grammar with an optional source name; resource limits and cancellation apply before and during parsing.

WordPerfect 5 and 6 use separate record grammars. Encrypted documents and malformed record gates, references, or table grids fail import. Deleted WordPerfect 6 undo text is excluded. Basic WPG1 vectors are rasterized at 96 DPI; `Image.GetSourceBytes()` and `Image.GetPngBytes()` return independent copies. `Limits.MaxResourceBytes` bounds graphics bytes and `Limits.MaxImagePixels` bounds total generated pixels across body and referenced stories. The five attributed WordPerfect fixtures and WPG1 fixture in `OfficeIMO.LegacyImport.Tests/Fixtures/WordPerfect` provide independent source evidence; specification-built tests cover tables, notes, malformed input, and shared budgets. Complete source-format preservation and pixel-identical layout are outside these profiles.

### Protection

```csharp
using DocumentFormat.OpenXml.Wordprocessing;

document.Settings.ProtectionPassword = "owner-password";
document.Settings.ProtectionType = WordDocumentProtectionType.ReadOnly;
```

### Editable equations from the shared math model

```csharp
using OfficeIMO.Drawing;

OfficeMathExpression equation = OfficeMath.Fraction(
    OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")),
    OfficeMath.Number("2"));

WordEquation wordEquation = document.AddEquation(equation);
wordEquation.Paragraph.AddText(" is editable Word math.");
```

`WordDocument.AddEquation(...)` and `WordParagraph.AddEquation(...)` map the shared expression directly to native OMML. Existing equations expose `ToExpression()`, `SetExpression(...)`, and `ToDrawing(...)`; `WordMathMarkup` converts between OMML and `OfficeMathExpression`. The adapter covers matrices and multi-column equation arrays, left/right scripts, centered limits, skewed fractions, delimiter lists, n-ary operators, and decorations. Display-equation replacement retains `oMathParaPr` presentation metadata. Shared `Stack` and `StretchStack` nodes fail closed because OMML has no lossless equivalent; use `OfficeMath.EquationArray(...)` explicitly if that alternate layout is acceptable. The reusable AST stays in `OfficeIMO.Drawing`, while Word owns only the OMML adapter.

Authored operator form and side-spacing metadata retained from MathML also fail
with `NotSupportedException`; editable OMML does not preserve that presentation
contract. Keep the MathML and use the shared renderer when those values must
survive, or explicitly choose an editable projection without them.
`WordDocument.AddEquation(...)` checks the expression before adding a paragraph.

### Convert Word content

PDF export lives in `OfficeIMO.Word.Pdf` and uses the first-party managed PDF engine. HTML and Markdown conversion use their adjacent adapter packages.

```csharp
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

string html = document.ToHtml(new WordHtmlExportOptions { IncludeDefaultCss = true });
using var markdownDocument = document.ToMarkdown(new WordMarkdownExportOptions());
string markdown = markdownDocument.ToMarkdown();
document.SaveAsPdf("report.pdf");
```

## Editable charts from shared data

Use `WordDocument.AddChart(OfficeChartType, OfficeChartData, ...)` or the corresponding
`WordParagraph.AddChart(...)` overload to create a native chart
with an embedded Excel worksheet. `WordChart.SetData(kind, data)` updates its caches and
worksheet together, preserving the drawing dimensions, title, name, and alternative text.
Existing embedded packages must be XLSX; updates reject other package formats before changing
native data or package bytes.
External or unresolved workbook links also reject the update; embed an XLSX
workbook before replacing chart data.
Formula-linked titles, axis titles and custom labels retain their cached text as
native rich text; custom error bars retain cached numeric values as literals.
Updates reject uncached bindings and unqualified workbook-linked extensions
before changing the chart or worksheet.
Native axes retain their referenced identity, and compatible repeated layers
retain separate formatting when they share an axis pair. Repeated layers of
the same family and axis group with different axis pairs reject shared updates.
The shared writer supports column and bar grouping variants, line and area grouping variants,
pie, doughnut, radar, scatter, and bubble charts. Supported category combinations use each
series' `RenderKind` and `AxisGroup`; scatter, bubble, horizontal bars, pie, doughnut, and radar
have the combination restrictions enforced by the shared chart contract.
For scatter and bubble charts, a series' `RenderKind` must match the chart kind
or be omitted.

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create();
var data = new OfficeChartData(new[] { "Pass", "Could not evaluate", "Fail" }, new[] {
    new OfficeSeries("Status", new[] { 8d, 2d, 1d }).WithPointStyles(new OfficeChartPointStyle?[] {
        new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#008000")),
        new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2),
        new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#C00000"))
    })
});
WordChart chart = document.AddChart(OfficeChartType.Doughnut, data, new OfficeChartOptions {
    Title = "Status",
    Width = OfficeLength.Pixels(360),
    Height = OfficeLength.Pixels(240)
});
chart.Name = "Status distribution";
chart.SetRadialLayout(new OfficeChartRadialLayout(firstSliceAngleDegrees: 90, doughnutHolePercent: 70));
chart.AltText = "Eight passed, two could not be evaluated, and one failed.";
document.Save("status.docx");
```

Omitting point styles preserves existing native overrides during a data update; supply an
explicit array containing null entries to clear those points' fill and outline overrides.
`chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot)` reads supported two-dimensional
charts into the shared Drawing contract, including numeric X coordinates, bubble sizes,
category combinations, secondary-axis assignments, native plotting order, and point styles.
Word page images and Markdown SVG fallbacks use this projection. PDF uses it for bubble and
combination charts; its existing single-family route retains its separate cached-data limits.
Secondary value axes retain their own linear bounds, major/minor tick units and appearance, numeric label format, and title.
Use `chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 1,
majorUnit: 0.2, numberFormat: "0%").WithTitle("Completion rate"))` on a chart with a secondary series.
Independent secondary colours, line styles and gridlines, and logarithmic or reversed scales still reject this snapshot.

Snapshots preserve chart and plot surfaces, primary axes and gridlines, radial geometry,
uniform body fonts, separate chart-title and axis-title fonts, uniform text size, style and
colour for titles, legends and axes, and basic data-label content,
separator, number format, and position. Mixed body fonts, per-point or series-specific labels,
best-fit radial leader lines, unsupported radial label positions, unsupported effects, and unrepresented outlines return false
without changing the native document. Native authoring remains available for these charts.
Outside-end radial leader lines are supported when they can be placed without crossing another doughnut ring.
Deleted axes suppress their lines, and primary tick-label placement and automatic/minimum/maximum
axis crossing are preserved where the renderer supports them; automatic zero crossing through a negative
value range and maximum crossing on horizontal bars are rejected. Nonstandard axis arrangements,
title shape fills or unsupported legend fills,
manual title layouts, unsupported numeric formats, and nondefault bar spacing
require a richer projection and return false. Analytical overlays, sparse, empty, unequal or nonnumeric value caches,
time-scaled date axes, inverted negative bars, unresolved native style presets, per-entry legend text,
multiline, rotated or aligned text layouts, chart data tables, hierarchical categories, visible secondary category axes,
category-label skipping, offsets or non-centered alignment, rounded chart frames, non-box bar shapes,
bar connector lines, chart drawing overlays, nondefault cross-between geometry, and nonidentity Word color-scheme mappings also reject projection. Visible-only
charts reject hidden rows or columns in their referenced workbook ranges; unrelated hidden cells do not affect the snapshot. Literal charts remain supported.
Pie and doughnut snapshots preserve supported per-slice explosion offsets in native charts and static exports.
Header and footer charts use relationships owned
by their containing story, including when the same relationship ID exists in the document body.

Category and scatter snapshots retain supported native line and marker appearance.
Call `snapshot.Data.Series[index].ToOfficeSeries()` to obtain the shared series,
including connecting-line visibility, stroke width and dash, marker shape and
size, marker outlines, and point overrides. Word page images and Markdown chart
drawings use that same series. Unqualified native dashes, curved lines, distinct
marker and connecting-line colours, unresolved automatic marker symbols, or unfilled marker treatments reject the
managed snapshot rather than changing their appearance. Cached projections are
bounded to 10,000 positions per Word cache and 100,000 positions across a chart.
Point overrides use a separate 1,000,000-record limit that counts stale and duplicate
records. Native series retain their plotting order. Gradients, custom or compound
outlines, and per-point marker overrides reject static projection. Supported filled
series outlines are inherited by points; explicit point outlines take precedence.
Picture markers reject static projection. Visible inherited markers require an explicit supported series symbol.
Small authored chart canvases retain explicit font, stroke and marker sizes; quality reports identify cramped or overflowing content.

Category discovery checks all populated series before generating fallback labels.
The longest available category cache supplies labels, and shorter series retain
their positions with zero padding. Marker-only plots use the marker fill; a
marker fill cannot replace an unresolved connecting-line colour in a static export.

`RadialLayout` reads native pie rotation and doughnut hole size. `SetRadialLayout(...)`
updates an existing two-dimensional pie or doughnut chart. These settings survive
save/reopen, data updates, snapshots, managed images, PDF, and chart projections to Markdown.
Rotation is clockwise from the top (0–360 degrees); hole size is 10–90 percent.

## Native doughnut charts

`WordChart.AddDoughnut(category, value)` creates an editable native doughnut chart with a
50-percent hole. It accepts finite, nonnegative `int`, `double`, or `float` values and can
append slices after reopening a chart authored with literal data, preserving gaps and disabled
labels. Numeric category literals accept finite invariant numeric category strings. Linked worksheet caches
and multi-ring imported doughnuts use the existing cached-data mutation APIs instead.

```csharp
WordChart chart = document.AddChart("Status", roundedCorners: false, width: OfficeLength.Pixels(360), height: OfficeLength.Pixels(180));
chart.AddDoughnut("Pass", 8).AddDoughnut("Could not evaluate", 2).AddDoughnut("Fail", 1);
chart.SetDataPointStyle(0, 1,
    new OfficeIMO.Drawing.OfficeChartPointStyle(noFill: true,
        outlineColor: OfficeIMO.OfficeColor.Black, outlineWidth: 2));
```

## Individual chart point styles

Use `chart.SetDataPointStyle(seriesIndex, pointIndex, style)` to apply an
`OfficeIMO.Drawing.OfficeChartPointStyle` to a native Word chart. Solid fill, explicit no-fill,
outlines with optional joins, and seven hatch patterns are supported without changing
the chart values.
Passing null clears the point's fill and outline overrides.

Native save/reopen, chart snapshots, managed images, PDF export, and Markdown SVG chart
fallbacks carry supported point styles. Pie legend swatches follow their slices.
Static area charts use the opaque series fill, honor an explicit no-outline series,
and report unsupported per-point styling. Imported native point overrides remain
in the editable chart package.
See the [shared style example](../OfficeIMO.Core/README.md#style-individual-chart-points).

## SmartArt in headers and footers

Create diagrams through a paragraph in the containing story:

```csharp
WordSmartArt diagram = document.HeaderDefaultOrCreate.AddParagraph()
    .AddSmartArt(WordSmartArtType.BasicProcess);
diagram.SetNodeText(0, "Review");
WordSmartArtNode approval = diagram.AddNode("Approve");
approval.Text = "Approve and publish";
```

Use `FooterDefaultOrCreate` for a footer. `Nodes` exposes attached nodes in diagram order; a node retains its identity when moved and its text remains editable after reloading. DOCX diagram data, layout, colours,
quick style and any cached drawing use relationships owned by that header or
footer. Reopening and editing resolve those local relationships independently
of diagrams in the body, even when relationship IDs are reused.

PDF and page images repeat the diagram with the containing story. Supported
cached preset geometry is used when present; unsupported appearance still
carries the [SmartArt projection limits](../OfficeIMO.Word.Pdf/README.md#page-images-with-pdf-pagination).

## Managed image export

Word pages use the shared Word-to-PDF layout and Drawing renderer and can be returned as PNG, JPEG, TIFF, lossless WebP, or SVG without Office automation:

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;

byte[] webp = document.ToImage(new WordImageExportOptions {
    Format = OfficeImageExportFormat.Webp, PageIndex = 0, Scale = 1.5
});

var options = new WordImageExportOptions {
    Format = OfficeImageExportFormat.Jpeg,
    PageIndex = 0,
    MaximumOutputWidth = 1600,
    MaximumOutputHeight = 1200
};
options.RasterEncoding.Jpeg.Quality = 90;
OfficeConversionReport report = document.SaveAsImage("page-1.jpg", options);
```

Image export shares Word-to-PDF pagination, line breaks, fonts, headings and table borders. Paragraphs and cells flow across pages without preview ellipsis. `WordImageExportOptions.LayoutOptions` accepts the same `WordPdfExportOptions` as PDF conversion. For PDF and images from one conversion, use the [page-image route](../OfficeIMO.Word.Pdf/README.md#page-images-with-pdf-pagination).

`OfficeIMO.Word` owns the Word page adapter and references the first-party managed `OfficeIMO.Pdf` layout and page renderer. No Office installation or external conversion executable is required. `CreateVisualSnapshots()` exposes the resulting Drawing scenes and `TextSpans` with page text geometry.

The document package owns Word pagination and diagnostics; `OfficeIMO.Drawing` owns sizing, pixels, and encoding. The same fit limit applies to SVG and raster output. `SaveAsJpeg`, `SaveAsTiff`, and `SaveAsWebp` are thin convenience wrappers over the same builder.

## Content provenance

Add the `OfficeIMO` front-door package for the shared inspection entry point.

Inspect C2PA and AI-specific IPTC metadata in the package and its supported embedded images, then remove only the selected carriers:

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Word;

OfficeProvenanceReport report = OfficeIMO.OfficeDocument.Inspect("input.docx", new OfficeIMO.OfficeInspectOptions { IncludeContent = false, Provenance = new OfficeProvenanceOptions() }).Provenance!;
OfficeProvenanceRemovalResult result = WordDocument.RemoveProvenance("input.docx", "clean.docx");
```

Mutation of a signed package is blocked by default. Set `SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` only when removing the now-invalid package signature is intentional. Optional cryptographic C2PA verification is provided by `OfficeIMO.Security`.

When the document is already in memory, inspect its encoded package bytes directly. This overload validates the package and inspects supported provenance without accessing the filesystem:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Provenance;

static OfficeProvenanceReport InspectUploadedWord(byte[] packageBytes) =>
    OfficeIMO.OfficeDocument.Inspect(new MemoryStream(packageBytes), "upload.docx",
        new OfficeIMO.OfficeInspectOptions { IncludeContent = false, Provenance = new OfficeProvenanceOptions() }).Provenance!;
```

## Concealed-content inspection and cleanup

`OfficeDocument.Inspect(path, new OfficeInspectOptions { ContentSafety = new OfficeContentSafetyOptions() }).ContentSafety` reports native/inherited hidden runs, deleted revisions, tiny or zero-geometry text, explicit low contrast, comments/notes, alternative text, and exact Unicode evidence. Pass reviewed finding IDs to `WordDocument.RemoveSelectedContent(...)`; the package is reopened after cleanup, and signed-document mutation fails closed by default. These findings describe ingestion risk, not AI authorship.

## Adjacent packages

`OfficeIMO.Word` owns the Word model. Conversion and export packages stay separate so consumers only take the dependencies they need:

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Word.Html](../OfficeIMO.Word.Html/README.md) | Word to/from HTML conversion. |
| [OfficeIMO.Word.Markdown](../OfficeIMO.Word.Markdown/README.md) | Word to/from Markdown conversion. |
| [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md) | Word to PDF export through `OfficeIMO.Pdf`. |
| [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md) | Planning and exporting Word content to Google Docs. |

## Related packages

- Use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for PowerShell examples and cmdlets.
- Use `OfficeIMO.Word.Pdf` for Word-to-PDF conversion and `OfficeIMO.Pdf` for direct PDF layout and manipulation.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

Runnable samples live under [OfficeIMO.Examples/Word](../OfficeIMO.Examples/Word).

## Dependency footprint

- **External:** Open XML SDK for `.docx` package mechanics. Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core`. The object model, native OMML adapter, legacy `.doc` reader/writer, lifecycle, validation, and PNG/JPEG/TIFF/WebP/SVG export are first-party.
- **Optional security:** install `OfficeIMO.Security` and pass `OfficeSecurityProvider.Default` only for OPC/VBA signing or cryptographic validation. It is not a transitive Word dependency.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 14 | 12 | 0 | 8 | 2 | 2 |
| Read | 24 | 0 | 9 | 0 | 2 | 3 |
| Edit | 12 | 0 | 22 | 3 | 0 | 1 |
| Preserve | 11 | 0 | 22 | 0 | 0 | 1 |
| Inspect | 5 | 1 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 31 | 26 | 0 | 8 | 0 | 1 |

The complete rows for `OfficeIMO.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
