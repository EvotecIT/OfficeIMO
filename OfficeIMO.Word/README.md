# OfficeIMO.Word - Word documents for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

`OfficeIMO.Word` is the main Word document package in the OfficeIMO family. It creates, edits, inspects, converts, and saves `.docx` files, and can import and write supported legacy `.doc` files, without COM automation and without Microsoft Office installed.

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
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");

document.AddParagraph("Quarterly report").Style = WordParagraphStyles.Heading1;
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

`AsFluent()` wraps the same `WordDocument`; `End()` returns that document for
direct object-model work:

```csharp
using var document = WordDocument.Create("report.docx");

document.AsFluent()
    .H1("Quarterly report")
    .Paragraph(paragraph => paragraph.Text("Created with OfficeIMO.Word."))
    .End();

document.Save();
```

## What it does

- Creates, loads, edits, saves, and appends `.docx` documents.
- Opens supported Word 97-2003 `.doc` files through the normal `WordDocument.Load(...)` path and projects them into the regular OfficeIMO Word model.
- Writes native `.doc` files for the currently supported simple-document subset, with preflight checks that block unsupported content before saving.
- Converts supported `.doc` and `.docx` files with `WordDocument.Convert(...)`, using the same import diagnostics and save preflight as normal load/save workflows.
- Works with paragraphs, runs, styles, sections, headers, footers, page numbers, tables, images, hyperlinks, bookmarks, fields, footnotes, endnotes, content controls, charts, shapes, and document protection.
- Applies optional shared package-security policy before parsing Open XML or compound DOC files, and preflights read, edit, template, render, and save capabilities.
- Inspects and manages VBA and embedded package/OLE/ActiveX payload bytes without executing active content.
- Creates and validates cross-platform OPC XML package signatures and managed VBA legacy, agile, and V3 signatures through an explicitly supplied `IOfficeSecurityProvider`; structural inspection remains provider-free.
- Exports estimated document page ranges as dependency-free PNG or SVG previews through `ExportImages(...)`, `SaveAsImages(...)`, and `ToImages()`.
- Keeps Office automation out of the runtime path, making it suitable for services, scheduled jobs, CI, desktop apps, and automation hosts.
- Provides fluent helpers for common authoring flows while keeping the lower-level Word object model available.
- Uses `OfficeIMO.Drawing` for shared colors, image metadata, page rendering, and the reusable math expression tree.

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

### Tables with structure

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
paragraph.AddImage("logo.png", width: 160, height: 64);
```

Use `paragraph.Image.Clone(destinationParagraph)` to copy an image to another paragraph,
header, footer, or document. The destination owns its image relationships and fresh DrawingML
identifiers. Cloning preserves VML shapes and their referenced definitions, and remaps
embedded SVG extension resources along with the primary image.

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
document.Paragraphs[0].AddField(WordFieldType.TOC);
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
    .SetFormat(WordFieldFormat.Numeric);
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
certificate trust. The cross-host `InspectPackageSignatures(...)`, `ValidatePackageSignatures(...)`,
`SignPackageSignature(...)`, and `TrySignPackageSignature(...)` APIs expose the same result types used by Excel,
PowerPoint, and Visio. The established `ValidateSignatures(...)` API additionally retains Word-specific timestamp and
diagnostic evidence.

### VBA macro-project signatures

Signature parts in saved `.docm` and `.dotm` files can be inspected on every
supported platform without executing VBA:

```csharp
WordMacroProjectSignatureInfo signatures =
    WordDocument.InspectMacroProjectSignatures("automation.docm");
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
and `offclearsig.exe` are not runtime dependencies. OfficeIMO does not execute
VBA or edit VBA source modules.

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

WordDocument.Convert("legacy-input.doc", "converted-output.docx");
WordDocument.Convert("openxml-input.docx", "legacy-output.doc");

using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport("legacy-input.doc");
if (result.HasDocument) {
    result.EnsureNoConversionLoss();
    result.Document.Save("converted-output.docx");
    string report = result.CreateAdvancedImportReport().ToMarkdown();
}
```

Legacy `.doc` support is first-party and dependency-free at runtime. The current
reader projects supported Word 97-2003 body paragraphs, simple zero-length,
same-paragraph, and cross-paragraph body bookmarks plus simple table-cell,
header/footer, and footnote/endnote paragraph bookmarks, simple external and
internal bookmark hyperlink fields with supported text, tab, soft/no-break
hyphen, and break display runs, simple static date/time and document-property
field display results,
common run and paragraph formatting including proofing exclusion, bidirectional
paragraph layout, mirror indents, contextual spacing, East Asian typography and
punctuation spacing flags, and automatic hyphenation suppression, built-in and
custom paragraph styles, simple tables, paragraph-boundary sections, page setup,
simple header/footer stories with tabs, text-wrapping and column breaks,
supported direct run formatting, and supported paragraph formatting, simple
footnote/endnote bodies with supported direct run and paragraph formatting and
soft/no-break hyphen runs, section note numbering and placement settings, and
document properties into the normal `WordDocument` model. Native `.doc` saving
is available for the supported simple subset: paragraphs, simple zero-length,
same-paragraph, and cross-paragraph body bookmarks plus simple table-cell,
header/footer, and footnote/endnote paragraph bookmarks, simple external and
internal bookmark hyperlinks with supported text, tab, soft/no-break hyphen,
break display runs, simple static date/time and document-property fields with
static display text and supported inline result characters including inside flattened inline content
controls, simple inline content-control display text, and simple block content
controls with nested
simple block controls plus nested inline content controls in
body/table/header/footer/footnote/endnote stories, common run and paragraph
formatting including proofing exclusion,
bidirectional paragraph layout, mirror indents, contextual spacing, East Asian
typography and punctuation spacing flags, and automatic hyphenation suppression,
tabs, soft/no-break hyphen runs, line/carriage-return/page/column breaks, simple
body tables with common formatting, including simple depth-2 nested tables,
supported table-style border, shading, layout, paragraph
formatting, run formatting, default-cell expansion, conditional table/cell
border, shading, paragraph formatting, run formatting, cell-layout expansion,
and conditional row height/header/no-split formatting, paragraph-boundary
sections, page setup, simple header/footer stories with tabs, soft/no-break
hyphen runs, text-wrapping, carriage-return, and column breaks, supported
direct run formatting, and supported paragraph formatting,
simple footnote/endnote bodies with supported direct run and paragraph
formatting and soft/no-break hyphen runs, supported section note settings, and
scalar document properties. Unsupported features such as macros, embedded OLE
objects, comments, text boxes, images, bookmark ranges outside supported
body/table-cell/header/footer/footnote/endnote paragraphs, richer
content-control children, richer visual table style effects, deeper or richer
nested table shapes,
richer note body structures, and richer header/footer or section shapes are
diagnosed or blocked rather than silently flattened. `WordDocument.Convert(...)`
uses those same load and save paths and blocks legacy sources with unsupported
or preserve-only content by default. Set `LossPolicy` to
`OfficeConversionLossPolicy.Allow` on `WordDocumentConversionOptions` or
`WordSaveOptions` only when that loss has been reviewed and is intentional.
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
foreach (LegacyWordParagraphContent paragraph in imported.Content.Paragraphs) {
    Console.WriteLine($"{paragraph.StyleName}: {paragraph.Text}");
}
foreach (OfficeCompatibilityFinding finding in imported.Report.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

imported.Value.Save("archive.docx");
```

The importer never saves back to these source formats, executes macros or embedded code, activates embedded objects, or resolves external links. Each result identifies structured or salvage recovery and reports feature-level loss. The source-oriented `Content` retains paragraphs, formatted runs, notes, and inert resource references beside the projected `WordDocument`; existing Word converter packages can export that document to ODT, HTML, Markdown, or PDF.

#### Profile coverage

| Family/profile | Quality | Recovered today | Explicit boundary |
| --- | --- | --- | --- |
| WordStar 3-7 character streams | Structured | hard and soft returns, paragraphs, common inline formatting, page breaks, selected dot commands, bounded notes/comments, paragraph-style names, and inert graphics references | printer/font/color/style-library sequences and unrecognized dot commands are reported; text-marker lists are identified as inferred |
| Ami Pro SAM 4 | Structured | style definitions, paragraphs, basic character styles, fonts, RGB color, alignment, spacing, page-break and keep properties, and source style names | the current structured profile is ASCII-only; code pages, frames, equations, images, tables, and additional inline tags remain open |
| Weak WordStar or non-SAM4 Ami Pro input with an explicit hint | Salvage | bounded text and paragraphs | a hint selects the family but does not upgrade weak input to structured quality |
| WordPerfect 5/6 | Salvage | bounded document-area text, paragraphs, offsets, and active-content marker inventory | prefix packets, formatting codes, notes, tables, graphics, and layout are not yet semantically decoded |
| Lotus Word Pro LWP | Salvage | bounded text plus compound-content safety inventory | document zones, styles, notes, tables, graphics, and layout are not yet reconstructed |
| Microsoft Works word 2-8 | Salvage | bounded text and paragraphs plus compound-content safety inventory where applicable | formatting, fields, notes, tables, images, and layout are not yet reconstructed |
| Microsoft Write WRI | Salvage | bounded text and paragraph runs | formatting runs, objects, headers, footers, and layout are not yet reconstructed |
| Microsoft Word for DOS 4-6 | Salvage | bounded text and paragraphs | formatting, annotations, objects, and layout are not yet reconstructed |

`Structured` means the input passed the documented profile grammar, not that conversion is lossless. Inspect `Report.Findings`, or call `imported.RequireNoLoss()` when salvage recovery, inert content, or any known approximation must fail the workflow. Detection combines stable signatures and validated grammar with an optional source name; resource limits and cancellation apply before and during parsing.

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

WordParagraph paragraph = document.AddEquation(equation);
paragraph.AddText(" is editable Word math.");
```

`WordDocument.AddEquation(...)` and `WordParagraph.AddEquation(...)` map the shared expression directly to native OMML. Existing equations expose `ToExpression()`, `SetExpression(...)`, and `ToDrawing(...)`; `WordMathMarkup` converts between OMML and `OfficeMathExpression`. The adapter covers matrices and multi-column equation arrays, left/right scripts, centered limits, skewed fractions, delimiter lists, n-ary operators, and decorations. Display-equation replacement retains `oMathParaPr` presentation metadata. Shared `Stack` and `StretchStack` nodes fail closed because OMML has no lossless equivalent; use `OfficeMath.EquationArray(...)` explicitly if that alternate layout is acceptable. The reusable AST stays in `OfficeIMO.Drawing`, while Word owns only the OMML adapter.

### Convert with adjacent packages

```csharp
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

string html = document.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true });
string markdown = document.ToMarkdown(new WordToMarkdownOptions());
document.SaveAsPdf("report.pdf");
```

## Editable charts from shared data

Use `WordDocument.AddChart(OfficeChartKind, OfficeChartData, ...)` or the corresponding
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
using OfficeIMO.Drawing;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create();
var data = new OfficeChartData(new[] { "Pass", "Could not evaluate", "Fail" }, new[] {
    new OfficeChartSeries("Status", new[] { 8d, 2d, 1d }).WithPointStyles(new OfficeChartPointStyle?[] {
        new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#008000")),
        new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2),
        new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#C00000"))
    })
});
WordChart chart = document.AddChart(OfficeChartKind.Doughnut, data, "Status", width: 360, height: 240);
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
Secondary value axes retain their own linear bounds, major/minor tick units and appearance, and numeric label format.
Use `chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 1,
majorUnit: 0.2, numberFormat: "0%"))` on a chart with a secondary series. Secondary titles,
independent colours, line styles and gridlines, and logarithmic or reversed scales still reject this snapshot.

Snapshots preserve chart and plot surfaces, primary axes and gridlines, radial geometry,
uniform body fonts, separate chart-title and axis-title fonts, uniform text size, style and
colour for titles, legends and axes, and basic data-label content,
separator, number format, and position. Mixed body fonts, per-point or series-specific labels,
outside or best-fit radial leader lines, unsupported effects, and unrepresented outlines return false
without changing the native document. Native authoring remains available for these charts.
Deleted axes suppress their lines, and primary tick-label placement and automatic/minimum/maximum
axis crossing are preserved where the renderer supports them; automatic zero crossing through a negative
value range and maximum crossing on horizontal bars are rejected. Nonstandard axis arrangements,
title shape fills or unsupported legend fills,
manual title layouts, unsupported numeric formats, exploded slices, and nondefault bar spacing
require a richer projection and return false. Analytical overlays, sparse, empty, unequal or nonnumeric value caches,
time-scaled date axes, inverted negative bars, unresolved native style presets, per-entry legend text,
multiline, rotated or aligned text layouts, chart data tables, hierarchical categories, visible secondary category axes,
category-label skipping, offsets or non-centered alignment, rounded chart frames, non-box bar shapes,
bar connector lines, chart drawing overlays, nondefault cross-between geometry, and nonidentity Word color-scheme mappings also reject projection. Visible-only
charts reject hidden rows or columns in their referenced workbook ranges; unrelated hidden cells do not affect the snapshot. Literal charts remain supported.
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
WordChart chart = document.AddChart("Status", roundedCorners: false, width: 360, height: 180);
chart.AddDoughnut("Pass", 8).AddDoughnut("Could not evaluate", 2).AddDoughnut("Fail", 1);
chart.SetDataPointStyle(0, 1,
    new OfficeIMO.Drawing.OfficeChartPointStyle(noFill: true,
        outlineColor: OfficeIMO.Drawing.OfficeColor.Black, outlineWidth: 2));
```

## Individual chart point styles

Use `chart.SetDataPointStyle(seriesIndex, pointIndex, style)` to apply an
`OfficeIMO.Drawing.OfficeChartPointStyle` to a native Word chart. Solid fill, explicit no-fill,
outlines with optional joins, and seven hatch patterns are supported without changing
the chart values.
Passing null clears the point's fill and outline overrides.

Native save/reopen, chart snapshots, managed images, PDF export, and Markdown SVG chart
fallbacks carry supported point styles. Pie legend swatches follow their slices.
Static area charts use the series fill and report unsupported per-point styling.
See the [shared style example](../OfficeIMO.Core/README.md#style-individual-chart-points).

## Managed image export

Word page previews use the shared Drawing renderer and can be returned as PNG, JPEG, TIFF, lossless WebP, or SVG without Office automation:

```csharp
using OfficeIMO.Drawing;

byte[] webp = document.ToWebp(new WordImageExportOptions { PageIndex = 0, Scale = 1.5 });

document.ToImage()
    .Page(0)
    .FitWithin(1600, 1200)
    .AsJpeg()
    .WithRasterEncoding(raster => raster.Jpeg.Quality = 90)
    .Save("page-1.jpg");
```

The document package owns Word pagination and diagnostics; `OfficeIMO.Drawing` owns sizing, pixels, and encoding. The same fit limit applies to SVG and raster output. `SaveAsJpeg`, `SaveAsTiff`, and `SaveAsWebp` are thin convenience wrappers over the same builder.

## Content provenance

Inspect C2PA and AI-specific IPTC metadata in the package and its supported embedded images, then remove only the selected carriers:

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Word;

OfficeProvenanceReport report = WordDocument.InspectProvenance("input.docx");
OfficeProvenanceRemovalResult result = WordDocument.RemoveProvenance("input.docx", "clean.docx");
```

Mutation of a signed package is blocked by default. Set `SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` only when removing the now-invalid package signature is intentional. Optional cryptographic C2PA verification is provided by `OfficeIMO.Security`.

When the document is already in memory, inspect its encoded package bytes directly. This overload validates the package and inspects supported provenance without accessing the filesystem:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Provenance;

static OfficeProvenanceReport InspectUploadedWord(byte[] packageBytes) =>
    WordDocument.InspectProvenance(packageBytes, "upload.docx");
```

## Concealed-content inspection and cleanup

`WordDocument.InspectContentSafety(...)` reports native/inherited hidden runs, deleted revisions, tiny or zero-geometry text, explicit low contrast, comments/notes, alternative text, and exact Unicode evidence. Pass reviewed finding IDs to `WordDocument.RemoveSelectedContent(...)`; the package is reopened after cleanup, and signed-document mutation fails closed by default. These findings describe ingestion risk, not AI authorship.

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
- **OfficeIMO:** `OfficeIMO.Core`. The fluent model, native OMML adapter, legacy `.doc` reader/writer, lifecycle, validation, and PNG/JPEG/TIFF/WebP/SVG export are first-party.
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
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
