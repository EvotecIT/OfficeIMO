# OfficeIMO.OpenDocument

`OfficeIMO.OpenDocument` creates and edits ODT, ODS, ODP, and ODG files directly. Its only runtime dependency is the zero-dependency `OfficeIMO.Core` foundation used across OfficeIMO for lifecycle and result contracts. It has no third-party runtime dependencies and does not invoke LibreOffice, Microsoft Office, or UNO.

```powershell
dotnet add package OfficeIMO.OpenDocument
```

## Create documents

Create an ODT document:

```csharp
using OfficeIMO.OpenDocument;

OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");

OdtTable table = document.AddTable(2, 2, "Results");
table.Cell(0, 0).Text = "Metric";
table.Cell(0, 1).Text = "Value";
table.Cell(1, 0).Text = "Revenue";
table.Cell(1, 1).Text = "42";

document.Save("summary.odt");
```

Add a native ODT field in paragraph order with the value currently shown in the document:

```csharp
OdtParagraph pageLine = document.AddParagraph("Page ");
pageLine.AddField(OdtFieldKind.PageNumber, "3");
pageLine.AddText(" of ");
pageLine.AddField(OdtFieldKind.PageCount, "12");
document.Save("summary.odt");
```

`OdtParagraph.Fields` reads page number, page count, date, and time fields in order. `OdtField.DisplayText` edits the cached text; an office application can refresh a dynamic field later. `OdtField.IsFixed` retains a date, time, or page number value instead of refreshing it. Page count cannot be fixed.

Add native ODT footnotes or endnotes at the current paragraph position:

```csharp
using OfficeIMO.OpenDocument;

OdtDocument document = OdtDocument.Create();
OdtParagraph paragraph = document.AddParagraph("The result is documented");
paragraph.AddFootnote("Source and calculation details.");
paragraph.AddText(" in the appendix.");
paragraph.AddEndnote("Additional context.");
```

`OdtParagraph.Notes` and `InlineNodes` expose note bodies and reference order after reopening the file. Note-body paragraphs are separate from document-body paragraphs, and their spans, links, and images stay on the note-body paragraph. Native note citations follow document order when notes are added to earlier paragraphs later; an imported custom citation label remains the displayed `OdtNote.Citation`. Adding a note to a document with `text:notes-configuration` for that note kind throws `NotSupportedException`, preserving its configured numbering and existing citations.

Create a sparse ODS workbook:

```csharp
OdsDocument workbook = OdsDocument.Create();
OdsSheet sheet = workbook.AddSheet("Metrics");
sheet.Cell(0, 0).SetString("Name");
sheet.Cell(0, 1).SetString("Value");
sheet.Cell(1, 0).SetString("Revenue");
sheet.Cell(1, 1).SetDecimal(42.5m);

OdsCell total = sheet.Cell(2, 1);
total.Formula = "of:=SUM([.B2:.B2])";
OdsValidation positive = workbook.AddValidation(
    "PositiveAmount",
    OdsValidationConditionSyntax.Create(
        OdsValidationValueKind.DecimalNumber,
        OdsValidationComparison.GreaterThan,
        "0"));
positive.SetHelpMessage("Amount", "Enter a value greater than zero.");
positive.SetErrorMessage("Invalid amount", "The amount must be positive.");
sheet.Cell(1, 1).ValidationName = positive.Name;
OdsRecalculationReport calculation = workbook.Recalculate();
if (calculation.FailedCells > 0) {
    Console.WriteLine(calculation.Diagnostics[0].Message);
}

workbook.Save("metrics.ods");
```

For a formula-based validation, use `OdsValidationConditionSyntax.CreateFormula("[.B2]>0")` and set `OdsValidation.BaseCellAddress` to an absolute sheet-qualified address such as `$'Metrics'.$B$2`. The typed formula condition accepts a local cell compared with a number, text value, or another local cell, plus `ISBLANK`, `ISNUMBER`, or `ISTEXT` on one local cell. These predicates can be combined with `AND`, `OR`, and unary `NOT`; numbers in comparisons may have a leading sign. Other OpenFormula expressions can be stored in the raw `Condition` property. The base cell determines how relative formula references apply to the validated cells.

ODS conditional cell styles can be authored and edited through `OdfStyle.AddConditionalMap`. Create a common named table-cell style for the desired appearance, add a mapping to a base style, and assign the base style to cells. For example, `baseStyle.AddConditionalMap("cell-content()>0", highlightStyle.Name, "$'Metrics'.$B$2")` applies the highlight style when the condition is true. The applied style must be a common named style in the same family as the base style; `Validate()` reports a missing, automatic, or different-family target. The map and its relative-reference base survive save and reopen. OfficeIMO preserves these native rules; its renderer does not evaluate them.

ODS data pilot tables can be inspected through `OdsDocument.DataPilotTables`. Use `AddDataPilotTable("SalesPivot", "Metrics.A1:Metrics.B20", "Metrics.D1:Metrics.E22")` to author a local-range pivot, then add row, column, page, or data fields with `AddField`; a data field names its aggregation, such as `AddField("Value", "data", "sum")`. The target range describes the output area; the API does not calculate cached pivot results. Imported grouping, member selection, and other advanced settings remain in the package XML but are outside the typed authoring subset.

Create an ODP presentation:

```csharp
OdpPresentation presentation = OdpPresentation.Create();
OdpMasterPage master = presentation.AddMasterPage("Brand");
master.BackgroundColor = OdfColor.Parse("#F8FBFF");
OdpPresentationLayout layout = presentation.AddLayout("Title");
layout.AddPlaceholder("title", OdfRect.FromCentimeters(2, 1, 28, 3));
OdpSlide slide = presentation.AddSlide("Summary");
slide.MasterPageName = master.Name;
slide.LayoutName = layout.Name;
OdpTextBox title = slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 28, 3), "Native ODP");
title.PresentationClass = "title";
slide.AddRectangle(OdfRect.FromCentimeters(2, 5, 8, 3)).FillColor = OdfColor.Parse("#D1E9FF");
slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");
presentation.Save("summary.odp");
```

Nest text styles and links when their formatting changes within a sentence:

```csharp
using OfficeIMO.OpenDocument;

OdtDocument document = OdtDocument.Create();
OdtParagraph paragraph = document.AddParagraph();
OdtSpan emphasis = paragraph.AddSpan("Read ");
emphasis.Bold = true;
emphasis.AddHyperlink("the guide", "https://example.com/guide").Italic = true;

OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Links");
OdpParagraph slideText = slide.AddTextBox(
    OdfRect.FromCentimeters(2, 9, 18, 2)).AddParagraph();
OdpRun label = slideText.AddRun("Open ");
label.Bold = true;
label.AddHyperlink("the guide", "https://example.com/guide").Underline = true;
```

`OdtTableCell.Paragraphs` and `OdpTableCell.Paragraphs` read repeated cells without expanding them. Edit through the selected logical cell to change only that row and column; child spans, links, fields, and runs follow the same rule.

`InlineNodes` exposes nested `Children` in document order. A nested run inherits text properties from its containing span or link until its own style overrides them.

Convert explicitly between OpenDocument and OfficeIMO Word, Excel, or PowerPoint models by installing the corresponding adapter package. Every conversion returns an `OdfConversionReport` that identifies mapped, approximated, skipped, and unsupported features.

```powershell
dotnet add package OfficeIMO.Word.OpenDocument
dotnet add package OfficeIMO.Excel.OpenDocument
dotnet add package OfficeIMO.PowerPoint.OpenDocument
```

## Create and edit Draw documents

`OdgDocument` uses the same package, style, image, security, and preservation engine as the other OpenDocument models. It reads packaged `.odg` and flat `.fodg` drawings.

```csharp
using OfficeIMO.OpenDocument;

OdgDocument drawing = OdgDocument.Create();
OdgPage page = drawing.AddPage("Workflow",
    OdfLength.Centimeters(24), OdfLength.Centimeters(16));
OdgShape box = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 3, 6, 3));
box.Text = "Receive order";
box.FillColor = OdfColor.Parse("#D1E9FF");
box.FontSize = OdfLength.Points(18);
page.Shapes.AddEllipse(OdfRect.FromCentimeters(12, 3, 6, 3)).Text = "Check stock";
page.Shapes.AddLine(OdfLength.Centimeters(7), OdfLength.Centimeters(4.5),
    OdfLength.Centimeters(12), OdfLength.Centimeters(4.5));
drawing.Save("workflow.odg");
drawing.SaveFlatXml("workflow.fodg");

OdgDocument edited = OdgDocument.LoadFlatXml("workflow.fodg");
edited.Pages[0].Shapes[0].Text = "Order received";
edited.Save("workflow-edited.odg");
```

Pages and shape collections support insertion, removal, and reordering. Shapes expose bounds, names, XML IDs, solid fill/stroke, text, basic font properties, line endpoints, embedded image bytes, and nested group children. `ToXml()` returns a detached copy for inspection. Imported pages can share a layout, so changing their dimensions changes every page referencing that layout.

`OdgDocument.ToDrawings()` projects all pages in order with original dimensions and page-qualified `OdfConversionReport` losses. It accepts a page-count limit, print or screen layer intent, strict loss policies and cancellation. `OdgPage.ToDrawing()` remains available for an individual page. The [Draw PDF package](../OfficeIMO.OpenDocument.Odg.Pdf/README.md) uses this shared projection for one PDF page per source page. Text stored inside image frames remains an explicit omission in the scene profile.

### Native Draw text editing

`OdgShape.Paragraphs` exposes native paragraphs and headings in reading order, including list items. `Runs`, `Hyperlinks`, `Fields`, and `InlineNodes` let you inspect and edit individual portions of a paragraph. The text wrappers use the shared ODF whitespace and style owners; they resolve inline, paragraph, and enclosing graphic formatting and copy shared styles before an individual formatting edit.

```csharp
OdfTextParagraph paragraph = box.AddParagraph("Status: ");
OdfTextRun status = paragraph.AddRun("Pending");
status.Bold = true;
paragraph.AddHyperlink(" Help", "https://example.com/help");
paragraph.AddText(" Page ");
paragraph.AddField(OdfTextFieldKind.PageNumber, "1");
box.AddList(ordered: true).AddItem("Review order");

status.Text = "Approved"; // Retains the sibling link and field.
status.Color = OdfColor.Parse("#167A36");
drawing.Save("workflow.odg");
drawing.SaveFlatXml("workflow.fodg");
```

Run, hyperlink, and field edits retain surrounding native syntax, list numbering declarations, shape geometry, and document metadata. List items expose their direct paragraphs and nested lists. `OdfTextListItem.AddList` creates a nested list through level ten; `StartValue` sets a nonnegative numbering restart. Page-number, page-count, date, and time fields expose cached display text and fixed state. Page-number/count fields also expose native numbering properties; native applications can refresh dynamic values. Other field kinds, bookmarks, annotations, and unknown inline content remain preserved as XML. Hyperlink targets are never fetched. `InlineNodes` contains snapshot text with live typed editing wrappers; read it again after edits. `ToXml()` returns detached inspection copies.

Assigning shape `Text` replaces all paragraph/list content with plain paragraphs. Assigning paragraph, run, or hyperlink `Text` replaces that element's children, including nested formatting and fields. Use a child wrapper to limit replacement to that portion. Paragraph/list editing is available on text-bearing shapes, text-box frames and image captions; groups and opaque frames reject it. Shape text and inline snapshots use the shared limit of 16,777,216 decoded UTF-16 characters. Text traversal and decoding allow 128 descended containers and 100,000 visited elements or nodes per traversal. A list and its item count as separate containers. Rejected case transforms retain the original XML.

### Draw image captions

Image shapes expose native `draw:image` paragraphs through the same `Text`,
`Paragraphs` and `Lists` APIs. Caption edits retain the image resource and frame
geometry. For example:

```csharp
var image = page.Shapes.AddImage(imageBytes, "server.png",
    OdfRect.FromCentimeters(2, 2, 6, 4), "Server");
var caption = image.AddParagraph("Application server");
caption.FontFamily = "Arial";
caption.FontSize = OdfLength.Points(10);
caption.TextAlign = "center";
```

Drawing and PDF projection use the existing paragraph layout over the original
image frame. Cropping and mirroring apply to the image pixels; caption placement
remains in frame coordinates. A missing or unsupported image resource receives
an omission report while supported caption text remains available. Embedded
object and annotation text stay in their own stories. Frames containing a text
box select it; otherwise the caption APIs select the first `draw:image` story.
Other image/frame text is preserved and reported instead of being combined with
the selected story. Native tables, sections and numbered-paragraph containers
in a frame story are retained; their omission from projection is reported.
Overwide unwrapped paragraphs retain their center or right
anchor; raster and PDF painting retain their horizontal glyph ink and PDF links
cover the complete text advance. Wrapped and vertical clipping retain the text
frame. Native comparisons cover the network fixture's twelve caption instances
and authored one-line ODG/FODG captions with explicit paragraph fonts, crop,
horizontal mirror and image opacity. Native font metrics, automatic fitting,
implicit padding and wrapping defaults and broader text-area layout remain outside this
profile.

Managed raster controls cover translated image frames and group `TransformChildren`
operations across ODG/FODG reopen with fonts registered after projection. Shared
effect controls also cover nested translations and horizontal scaling, explicit
clips, mask registration, cancellation and pixel limits. Native appearance
qualification remains limited to the producer and authored controls above.

`ToDrawing` projects separate paragraphs and nested styled runs through `OfficeDrawingRichText.Paragraphs`. Font family, size, color, bold, italic, supported decorations, background and baseline placement retain their run boundaries. Paragraphs carry alignment, absolute or relative line spacing, absolute margins and indentation; frames carry padding and vertical alignment. SVG, raster and PDF resolve layout with the drawing's render-time fonts, including fonts added after projection. Safe hyperlink targets remain attached to runs and are emitted in SVG and PDF. Fields use the [Draw field projection profile](#draw-fields). Leading explicit spaces and unsupported typography receive explicit loss mappings. Literal XML whitespace uses the [paragraph whitespace profile](#paragraph-whitespace); native tab elements use the [tab-stop profile](#draw-tab-stops). Percentage font sizes are resolved; named-style relative size changes, word-only decorations, nondefault decoration widths and decoration colors differing from the text color are reported as unsupported. Malformed or oversized text retains supported shape geometry; supported text also remains available when custom geometry is skipped. Projection input is bounded to 100,000 UTF-16 characters, charging the larger of each scalar field's source cache and rendered value, 100,000 visited inline nodes across the shape's paragraphs, and 4,096 runs including list labels and paragraph separators. Native XML retains the broader editing limit.

Text projection remains an approximation. Native comparisons cover text-box frame wrapping and indentation, mixed styles, paragraph and vertical alignment, and Latin, Polish, Greek and Cyrillic text with explicit paragraph fonts and side padding. Visual line content matches in the tested frame sample; exact glyph placement differs. The tested LibreOffice build ignores graphic-level font defaults and padding shorthand, and rectangle labels remain unwrapped even with `fo:wrap-option="wrap"`. Omitted wrapping uses the shared wrapped-text default and receives an approximation mapping. Auto-sizing, contour fitting, broader script formatting, complex-script shaping and font fallback are not qualified for identical native appearance. Inspect the conversion report or select `ThrowOnAnyLoss` when approximation is unacceptable.

Native qualification covers an edit to the independent LibreOffice text fixture and targeted run text/color edits in a LibreOffice-resaved mixed paragraph, link, page-field and numbered-list sample. The tested LibreOffice build drops a span nested inside a hyperlink, including its text; that structure passes the ODF schema and remains preserved in OfficeIMO saves, but its native retention is unqualified. Native-produced files can also contain LibreOffice extensions. OfficeIMO retains them; strict OASIS schema validity applies to the authored standard-only sample, not to those extension-bearing files.

### Draw text style cascade

Draw text resolves nested inline styles and their parents, the paragraph style,
`draw:text-style-name`, and the shape's graphic style and parents. Missing
properties then use the initial text or paragraph family's defaults, followed by
graphic defaults. Percentage font sizes resolve against the first absolute size
in that cascade. The original declarations remain in ODG and FODG saves.

Controlled native comparisons cover graphic, paragraph and text-family defaults,
direct graphic formatting, common graphic parent styles, shape-bound paragraph
styles, explicit paragraph formatting, relative inline sizes and bold-only spans
and paragraphs. The tested LibreOffice build retains the common graphic parent
size/color and explicit paragraph size/color, but ignores or rewrites other
bindings, including the relative inline size. Native ODG and FODG resaves keep the
same native PDF paint in these controls. OfficeIMO reopens their saved declarations;
this does not establish identical native appearance. The
[controlled source and native fixture](../OfficeIMO.OpenDocument.Tests/Fixtures/Drawing/README.md)
are separate from the independently authored corpus.

Window-dependent foregrounds (`style:use-window-font-color="true"` or `"1"`)
retain a fixed RGB fallback and receive an unsupported `text-window-color`
mapping, including fully opaque text. Invalid explicit policies receive the same
mapping. Strict conversion rejects this fallback; no window theme color is
inferred. An explicit `false` or `0` restores fixed-color inheritance. As in the
existing foreground profile, a nearer explicit RGB declaration stops an inherited
window policy. This rule is not qualified for identical native appearance.

### Draw fields

Use the same field wrappers on page shapes and master artwork:

```csharp
var document = OdgDocument.Create();
var page = document.AddPage("Overview");
var header = page.MasterShapes.AddTextBox(
    OdfRect.FromCentimeters(1, 1, 14, 2), "Page ", "Header");
var paragraph = header.Paragraphs[0];
paragraph.AddField(OdfTextFieldKind.PageNumber).NumberFormat = "1";
paragraph.AddText(" of ");
paragraph.AddField(OdfTextFieldKind.PageCount).NumberFormat = "1";
document.ClonePage(0, "Review");
// The shared master displays "Page 1 of 2" and "Page 2 of 2".
var scene = document.Pages[1].ToDrawing(
    OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
```

Projection resolves dynamic page numbers in the rendered page's current logical
order, including shared masters, groups, captions, cloned/reordered pages and
cross-document imports. Page count uses the current drawing pages rather than a
saved metadata statistic or printer pagination. SVG, raster and PDF receive the
same resolved runs. Projection leaves native field caches and XML unchanged.
Fields resolve before whitespace normalization and retain surrounding span/link
styles. These logical page values receive approximation mappings.

`NumberFormat` supports decimal `1`, alphabetic `a`/`A`, Roman `i`/`I` through
3999, and an empty string for no number. `NumberLetterSync` selects repeated
letters (`AA`, `BB`) or ordinary alphabetic overflow (`AA`, `AB`). A null format
removes the override; page-number fields inherit the rendered page layout's
number style. If no format is declared, projection uses decimal and reports the
default. `PageSelection` chooses current, previous or next; `PageAdjustment` adds
an integer. A missing selected or adjusted page displays no number. Other native
number formats, malformed attributes and structured field content retain their
cached fallback and receive unsupported mappings.

Fixed page numbers use their cached display text. Date/time projection defaults
to saved display snapshots, with approximation mappings for fixed and dynamic
fields. A missing cache is unsupported in this default mode. Dynamic page-number/count
fields can resolve empty caches. `ThrowOnAnyLoss` rejects these approximations;
`ThrowOnSkippedOrUnsupported` accepts the supported snapshot and logical-page profile.

Opt into saved-value formatting or refresh with explicit settings:

```csharp
document.Styles.CreateDateStyle("ISODate");
OdfTextField date = paragraph.AddField(OdfTextFieldKind.Date, "saved display");
date.DataStyleName = "ISODate";
date.DateTimeValueLexical = "2024-02-29";
date.DateTimeAdjustmentLexical = "P1D";
var fieldOptions = new OdfDateTimeFieldProjectionOptions {
    Mode = OdfDateTimeFieldProjectionMode.StoredValues
};
var formatted = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported,
    false, CancellationToken.None, fieldOptions); // 2024-03-01

fieldOptions.Mode = OdfDateTimeFieldProjectionMode.RefreshDynamic;
fieldOptions.RefreshTimestamp = new DateTimeOffset(2024, 3, 1, 13, 7, 9, TimeSpan.Zero);
var refreshed = document.ToDrawings(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported,
    false, 1000, CancellationToken.None, fieldOptions);
```

`StoredValues` formats both fixed and dynamic fields from their native values.
`RefreshDynamic` uses one supplied timestamp across the whole operation and keeps
fixed fields on their saved values. Projection never reads the clock or changes
field values, display caches or source styles. It formats civil components without
converting supplied offsets to the host time zone. Missing saved values in the
stored/fixed path receive an unsupported mapping rather than an implicit current time.
An explicit data-style binding is required in both evaluation modes.

`Styles.CreateDateStyle` creates a common Gregorian ISO date style;
`CreateTimeStyle` creates a common 24-hour hours/minutes/seconds style.
`FindDataStyle(name, partPath)` resolves that part's automatic definition before
common styles and rejects duplicate names. `OdfDataStyle.ToXml()` returns a detached
inspection copy. ODS Excel number-format APIs continue to use the same native
definitions. Flat serialization keeps common and part-local bindings distinct.
The date/time lexical properties retain offset spelling; null removes an attribute.
Editing accepts Gregorian years 1–9999, clocks through 24:00:00 and at most seven
fractional digits. Imported values outside this profile remain preserved and inspectable.

Evaluation supports explicit-order Gregorian year/month/day, weekday, clock,
AM/PM, literal components and seconds with zero through seven decimal places.
Textual and possessive month names, weekday names, AM/PM and decimal separators
use the declared locale, or `DefaultCultureName` when omitted (invariant by default).
An omitted month inflection uses genitive names when the style includes a day
component; explicit `possessive-form=false` selects nominative names.
Fractional output rounds with carry; integer-second output discards subsecond digits.
Date adjustments apply combined calendar months first, clamping to the destination
month's final day, then day/clock durations. Time adjustments discard fractional
minutes toward zero and wrap when the style requests truncation; the largest clock
component can instead show extended elapsed time. Negative extended time and extended
time combined with AM/PM remain unsupported. Formatting is bounded by the shared
100,000-character drawing story limit; parsed styles are reused only within an operation.

Automatic locale ordering, locale-selected format sources, non-Gregorian calendars,
transliteration, conditional maps, style-level text properties, other components,
ambiguous locale declarations and out-of-range values retain cached text and produce
unsupported mappings. Evaluated values receive approximation mappings because native
applications can normalize offsets or ignore native formatting and adjustments.
SVG, raster and PDF receive the same evaluated runs.

The tested LibreOfficeDev 26.8 build refreshes page numbers and counts but ignores
the tested Roman formats, fixed page-number values, selections and adjustments.
It drops the tested page-count field inside a hyperlink, including populated display
caches; keep page-count fields outside hyperlinks when native Draw retention matters.
It saves master page numbers as native placeholders, which OfficeIMO resolves per page.
The controlled date/time Draw sample covers eleven data styles and twenty saved-value
cases: this native build normalizes offsets but ignores the authored style bindings and
tested adjustments. These observations establish interoperability differences, not
native appearance acceptance for OfficeIMO's component formatter.
The fixture-backed Writer control confirms selected numeric, weekday, localized month
and clock/AM-PM cases. It also normalizes offsets, truncates the tested fractional
seconds, ignores the tested calendar-month adjustment and handles the negative
clock adjustment differently. These differences remain reported approximations.
Native linked-field
retention, exact typography, broader locales/calendars and independent producer corpora
remain outside this qualification.

### Draw text opacity

Edit a paragraph, span or hyperlink through `OdfTextContent.TextOpacity`, or set a shared named/automatic style through `OdfStyle.TextOpacity`:

```csharp
OdfTextParagraph paragraph = box.AddParagraph("Partly transparent text");
paragraph.Color = OdfColor.Parse("#FF0000");
paragraph.TextOpacity = 0.5;
OdfTextRun run = paragraph.AddRun(" Fully opaque span");
run.TextOpacity = 1;
run.TextOpacity = null; // Remove the local override and inherit the paragraph's 0.5.
```

Values are finite fractions from zero to one. A style getter reads its own declaration; a text getter resolves inline, paragraph and graphic inheritance. Invalid assignments leave the document unchanged. Local text edits detach shared automatic styles; editing an `OdfStyle` updates its references. Reading leaves imported declarations unchanged. An explicit edit writes `loext:opacity` and removes the legacy `draw:opacity` alias from that local style; assigning `null` removes both local aliases. Invalid or conflicting imported aliases throw `InvalidDataException` from the typed getters and remain available for repair through the setter. Percentages retain small nonzero fractions without exponent notation. These declarations use the LibreOffice extension profile and do not establish standard-only OASIS schema validity.

Native editing checks cover opacity changes in the independent transparent-text fixture, authored rectangle, line and connector paragraphs, shared automatic-style isolation, clearing a run override, and nested spans. The tested LibreOffice build ignores the authored common named paragraph style, including an automatic paragraph inheriting it, and ignores a hyperlink's own text style from either automatic or common styles. It also drops text in the tested styled span inside a hyperlink. OfficeIMO retains and projects these declarations, but their native appearance is unqualified. The controlled ODG/FODG pairs retain the same native paint; exact typography across producers remains unqualified.

For a styled hyperlink, apply opacity to an outer run that contains the link:

```csharp
OdfTextRun linkContainer = paragraph.AddRun();
linkContainer.TextOpacity = 1;
linkContainer.AddHyperlink(" Help", "https://example.com/help");
```

The tested native PDF export and ODG resave retain this outer-run opacity, hyperlink text and URL. Native SVG export uses a different alpha in the controlled hyperlink case, so that route remains unqualified. Explicit local paragraph color, font and opacity retain their appearance in the tested native SVG, PDF and resave routes; this does not qualify named-style inheritance.

`ToDrawing` reads body-text opacity from LibreOffice's `loext:opacity` extension and the legacy `draw:opacity` import alias on text properties. Native percentages from `0%` to `100%` resolve through inline, paragraph and graphic styles; the nearest declaration takes precedence. Projection applies the resolved alpha to an explicit foreground color and its default decorations through the shared Drawing renderer. ODG and FODG saves retain the original declarations. Fully transparent text stays in the native XML and drawing model while SVG paint is omitted.

Bullets and numbers with an explicit `fo:color` on the list level's `style:text-properties` remain fully opaque, independently of body-text opacity. Valid list-level opacity declarations are retained in XML but do not reduce marker alpha. This profile requires a level-owned foreground color without an automatic/window-color override; a color inherited only from a named label, paragraph or graphic style does not qualify it.

Invalid values, conflicting aliases in one style and unknown opacity namespaces retain opaque text with an unsupported `text-opacity` mapping. Window foregrounds receive a separate unsupported `text-window-color` mapping regardless of opacity. Automatic/window body foregrounds and nonopaque declarations on labels without a qualified level color retain this fallback. `ThrowOnSkippedOrUnsupported` rejects those fallbacks. Text backgrounds and separately colored decorations are outside this opacity profile.

The independent transparent-text fixture and derived percentage controls qualify body alpha separately from broader enhanced geometry and layout. Native controls derived from a resaved numbered caption cover opaque black and blue numbers and bullets beside transparent or partially transparent body text. The tested LibreOffice build removes list-level opacity and named label-style references during resave, and ignores the tested named label color when no inline level color is present. OfficeIMO retains those original declarations; named and default label foregrounds remain outside native appearance qualification. Native SVG and PDF retain the tested partial body opacity, but use different alpha rounding; exact pixels and typography are not qualified across renderers or producers.

### Paragraph whitespace

Native paragraph text and inline snapshots in ODT, ODS, ODP and ODG share the ODF paragraph whitespace decoder. Literal XML spaces, tabs, carriage returns and line feeds collapse to one space across spans and hyperlinks, with leading and trailing literal whitespace removed. The collapsed space belongs to its first source text node, retaining that node's formatting. Explicit `text:s`, `text:tab` and `text:line-break` elements retain their declared spaces, tab and line break and separate adjacent literal whitespace sequences. Nonbreaking spaces and other Unicode spaces keep their characters. Ruby base text is visible; ruby annotations are excluded from the paragraph text.

Reading leaves the source XML unchanged. Assigning typed `Text` writes explicit whitespace elements, so intentional spaces, tabs and line feeds survive saving. Cached field characters retain their opaque fallback; other opaque inline vocabularies and legacy empty-field whitespace barriers are outside this normalization profile. Source character limits include whitespace removed during normalization and expanded `text:s` counts.

Managed checks cover native views, styled Draw projection, ODG/FODG reopen and source limits. A standard-only seven-row drawing passes OASIS 1.4 schema validation and retains its logical text after a LibreOfficeDev Draw resave. Raster and PDF output retain the expected line content. The native exporter spaces styled span and hyperlink boundaries differently; exact appearance, broader independently authored whitespace corpora and nonbreaking-space wrapping remain unqualified. This controlled input is authored by OfficeIMO, not an independent producer corpus.

### Draw list labels

```csharp
OdfTextList steps = box.AddList(ordered: true);
OdfTextListItem first = steps.AddItem("Review the request");
first.AddParagraph("Keep the evidence with the request.");
first.AddList().AddItem("Check attachments");
steps.AddItem("Continue the workflow").StartValue = 7;
```

`ToDrawing` places each item's label on its first visual line. Continuation paragraphs and wrapped lines retain their text indent; list headers remain unnumbered. The profile resolves native bullets, decimal counters including zero, alphabetic numbering with letter synchronization, Roman numbering from 1 through 3999, prefixes/suffixes, per-item styles and restarts, and displayed ancestor levels. Native continuation declarations resolve preceding lists with the same style and level in the shape's text story; ID continuation takes precedence over `continue-numbering`.

Labels have independent fonts and spacing in the shared Drawing layout. Legacy minimum-width label boxes and modern label alignment support left/center/right positioning, measured space separators and no separator. Modern list tabs use the declared position and receive an unsupported `list-tab-stops` mapping because interception by other stops and overflow to the next default stop are not reproduced. Unresolved styles or continuation targets, consecutive numbering, image labels, named number-string formats and unsupported numbering/layout values receive explicit loss mappings; supported item body text remains available.

Label sizes resolve native percentages and bullet-relative sizing. Percentage `fo:font-size` values use the first displayed character's effective size, including span formatting introduced by native resaves; empty item bodies use the paragraph size. Absolute label sizes keep their declared value. `text:bullet-relative-size` uses the paragraph size. An omitted label size inherits the first character size in this projection and receives an approximation mapping; the tested LibreOffice build uses different label defaults. OfficeIMO ODG/FODG round trips retain the tested bullets, nested items, continuation paragraphs, decimal restarts, alphabetic and Roman labels, and targeted text edits. Native PDF export displays those labels, but the tested LibreOffice build resets explicit item start values during both ODG and FODG resave. Native counter restart retention is unqualified. A declared label size of `100%` retains the first character size in that sample; the tested producer replaces an absolute label size with its smaller percentage default. Spacing beside wide alphabetic and Roman labels also differs. Exact native font metrics, RTL list placement and interception of list tabs by paragraph stops remain outside the qualified profile. Inspect the report before relying on layout fidelity.

### Draw tab stops

```csharp
OdfTextParagraph prices = box.AddParagraph("Item\t123,45");
prices.TextAlign = "right";
prices.SetTabStops(new[] {
    new OdfTabStop(OdfLength.Centimeters(6), "char", ",").WithLeader(".")
});
```

`SetTabStops` replaces the paragraph's complete explicit set with at most 256 unique absolute positions. An empty set clears inherited explicit stops without editing the parent style. The native types are `left`, `center`, `right` and `char`; character stops require one Unicode scalar delimiter. `WithLeader` returns an independent stop with one non-control Unicode scalar for native `style:leader-text`; `null` removes it and SPACE leaves a blank gap. ODG and FODG save/reopen retain these declarations. Page cloning and import retain leader text-style references in their original automatic or common scope; missing dependencies reject import before either document changes.

`ToDrawing` resolves the nearest complete tab set, including an empty override, and measures following fields across styled runs. Default spacing and the inner/outer margin origin come from the default paragraph style; properties with these names in ordinary styles do not change the document defaults. For relative body tabs, a generated legacy list-body indent does not move the paragraph's original tab origin. When a modern list supplies an omitted paragraph margin, that effective margin remains the relative body-tab origin. An omitted distance uses 36 points with an approximation mapping. Unsupported or duplicate stops and writing directions receive loss mappings while supported text remains available. List-label tab interception remains separate from body tabs. Soft wrapping and tabs beyond the frame follow the shared Drawing overflow contract.

Centered and right/end paragraphs align each complete tabbed line within the text area. Tab fields retain their declared left, center, right or character relationship within that line; plain lines retain their paragraph alignment. Justified tabbed lines retain their tab spacing without distributing extra word space. They receive an approximation mapping and pass `ThrowOnSkippedOrUnsupported` when the remaining text profile is supported. The text area uses its declared width when the horizontal anchor is omitted or `justify`. Explicit left/center/right area anchors use the separate [line-caption](#draw-line-and-connector-labels) and [rectangle/text-box](#draw-text-areas) profiles; their native intrinsic block placement is distinct from paragraph alignment. Centered or right/end tab paragraphs with a hanging indent, and non-left list-body tabs, retain their unsupported mappings. Hard-break justification keeps its existing unsupported mapping.

Controlled native Draw exports with full-width justified text areas cover all four tab types, start/end aliases, multiple stops, paragraph margins, positive first-line indentation and plain lines around hard breaks. Fitting explicit field starts agree within 0.03 points. Default-grid starts differ by up to 0.6 points even with declared default spacing. Native exports of the separate left-area indentation and wrapped-prefix controls place fields outside the page; shared layout retains the fields under its bounded overflow contract. These cases, exact wrapping, baselines, font metrics and additional independently authored producers remain outside native appearance qualification.

Textual leaders take precedence over native line-pattern declarations. Shared SVG, raster and PDF layout repeats the character in the tab run's active font and formatting without changing the following field's position. Glyph phase and typography receive an approximation mapping. Generated text paint is bounded to 100,000 UTF-16 characters per text frame across its paragraphs; exhaustion marks layout clipping while retaining tab spacing and body text. Frame fitting does not shrink body text merely because leader paint reaches this limit. Logical drawing text retains its tabs; rendered PDF text includes the visible leader glyphs. Invalid leader text or a missing referenced text style retains its source XML and blank tab gap with an unsupported mapping.

Bind a textual leader to a common or automatic text style:

```csharp
OdfStyle leaderStyle = drawing.Styles.CreateNamed("PriceLeader", OdfStyleFamily.Text);
leaderStyle.Color = OdfColor.Parse("#0000FF");
leaderStyle.FontSize = OdfLength.Parse("150%");
prices.SetTabStops(new[] {
    new OdfTabStop(OdfLength.Centimeters(6), "char", ",").WithLeader(".").WithLeaderTextStyle(leaderStyle)
});
```

`SetTabStops` checks the style's family, document and package-part visibility before changing XML. Projection resolves common-only bindings from common/default styles and part-local bindings from automatic styles, including inherited graphic paragraph properties. Save/reopen, cloning and import retain the original binding across name collisions. The referenced parent chain overrides font family, size, foreground/opacity, bold/italic state, background, decorations and baseline; omitted properties inherit the actual tab run. Relative-only sizes multiply its font size; percentages above an absolute parent size become an absolute size. Text casing uses the existing Draw run profile; expansion beyond one Unicode scalar receives an unsupported mapping. Advanced decoration, script and font metrics retain the existing text loss reports. Separately styled glyphs receive an approximation mapping: the tested native Draw exporter ignores the binding and removes it on resave, so identical native appearance and binding retention are unqualified.

Use `WithLineLeader` to author native line declarations:

```csharp
prices.SetTabStops(new[] {
    new OdfTabStop(OdfLength.Centimeters(6), "right").WithLineLeader(
        new OdfTabLineLeader("dot-dash", "double", "0.75pt", OdfColor.Parse("#0000FF")))
});
```

ODG/FODG reopen, page cloning and import retain `leader-style`, `leader-type`, `leader-width` and `leader-color`. Shared projection supports solid, dotted, dash, long-dash, dot-dash, dot-dot-dash and wave, with single or double lines. Style or type `none` leaves a blank gap. Explicit RGB colors override the active text color; absent color or `font-color` inherits it. A declared text character, including SPACE, takes precedence. `leader-text-style` has no effect on line paint. Invalid active line declarations retain their source and tab spacing with an unsupported mapping.

Line rendering receives an approximation mapping. Absolute lengths retain their point width. The font-relative profile uses 5% of rendered font size for `auto`, `normal` and `medium`, 2.5% for `thin`, and 10% for `bold` and `thick`. Positive integer and percentage widths multiply the 5% automatic width; this is OfficeIMO's declared evaluation profile, not a claim that ODF defines that base. Named widths are implementation-defined in the [ODF 1.4 contract](https://docs.oasis-open.org/office/OpenDocument/v1.4/OpenDocument-v1.4-part3-schema.pdf#page=481). Line paint adds no characters to SVG or PDF text extraction. Shared outlines are bounded to 8,192 vertices per tab and 100,000 per frame; truncation marks clipping while retaining field positions and body font sizes.

Native qualification covers left, center, right and comma alignment, a missing delimiter, mixed font sizes, leading/consecutive/default tabs and an inset paragraph in one LibreOfficeDev Draw sample. Its horizontal field starts agree within 0.2 points with the shared PDF output; vertical font placement differs. A related typed probe qualifies dot, hyphen and underscore leader export and resave: authored ODG and FODG have equal native PDF text and field bounds, and six compared field starts agree with shared PDF within 0.2 points. Typed textual stops include an explicit activating line-style declaration because the tested producer omits text-only leaders when that declaration is absent. The retained native FODG fixtures are exports of OfficeIMO-authored probes, separate from independently authored drawing corpora. An additional explicit-padding probe qualifies zero advance at an overlapping right stop and selection of the following stop. A 33-case line control in LibreOfficeDev 26.8.0.0.alpha0 renders dotted declarations as text dots and other visible patterns as underscores; its resave drops line count, width and color. OfficeIMO preserves and paints those declarations through its own containers and reports approximate appearance. Identical native line-pattern appearance and native declaration retention remain unqualified. Broader producers, native retention of separately styled text leaders, non-left paragraphs, list tabs and exact glyph phase/font metrics also remain unqualified. Literal XML whitespace is handled separately by the [paragraph whitespace profile](#paragraph-whitespace).

ODF groups do not have a native `draw:transform` attribute. Add their children first, then call `group.TransformChildren("scale(1.5 1) translate(2cm 1cm)")` to apply an affine transform to the existing contents. This composes leaf transforms and bakes free connector and line coordinates while retaining nested groups. Attached connectors, rotated/scaled/sheared connector labels, and unknown elements are rejected before any child changes; detach connector ends before transforming that group. Assign `Transform` directly only to leaf shapes.

Attach shapes through persistent glue points:

```csharp
OdgShape target = page.Shapes.AddEllipse(OdfRect.FromCentimeters(12, 3, 6, 3));
OdgShape connector = page.Shapes.AddConnector(
    box.AddGluePoint(OdgGluePointAlignment.Right),
    target.AddGluePoint(OdgGluePointAlignment.Left));
box.Bounds = OdfRect.FromCentimeters(2, 3, 7, 3); // The connector follows the new edge.
drawing.Layers.Add("Review", OdgLayerDisplay.Screen);
box.Layer = "Review";
```

Connectors expose routing type, attached shape IDs, and resolved endpoints. Straight connectors and saved orthogonal, segmented, and curved routes render through the Drawing engine when their endpoints match the current attachments. Renaming shape IDs updates connector references. Removing a shape detaches its connected endpoints at their last resolved positions; native routes that cannot be resolved use saved endpoint coordinates when available. Explicit glue points use absolute offsets from an edge or corner; standard imported edge centers, relative percentage points, and legacy Draw relative-length points are also resolved. `OdgGluePointAlignment.Bottom` creates a bottom-center relative point that follows resizing and requires zero offsets. `EscapeDirection` accepts automatic, left, right, up, down, horizontal, and vertical constraints. Editing a shared glue point changes the native constraint for every connector referencing it; saved path projection does not recalculate routes from that constraint.

Create a free curve and provide its route in points:

```csharp
using OfficeIMO.Drawing;

OdgShape routed = page.Shapes.AddConnector(
    new OfficePoint(50, 80), new OfficePoint(250, 140), OdgConnectorKind.Curve);
routed.SetConnectorRoute(new[] {
    OfficePathCommand.MoveTo(50, 80),
    OfficePathCommand.CubicBezierTo(120, 30, 180, 190, 250, 140)
});
IReadOnlyList<OfficePathCommand> route = routed.ConnectorRouteCommands;
```

Route commands use the connector's coordinate space before its own and parent transforms. `SetConnectorRoute` accepts one open path, updates free endpoints, and requires attached ends to be within 0.1 point of their glue positions. Accepted attached ends are snapped to their glue positions before native-grid rounding. A `Line` connector accepts one straight segment; choose another kind before adding bends or curves. The writer rounds coordinates to the native 1/100 mm grid, requires signed 32-bit native coordinates and canvas dimensions, and replaces previous line-offset parameters. It shares the path limits of 20,000 commands and 1 MiB of native geometry text; invalid edits leave the previous route intact. Native applications may regenerate a saved route from its routing type and attachments. Changing a free endpoint, attachment, or routing type clears the saved route. Moving an attached shape can make a saved route stale: replace it explicitly or recalculate it in a native application. `ToDrawing` reports stale routes and non-straight connectors without a saved route instead of projecting incorrect geometry.

### Partially attached straight connectors

After moving or resizing a target, save the current straight route before handing a partially attached connector to a native drawing application:

```csharp
connector.ConnectorKind = OdgConnectorKind.Line;
connector.SetConnectorRoute(new[] {
    OfficePathCommand.MoveTo(connector.X1.ToPoints(), connector.Y1.ToPoints()),
    OfficePathCommand.LineTo(connector.X2.ToPoints(), connector.Y2.ToPoints())
});
```

This keeps the attachment and writes both endpoint coordinates and the straight path. LibreOfficeDev 26.8.0.0.alpha0 ODG/FODG save/reopen checks retain the tested geometry and bindings for one custom glue point on an untransformed rectangle and one free endpoint. The controls cover start and end attachments, page and master artwork, either shape order, target movement/resizing, and an independently edited page/master clone. Shared SVG, raster and PDF projection remains available under strict conversion after native saving. These are controlled authoring samples, separate from an independently authored diagram corpus.

Refresh the route after further target edits; an existing cache does not follow those edits automatically. The tested native build can save a cache starting at zero when attached endpoint coordinates are omitted, even though it renders the live attachment correctly. OfficeIMO reports that inconsistent cache rather than using it for projection. Other route kinds, automatic glue selection, transforms, labels, broader attachment geometry and other producers remain outside this native profile.

### Connector routing

`AttachStartToShape` and `AttachEndToShape` select among a bounded shape's four edge centers; passing null detaches the end. A valid saved edge position takes precedence, otherwise the nearest edge center toward the other shape or free end is selected. Projection reports this bounded selection as an approximation. Attachment selection does not generate a route. Attachments to groups remain outside this profile.

Generate a route after attaching the endpoints:

```csharp
routed.RouteOrthogonal(OdgConnectorRouteOrientation.HorizontalFirst, offset: 12);
routed.RouteOrthogonalAroundShapes(page.Shapes, padding: 6, maxLanes: 12);
```

Both methods replace the saved route and select `Standard` routing. Offset and padding use connector-local points before transforms. Obstacle routing searches deterministic orthogonal lanes around the transformed declared bounds of the supplied shapes, with an extra 0.1 point clearance for coordinate rounding. It accepts up to 4096 input shapes and 32 lanes in each direction. Attached shapes and line/connector inputs are ignored; supply individual bounded children for groups. A failed enumeration, unsupported shape, invalid geometry, or exhausted search leaves the previous route intact. Increase the lane count, adjust endpoints or padding, or supply an explicit route when no clear lane is found. This bounded search does not guarantee a solution, enforce custom escape directions, avoid other connectors, or reproduce a native application's routing algorithm. Native applications can subsequently recalculate the saved route.

Enable glue-point constraints when the route must leave its attachments in specific directions:

```csharp
startPoint.EscapeDirection = OdgGluePointEscapeDirection.Down;
endPoint.EscapeDirection = OdgGluePointEscapeDirection.Up;
connector.RouteOrthogonalAroundShapes(page.Shapes,
    respectEscapeDirections: true, padding: 6, maxLanes: 12);
```

With `respectEscapeDirections: true`, routing and padding use page coordinates. The search honors both explicit departure constraints and includes the attached shapes' transformed declared bounds, even when they are omitted from the obstacle collection. Only an endpoint's terminal segment may cross its own attached shape's bounds. Free, automatic, and standard glue endpoints permit either axis. Short routes are tried first, followed by bounded exit-segment detours; opposing departures and same-shape loops can require five segments. Affine connector transforms are retained while the page route is mapped into the connector's coordinate space. Native-grid rounding must retain orthogonality, departures, and endpoints within 0.1 page point; failure leaves the document unchanged. This is a bounded candidate search, with no guarantee of finding every possible route.

LibreOffice save/reopen checks retain the tested horizontal, vertical, opposing, same-shape, and partially attached departure constraints, including an edit to an independently produced diagram. Recalculated `Standard` paths can cross obstacles, and transformed connectors can acquire a straight saved cache even when a later native load draws a detour. Saved-cache projection and native rendering are separate outcomes. Use the three-segment native profile below when the route is eligible; arbitrary cached-route retention remains outside that profile.

LibreOffice recalculates attached `Standard` connectors on load and can replace a cached detour with a path through an obstacle. Its handling of a connector's own transform can also invalidate that cache. Bake a free connector's own transform into its endpoints, bends, and curve controls before saving:

```csharp
routed.Transform = "scale(1.3 0.7) translate(20pt 30pt)";
routed.BakeConnectorTransform();
```

`BakeConnectorTransform` clears the transform and routing offsets while retaining the routing kind, metadata, and declared styles. Scaling changes geometry without scaling the declared stroke width. Both endpoints must be free; transformed parents, rotated/scaled/sheared labels, invalid routes, and native coordinate overflow are rejected before mutation. Translation baking retains connector text and its styles. `TransformChildren` uses the same free-connector implementation. Native save/reopen qualification covers polyline detours after translation, nonuniform scale, quarter-turn rotation, and shear, plus a cubic curve after an affine matrix. Broader paths and transformed label placement remain open qualification work.

Encode a saved three-segment route as native routing parameters while keeping its attachments:

```csharp
connector.RouteOrthogonalAroundShapes(page.Shapes, padding: 6);
connector.UseNativeThreeSegmentRouting();
```

`UseNativeThreeSegmentRouting` selects `Lines` routing and writes departure directions, zero endpoint spacing, and two segment offsets. Both ends must attach to untransformed rectangles, full ellipses/circles, or frames. The outer segments must be nonzero and horizontal or vertical in page coordinates; the middle segment may be diagonal. The method reuses equivalent directional glue points or creates copies, preserving points and styles used by other connectors. It bakes the connector's own transform into its route and removes that transform; transformed parents and transformed connector text are rejected. Invalid or unsupported input leaves the document unchanged.

Native save/reopen qualification covers upward/downward detours, horizontal departures, untransformed shapes within groups, same-shape loops, baked connector translations/scales/quarter turns, moving an attachment, and edits to the [independent LibreOffice diagram fixture](../OfficeIMO.OpenDocument.Tests/Fixtures/Drawing/README.md). Native applications can shift bends for painted shape bounds, strokes, and subsequent edits; this profile preserves departure constraints and offsets rather than exact bend coordinates. It does not guarantee obstacle avoidance after native recalculation or shape edits. Partially attached three-segment routes, other attachment geometry, transformed targets, and arbitrary attached cached paths remain outside this native retention profile.

`drawing.Layers` declares document layers; `page.MasterLayers` declares layers on the referenced master; `page.Layers` declares a page override. Each explicit set replaces the inherited set. `page.EffectiveLayers` resolves page, master, and document definitions. `Display` distinguishes screen and print visibility, and `IsProtected` records the native editing hint. Older LibreOffice saved-view layer masks are read when explicit layer attributes are absent; editing a document layer flag updates existing view masks too. Page and master layer edits leave those global masks unchanged. Visibility edits validate all affected masks before changing any saved view. Malformed masks leave existing layers unchanged, including when a new layer cannot be added. Reader text extraction deliberately includes hidden layers.

Share a master layout and its layers between drawing pages:

```csharp
page.MasterLayers.Add("Guides", OdgLayerDisplay.Screen);
OdgPage detail = drawing.AddPage("Detail");
detail.MasterPageName = page.MasterPageName;
```

Assigning `MasterPageName` requires an existing, unambiguous master with a resolvable page layout. It changes the inherited layout and layers while retaining the page's shapes and explicit layer set. Editing a shared master layer or layout affects all pages that inherit it. Master selection and scoped layer edits survive OfficeIMO ODG/FODG save/reopen. `ToDrawing` projects supported master artwork beneath page content, using the same geometry and text profile as page shapes.

### Clone drawing pages and masters

Duplicate a page, then give the copy an independent master when its layout or inherited layers need to change:

```csharp
OdgPage copy = drawing.ClonePage(0, "Review");
copy.MasterPageName = drawing.CloneMasterPage(page.MasterPageName, "ReviewMaster");
copy.Width = OdfLength.Centimeters(30);
copy.Shapes[0].Text = "Review copy";
```

`ClonePage` appends a copy within the same document. It assigns fresh native IDs and shape names, remaps copied connector targets, navigation order, chained text boxes, continued lists, numbered-paragraph list groups, note references, and simple local fragment links, and retains geometry, styled text, transforms, layers and cached routes. Glue-point numbers stay local to their copied shapes. Note names and numbered-paragraph groups use separate mappings from XML IDs. Links to other pages or external documents remain unchanged; external resources are not fetched. Styles and immutable image package entries are shared. Shape and paragraph property edits use the existing automatic-style copy-on-write behavior; editing named styles directly still affects their consumers.

Pages share the original master until a different `MasterPageName` is assigned. `CloneMasterPage` copies the master artwork and layers with a new page layout, remaps local IDs and references, and returns its new name. Both methods generate unique names when the destination name is omitted; master names must be XML NCNames, such as `ReviewMaster`. They reject ambiguous identifiers, invalid local attachment or chain targets, embedded editable objects, forms, animations, tracked changes, index-mark ranges and named text/table definitions before modifying the document.

ODG/FODG round trips qualify independent edits, references and image bytes. The independent routed-glue and sheared-geometry fixtures produce identical original/copy page pixels in LibreOffice PDF export. Native resave retains the tested copied connector targets and separate master-layout declarations. LibreOfficeDev 26.8.0.0.alpha0 exports different-sized Draw pages using the first page's paper size, which can clip a wider copy; mixed page sizes are not qualified for that native PDF exporter.

Scoped layer visibility is not qualified for identical native appearance. The tested LibreOffice build removes page and master layer sets during resave and writes layer flags globally. Its PDF export also excludes layers hidden on screen even when they are printable. OfficeIMO retains the ODF declarations and uses the selected screen or print intent in `ToDrawing`; native printing remains unqualified.

### Copy a page from another drawing

Import a page into a different document with its own master, layout, reachable styles and embedded resource bytes:

```csharp
OdgDocument source = OdgDocument.Load("diagram.odg");
OdgDocument destination = OdgDocument.Create();
OdgPage imported = destination.ImportPage(source, 0, "Review");
imported.Shapes[0].Text = "Review copy";
destination.Save("combined.odg");
```

`ImportPage` accepts drawings loaded from ODG or FODG. It preserves named and part-local automatic style dependencies, paint definitions, list styles and font declarations with fresh names. Repeated references to an embedded image share one copied package entry; existing destination entries are not replaced. Page and master IDs, local links and connector targets are remapped together. The imported page snapshots its effective source layer visibility and protection. The source and existing destination wrappers remain usable; edits to the imported layout, styles and shapes are independent.

Import resolves dependencies before attaching XML or resource bytes. It rejects the unsupported cloning content listed above, scripts, event listeners, references outside the copied page/master, other-master dependencies, relative external links, externally linked resources, ambiguous or missing definitions, and style dependency chains deeper than 256. Absolute external hyperlinks are retained without fetching them. The destination must use the source ODF version or newer; transferring a nonzero unitless native gradient/opacity angle across ODF 1.2 is rejected because its meaning can differ by producer.

Source default properties are captured in imported styles. Paragraph and inline fallbacks are captured in their original graphic context so they do not override explicit shape formatting. Captured fallback properties become local formatting; change the imported paragraph or span style to change those values. Percentage font sizes that depend on source defaults are resolved to absolute sizes in the imported text styles; percentages with an explicit base retain their native binding. Percentage sizes combined with nonzero relative size changes, class-style bindings and conditional text styles are outside this snapshot profile. Destination-only default properties that cannot be canceled by the snapshot are rejected; importing into a new drawing provides a destination without those defaults. This profile does not copy the source document's metadata, scripts, saved views or unused styles and resources.

ODG/FODG round trips cover collisions, independent editing, images, styled text, lists and paint definitions. Native PDF export produces identical source/import page pixels for the generated approval diagram and the independent routed-glue, sheared-geometry and transparent-text fixtures. This evidence does not qualify every producer, scoped native layer behavior or mixed page sizes; the native limitations above still apply.

Create and edit local geometry without changing its coordinate system:

```csharp
OdgShape curve = page.Shapes.AddPath(OdfRect.FromCentimeters(2, 8, 6, 3),
    new OdfViewBox(0, 0, 200, 100), "M0 100 Q0 0 100 0 T200 100 Z");
curve.Bounds = OdfRect.FromCentimeters(2, 8, 9, 3); // Scales the same path horizontally.
curve.FillRule = OfficeIMO.Drawing.OfficeFillRule.EvenOdd;
page.Shapes.AddRoundedRectangle(OdfRect.FromCentimeters(13, 8, 6, 3),
    OdfLength.Centimeters(0.4));
```

`PathData` retains supplied SVG notation; `PathCommands` and `SetPathCommands` use shared Drawing commands and normalize shorthand and arcs. `AddPolygon`, `AddPolyline`, `Points`, and `SetPoints` use integer vertices. `ViewBox` requires signed 32-bit coordinates and positive integer dimensions; bounds provide the page position and scale, including a nonzero local origin. Geometry reads and edits accept at most 20,000 commands or vertices and 1 MiB of native geometry text. Invalid edits leave the previous geometry intact. `CornerRadius` sets circular rectangle corners; `CornerRadiusX` and `CornerRadiusY` set elliptical corners. Each form clears the alternate attributes when assigned.

Create reusable native gradient fills:

```csharp
document.Styles.CreateGradient("Ocean", new OdfGradientPattern(
    OdfGradientStyle.Linear, OdfColor.Parse("#2040E0"), OdfColor.Parse("#DC3020"),
    angleDegrees: 45, border: 0.2));
box.FillGradientName = "Ocean";
box.GradientStepCount = 0;
```

`Styles.Gradients` and `FindGradient` expose common definitions. `gradient.Pattern` reads or replaces a native two-color definition. The model covers linear, axial, radial, ellipsoid, square and rectangular styles, with angle, border, endpoint intensity and center. Fractions use one for 100%; angles are clockwise from vertical. Explicit authoring writes degree units. Imported omitted centers resolve to the native zero-offset default; constructor defaults author a centered gradient. Names require XML NCNames, and invalid edits leave definitions and bindings intact.

`FillGradientName` and `GradientStepCount` resolve inherited graphic styles and use copy-on-write for shape edits. A gradient binding selects gradient fill while retaining inactive solid-color attributes. Setting the name to `null` explicitly disables fill. Replacing a definition changes referencing shapes and retains unrelated attributes. Equivalent LibreOffice endpoint stop extensions are retained and updated with the colors; other extended stops and SVG gradient definitions remain preserved but cannot be overwritten through the two-color model. Duplicate names across native and SVG definitions are rejected.

`ToDrawing` projects linear, axial and radial fills using actual geometry bounds, including padded path canvases, borders and color intensities. Linear angles retain their physical direction on nonsquare shapes. Axial fills mirror the start color about the axis center; radial fills use a physical circle with the end color at its center. Uniform fill opacity is applied to the gradient. Automatic native banding is approximated by continuous interpolation. Native save/reopen and rendered comparisons cover angles, borders, intensities, offset radial centers, compound holes and opacity. The tested native producer relocates padded paths and curves during import; OfficeIMO retains their declared geometry. Rotation, reflection, shear and nonuniform scale retain the declared local gradient but add an unsupported `fill-gradient-transform` mapping because native gradient refitting is not reproduced. These geometry and affine cases are not qualified for identical native appearance.

Ellipsoid, square and rectangular fields, fixed step counts, off-area radial centers, the singular center seam of a 100% axial border, and ambiguous nonzero unitless angles from ODF 1.2 produce explicit unsupported mappings. Active missing, malformed or ambiguous definitions retain the stroke geometry and original XML. Open paths do not resolve or paint inactive gradient bindings. Inspect the conversion report or select a throwing loss policy when an unsupported fill must stop conversion.

Saving an ODF 1.2 source with ambiguous nonzero unitless native gradient or opacity angles to a newer ODF version is rejected before writing. Use `new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }` to retain the source interpretation, or explicitly replace each affected definition with an angle unit. Flat XML output retains the source version.

Set independent paint opacity as fractions from zero to one:

```csharp
box.FillOpacity = 0.5;
box.StrokeOpacity = 0.25;
// For an embedded image frame, ImageOpacity changes image pixels.
// FillOpacity describes the frame's background fill separately.
```

`FillOpacity`, `StrokeOpacity`, and `ImageOpacity` are shared ODF shape properties. They resolve named, automatic, and default graphic styles; assigning `null` removes the local override. Invalid numeric edits are rejected before styles change. Fill and image values use native percentages; stroke reads also accept ODF fractional notation. The drawing projection applies uniform fill/stroke opacity to geometry and image opacity to embedded pixels through the shared Drawing renderer. Native save/reopen retains the tested rectangle, ellipse, line, transformed path, and image values; rendered interior colors match within raster rounding. Group paint opacity and opacity gradients remain outside this projection profile and produce unsupported diagnostics. Uniform fill-opacity access rejects an active opacity gradient instead of replacing it.

Crop an embedded image without changing its bytes:

```csharp
var image = page.Shapes.AddImage(pngBytes, "picture.png",
    OdfRect.FromCentimeters(1, 1, 6, 3));
image.Crop = OdfInsets.FromCentimeters(0.2, 0.3, 0.2, 0.4);
image.Mirror = OdfImageMirror.Horizontal;
```

`Crop` stores native `fo:clip` offsets in top, right, bottom, left order. The offsets refer to the image's intrinsic physical dimensions, calculated from pixels and embedded resolution, so resizing the frame retains the source region. Negative offsets add empty margins inside the frame. Named, automatic, and default graphic styles participate in inheritance; assigning `null` removes the local override, while explicit zero insets disable inherited cropping. Edits use copy-on-write styles and retain the original image resource. Crop access on a shape without an image throws. `OdpImage.Crop` uses the same native codec.

`Mirror` stores native `style:mirror` without changing image bytes. It uses the same inherited graphic styles and copy-on-write rules as cropping. Assign `null` to remove the local override or `OdfImageMirror.None` to disable inherited mirroring. Vertical mirroring may be combined with one horizontal mode; conflicting horizontal modes and undefined flags fail before modifying styles.

The shared Drawing projection applies raster crop windows, negative margins, horizontal mirroring, image opacity, and shape transforms. Mirroring reflects both the cropped source pixels and blank margins around the original frame center. An empty intersection with source pixels remains empty. Invalid crop syntax, collapsed intrinsic dimensions, and cropped vector images produce explicit unsupported mappings with an uncropped image fallback. Vertical, combined, and page-dependent mirror modes remain unsupported and use an unmirrored fallback that retains supported cropping. `ThrowOnSkippedOrUnsupported` rejects these fallbacks. Native comparisons cover PNG crop regions, resized frames, negative margins, horizontal mirroring, transparency, rotation, nonuniform scale, shear, and transformed group children. The tested LibreOffice build ignores vertical and page-dependent `style:mirror` modes and rewrites them as `none`; OfficeIMO retains their declared values. Resaving equal-pixel PNGs with different resolution metadata can remove that metadata and change the visible crop region. OfficeIMO package and flat saves retain the original bytes; broader native producer and image-format qualification remains open.

Create a named dash definition, then reference it from a shape:

```csharp
var dash = document.Styles.CreateStrokeDash("LongShort", new OdfStrokeDashPattern(
    1, OdfLength.Parse("300%"), OdfLength.Parse("150%"),
    secondDashCount: 1, secondDashLength: OdfLength.Parse("100%")));
line.StrokeDashName = dash.Name;
line.StrokeLineCap = OfficeStrokeLineCap.Butt;
line.StrokeLineJoin = OfficeStrokeLineJoin.Bevel;
```

`Styles.StrokeDashes` and `FindStrokeDash` expose common named definitions. Replacing `dash.Pattern` updates every referencing shape and retains unrelated definition metadata. Patterns contain one or two repeated sequences with a total of 1–1024 dashes, finite nonnegative lengths, and a positive gap. Metrics accept absolute ODF lengths or percentages of stroke width. The shared Drawing projection expands the two sequences into an exact dash array; changing `StrokeColor` retains an active dash selection. Assigning `StrokeDashName = null` explicitly selects a solid stroke. Names and references require XML NCNames such as `LongShort`; metric units use lowercase ODF notation such as `8pt`. Invalid names and patterns are rejected before XML changes.

Caps and joins resolve named, automatic, and default graphic styles. Assigning `null` removes only the local override. A shape cap overrides the dash definition's round/rect cap; the projection uses butt caps otherwise and round joins when no join is declared. Native save/reopen and rendered stroke runs cover explicit butt/round/square caps, miter/round/bevel joins, and uniform absolute or percentage dash metrics. Mixed absolute/percentage metrics and round caps supplied only by a dash definition are not qualified for identical native appearance: the tested LibreOffice build changes the former and does not reproduce the latter. Native grid rounding can also introduce a short terminal dash at a pattern-cycle boundary, so exact endpoint dash phase is not qualified. Imported omitted dash metrics, excessive counts, unresolved or duplicate names, and native middle/none joins remain preserved but produce conversion loss diagnostics. Additional overlaid dash definitions are preserved and reported as unsupported while the primary geometry remains projected.

Create a named marker path and bind it to either end of a line or connector:

```csharp
var arrow = document.Styles.CreateMarker("Arrow", new OdfMarkerGeometry(
    new OdfViewBox(0, 0, 20, 30), "M10 0l-10 30h20z"));
line.StrokeEndMarkerName = arrow.Name;
line.StrokeEndMarkerWidth = OdfLength.Points(12);
line.StrokeEndMarkerCentered = false;
```

`Styles.Markers` and `FindMarker` expose common definitions. Replacing `marker.Geometry` updates referencing shapes and retains unrelated metadata. Geometry accepts bounded SVG paths, including curves, arcs and multiple contours, with a finite positive view box. The string constructor reads decimal producer view boxes; creating or replacing a definition writes an enclosing integer view box as required by ODF. Reading an imported definition retains its original XML. Names require XML NCNames. Invalid geometry, widths and references are rejected before editing styles.

Start/end marker names, absolute widths and centering resolve inherited graphic styles and use copy-on-write when edited. Setting a name to `null` writes an explicit empty reference to disable an inherited marker. Setting width or centering to `null` removes the local override. A zero width suppresses painting; an omitted effective width is reported as unsupported. Names do not enable a disabled stroke.

The Drawing projection scales marker paths from their actual curve bounds, preserves their aspect ratio, and aligns the top or center to each drawable contour endpoint. Draw uses one closure state for a whole path: fully closed shapes suppress markers; any open contour, including an empty move, opens all contours while retaining their closing edges. Empty and zero-length contours receive no markers. An open shape does not paint its declared fill; `inactive-fill` diagnostics identify this behavior while ODG/FODG retain the original fill and geometry. Fully closed compound paths retain their complete fill and holes.

Marker height consumes the outline, with a marker-width/15 overlap. Notched polygon backs use their centerline exit. Curved strokes retain Bezier commands and orient markers along the consumed path chord. Stroke color and opacity cover each contour's shaft and markers together, so hollow markers do not expose the original shaft and overlaps within that contour do not darken. Distinct contours accumulate transparency where they overlap. Short contours consumed by both markers have no remaining shaft. Path measurement shares a bounded sampling budget across the whole shape; unsupported measurement retains the underlying geometry and reports the marker loss without claiming partial marker success.

Transformed paint canvases include marker and stroke extents. Missing or malformed active definitions produce diagnostics while retaining the underlying geometry. Closed shapes retain inactive marker bindings without resolving them. Native save/reopen and rendered bounds cover straight and diagonal arrows, centered markers, curved marker definitions and rings. Native comparisons also cover varied marker/stroke widths, notched backs, transparent overlaps, ring holes, exhausted short lines, compound paths, mixed closure, commands after a close, empty contours and per-contour transparency. Native single-segment curve trimming can alter curve controls; OfficeIMO retains the original analytic curve. Exact native curve appearance, dash phase at marker docking, and nonuniform affine marker appearance remain unqualified.

`page.ToDrawing()` returns a point-based `OfficeDrawing` with an `OdfConversionReport`. The projection covers rectangles and rounded rectangles, full ellipses and circles with bounding dimensions, SVG paths, polygons, polylines, [literal enhanced custom shapes](#draw-enhanced-geometry), lines, text boxes, supported embedded images and raster crop windows, attached and free connectors with supported saved routes, and nested groups. ODF affine transforms use radian angles and ordered operations, including LibreOffice's rotation/skew direction convention. Shapes outside the page remain positioned in the shared scene. Pass its value to the shared Drawing SVG/raster exporters or compose it into a PDF using `OfficeIMO.Pdf`. Inspect the report before using the result: native text-frame transforms beyond the line/connector label profile, enhanced geometry outside the literal profile, ellipse/circle arcs and sectors, backgrounds beyond the [solid, native gradient and stretched bitmap profile](#draw-page-backgrounds) and advanced styles are outside this projection profile. Styled text uses the paragraph profile described above, with explicit losses for unsupported formatting and native layout differences. Native zero-width hairlines use a reported 0.25-point approximation. `ThrowOnAnyLoss` rejects approximations as well as omissions. Use `ToDrawing(forPrint: true)` for print-visible layers; the default uses screen-visible layers. This is an explicit drawing projection, not a full Draw layout engine or an ODG-to-Visio converter.

The geometry profile follows native ODF attributes, but producer rendering can differ. LibreOffice validation retains ordinary curves and polygons under skew and matrix transforms. The tested native resave drops elliptical corner radii and explicit nonzero fill rules, and renders oversized circular radii and zero-height polylines differently. OfficeIMO retains those attributes and projects their declared geometry; these cases are not qualified for identical LibreOffice appearance. See the [fixture evidence](../OfficeIMO.OpenDocument.Tests/Fixtures/Drawing/README.md).

Unknown drawing XML and package entries remain preserved during targeted editing. Image projection reports unsupported encodings, rasters outside the shared PDF transcode limit and SVGs outside the shared renderable profile; supported captions remain available when image paint is omitted. Partial SVG content receives its own loss mapping. Flat XML also retains bounded StarView (`SVM`) image previews; these bytes are opaque and are not rendered. Embedded object conversion follows the existing flat-XML loss report.

### Draw master artwork

`page.ToDrawing()` includes supported shapes from the referenced master before painting the page's own shapes. Master graphic, paragraph, inline and list styles resolve in `styles.xml`; page content resolves in `content.xml`. Names may overlap between those parts. Nested geometry, embedded images and master-local connector attachments use the same Drawing projection as page content. `page.Shapes` contains the page's own shapes.

Master artwork follows `page.EffectiveLayers` and the selected screen or print intent. Presentation-only visibility flags remain reported as unsupported effects; they do not hide Draw artwork. Page/master cloning and page import preserve artwork, style dependencies and resources. `page.MasterShapes` exposes the referenced master's artwork through the same `OdgShapes` and `OdgShape` APIs as page content. Add, remove, reorder and edit supported shapes, groups, geometry, images, connectors and native text there. Edits affect every page referencing that master. Connector attachments stay within one page or one master; page shapes cannot attach to master shapes. Removing a connected master shape detaches surviving connectors while retaining their current endpoint positions.

For independent artwork, clone the master and assign its name before editing. Graphic and text formatting detach shared automatic styles on first write; image bytes remain owned by the document and shared by clones. Unknown native drawing XML remains preserved, and unsupported geometry or text content keeps its existing editing and conversion limits.

```csharp
var drawing = OdgDocument.Create();
var page = drawing.AddPage("Overview");
var shared = page.MasterShapes.AddTextBox(
    new OdfRect(OdfLength.Points(24), OdfLength.Points(24),
        OdfLength.Points(200), OdfLength.Points(40)), "Shared heading");
shared.FontSize = OdfLength.Points(18);
var copy = drawing.ClonePage(0, "Review");
copy.MasterPageName = drawing.CloneMasterPage(page.MasterPageName, "ReviewMaster");
copy.MasterShapes[0].Text = "Review heading";
drawing.Save("master-artwork.odg");
drawing.SaveFlatXml("master-artwork.fodg");
```

Enhanced geometry and backgrounds outside their documented profiles retain separate loss diagnostics. Strict conversion rejects visible unsupported master content. The independent Draw master fixture contains supported text, rectangles, ellipses and a diamond alongside formula-driven artwork, master bitmaps and a page square gradient; its projected subset does not establish complete page appearance. Native layer resave/printing and exact text metrics retain the qualification limits described above.

### Draw page backgrounds

Use `page.Background` for page-local paint and `page.MasterBackground` for paint shared by every page referencing the master. Both expose `FillColor`, `FillGradientName`, `FillOpacity`, `GradientStepCount`, `SetBitmap(bytes, fileName)` and `GetBitmapBytes()`. `FillMode` and `FillImageName` describe the resolved native declarations in that owner's style chain. Page-local edits use copy-on-write for shared styles. Master edits affect all referencing pages; clone the master and assign its name before independent master edits.

```csharp
var drawing = OdgDocument.Create();
var page = drawing.AddPage("Overview");
page.MasterBackground.FillColor = OdfColor.Parse("#F8FBFF");
page.MasterBackgroundSize = OdgBackgroundSize.Full;
page.Background.FillColor = OdfColor.Parse("#DDEEFF");
var copy = drawing.ClonePage(0, "Detail");
copy.Background.UseNoFill(); // Uses the shared master paint.
copy.MasterPageName = drawing.CloneMasterPage(page.MasterPageName, "DetailMaster");
copy.MasterBackground.SetBitmap(File.ReadAllBytes("background.png"));
copy.MasterBackground.FillOpacity = 0.5;
drawing.Save("backgrounds.odg");
drawing.SaveFlatXml("backgrounds.fodg");
```

`UseNoFill()` selects native no-fill while retaining inactive declarations. A null color or gradient binding selects no-fill and clears that local value. Page no-fill retains master paint; master no-fill removes master paint. Null opacity or band count removes the local override and exposes inherited values. Uniform opacity edits reject active opacity gradients. Gradient bindings use existing common `Styles.Gradients` definitions; definition edits remain shared. `SetBitmap` accepts complete rasters inside the shared drawing/PDF profile, snapshots and deduplicates embedded bytes, and creates an independent fill-image binding. It selects stretch mode without changing opacity or paint area. Linked content is never fetched.

`page.ToDrawing()` paints solid, native two-color gradient and stretched embedded raster backgrounds beneath master artwork and page shapes. Drawing-page styles resolve in their owning parts and through named parents. Page paint overrides the master; page `fill="none"` retains the master fill in the qualified native Draw profile. `MasterBackgroundSize` selects the full page or the area inside the master's absolute nonnegative margins; the default is `Border`, and this area also applies to page-local fill. A conflicting imported page-area declaration and presentation-only visibility flags remain preserved with unsupported-effect diagnostics.

Uniform `draw:opacity` uses the existing Drawing paint opacity. Stretched bitmaps in the shared PNG, JPEG, GIF, BMP, TIFF and WebP raster profile retain source bytes and use the shared pixel-preserving image mode for SVG, raster and PDF output. In stretch mode, ODF's source-size and tile-position declarations are inactive. Package references are resolved without fetching linked content. ODG/FODG saves retain referenced common fill-image bytes, including bounded SVG and recognized opaque image resources outside the projection profile; FODG embeds the payload without adding an unsupported `draw:mime-type` attribute to the definition. Ordinary SVG images keep their existing safety normalization during flat import.

Linear, axial and radial backgrounds resolve common `Styles.Gradients` definitions and reuse the [shape-gradient profile](#create-and-edit-draw-documents), including angles, borders, color intensities, radial centers inside the paint area and uniform opacity. Automatic native banding uses continuous interpolation and produces an approximation mapping; `ThrowOnAnyLoss` rejects it.

Ellipsoid, square and rectangular fields, fixed bands, extended/SVG gradient definitions, off-area radial centers, singular axial borders, hatches, bitmap repeat/scale modes, opacity gradients, vector backgrounds and relative, negative or empty border areas remain preserved with omission diagnostics. The native authoring model can bind two-color gradient kinds and fixed bands outside that projection profile; the conversion report remains the boundary. Missing, malformed and ambiguous fill-image or gradient definitions are reported. Native PDF controls and ODG resaves qualify solid and PNG backgrounds, opacity, master inheritance, page color overrides and full/border paint areas from both containers. Native PDF controls also qualify the documented linear, axial and radial background profile in both containers. Public authoring controls cover page/master clone isolation, bitmap replacement and cross-document background imports; a targeted edit of the independent master fixture retains its separate unsupported-artwork diagnostics. Broader image formats, producers and combined appearance remain outside this native qualification; see the [fixture evidence](../OfficeIMO.OpenDocument.Tests/Fixtures/Drawing/README.md).

### Draw enhanced geometry

Imported `draw:custom-shape` elements project an explicit literal `draw:enhanced-path` with `M`, `L`, `C`, `Q` and `Z` commands in one paint set terminated by `N`. Multiple subpaths in that set use even-odd fill. One full `U` ellipse is supported, optionally closed, with positive radii and start/end angles exactly one turn apart. Its starting angle is a radial vector; the full ellipse draws clockwise. Named `draw:type` values do not replace missing or unsupported paths. These rules follow the [ODF enhanced-path contract](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/part3-schema/OpenDocument-v1.4-os-part3-schema.html).

The explicit integer `svg:viewBox` supplies the local canvas. Geometry scales to the existing `Bounds`; `draw:mirror-horizontal` and `draw:mirror-vertical` reflect it within those bounds. Imported page shapes retain their enhanced XML when their bounds, paint or ordinary text properties are edited and saved as ODG or FODG. Projection uses the same profile for supported master shapes.

Coordinates and curve controls must fit the view box for shared SVG, raster and PDF export. Paths retain the limits of 1 MiB of text and 20,000 normalized commands. Equations/modifier references, partial arcs, other commands, multiple independently painted sets, stretch points and extrusion remain preserved with omission diagnostics. Enhanced text areas, text rotation and text-on-path declarations remain reported as unsupported; supported captions use the ordinary shape frame. There is no public enhanced-path authoring API.

### Draw line and connector labels

`ToDrawing` projects complete unwrapped paragraphs on straight lines and supported saved connector routes containing line segments, quadratic or cubic curves, or a mixture of them. Line captions follow the directed segment, including reversed lines; connector captions remain horizontal and use the route's actual bounds, including bends and curve extrema outside the endpoint rectangle. Curve control points can lie outside those bounds and do not define the caption anchor. Paragraph styles, font sizes and emphasis use the shared text-layout owner. Labels can extend beyond a narrow route or the page edge. Raster paint canvases include horizontal glyph overhang and ink overflowing condensed line spacing while retaining the logical caption frame and anchor.

```csharp
OdgShape connector = page.Shapes.AddConnector(new OfficePoint(40, 80), new OfficePoint(180, 80));
OdfTextParagraph caption = connector.AddParagraph("Approved");
caption.FontFamily = "Arial";
caption.FontSize = OdfLength.Points(12);
caption.TextAlign = "center";
caption.Color = OdfColor.Parse("#000000");
var projected = page.ToDrawing();
string svg = OfficeDrawingSvgExporter.ToSvg(projected.Value);
```

The qualified placement profile uses identity or translation transforms and rotations of one detached straight segment, left-to-right horizontal writing, paragraph alignment and no wrapping. Rotations materialize caption anchors from transformed endpoints without scaling the declared fonts. Line captions follow the rotated directed segment; connector captions remain horizontal around its bounds. Native controls cover positive and negative 30-degree rotations, 90 and 180 degrees with translation, and an equivalent rotation matrix. An omitted wrapping option uses unwrapped projection with an approximation diagnostic. A declared `fo:wrap-option="wrap"` retains the complete caption in unwrapped projection and reports `label-wrapping` as unsupported; `ThrowOnSkippedOrUnsupported` rejects it. The tested native line and connector controls also remain unwrapped under this declaration. Width-constrained wrapping is outside the profile; use explicit paragraph or line breaks for multiline captions. Scaled, reflected or sheared captions and rotated attached/curved/multisegment routes produce explicit losses while retaining their native XML and line geometry. Graphic and paragraph writing directions, including page-layout inheritance, are reported when unsupported.

Captions retain [list labels](#draw-list-labels) and [measured body tabs](#draw-tab-stops) through the shared paragraph layout. Frame measurement includes marker fonts, spacing and tab fields before centering the caption on its route, including short lines and empty item bodies with visible markers. Native controls qualify left-aligned paragraphs with bullets, decimal labels, and left/center/right/comma tab stops on straight and diagonal lines and straight and curved connectors. Label sizes of `100%` retain the paragraph size in these controls; the native absolute-size limitation described above still applies. Field and numeric-label origins agree within 0.2 points along the caption direction; exact font metrics and vertical baselines remain unqualified. Unsupported list/tab styles retain their existing conversion mappings.

Explicit side padding and top/middle/bottom alignment use the route's inset anchor before the unwrapped caption frame grows to fit its text. On positive-sized routes, excessive padding distances are reduced equally, with each side stopping at zero. Inverted inset endpoints are normalized; zero-width routes also bypass vertical reduction. Consequently, top/bottom captions on a zero-height line can extend across the line instead of moving further away from it. The declared padding remains on the projected frame and in native XML. Native controls cover zero-height and diagonal lines, horizontal, vertical, narrow and short straight connectors, and a tall curved connector, with centered two-paragraph captions and zero, symmetric and asymmetric side padding. Relative caption-origin changes agree within 0.2 points; exact baselines, padding shorthand, other producers and broader combined list/tab/rotation inset layouts remain outside this qualification.

For line and connector captions, `draw:textarea-horizontal-align` positions the whole paragraph block independently of each paragraph's `fo:text-align`. The default `justify` spans the inset route width; `left`, `center` and `right` position the measured intrinsic block at that side or center. Text wider than the route can extend beyond it. Native controls cover all four area values with left/center/right paragraphs, explicit Arial 10-point fonts, middle vertical alignment and zero or asymmetric padding on horizontal and diagonal lines and horizontal, curved, narrow and vertical connectors. Visible caption origins agree within 0.3 points along the caption direction. Exact baselines, other fonts/producers and broader combined list/tab/rotation area layouts remain unqualified.

The area and inset profile also qualifies left-aligned bullet/decimal captions, left/center/right/comma tabs, and numbered items containing comma-tab fields on horizontal connectors, diagonal lines and curved connectors. These controls use explicit Arial 10-point fonts, `100%` list-label sizes, middle vertical alignment and asymmetric side padding. Field and numeric-label origins agree within 0.3 points along the caption direction. A list's generated body indent does not shift the body tab columns or inflate the intrinsic caption width. The numbered-tab controls also cover a field overlapping the selected stop, which retains that stop and advances zero. Modern list-spacing modes, non-left tab paragraphs, broader transform-attribute rotation combinations, exact baselines and other fonts/producers remain outside this combined qualification.

Numbered comma-tab captions also qualify positive/negative 30-degree rotations and an equivalent rotation matrix on detached straight lines and connectors, combined with all four horizontal text areas and asymmetric side padding. These controls use left paragraphs, middle vertical alignment, explicit Arial 10-point text, `100%` label sizing and black label color. Field and label starts agree within 0.2 points along the caption direction. ODG/FODG reopening retains the transform declarations; native resaving materializes them into route coordinates. The tested native exporter quantizes the line-caption angle from the rotation string to 29.9 degrees while the equivalent matrix exports 30 degrees, so identical native glyph angles and pixels are not qualified. Wrapped captions, broader list modes, other fonts/producers and rotated attached or curved routes remain outside this profile.

On a zero-width route, the tested native producer expands left/right text areas and noncentered justified paragraphs far off-page. OfficeIMO retains those captions near the route with an explicit `label-degenerate-area` approximation; `ThrowOnAnyLoss` rejects it. Centered areas and centered justified paragraphs retain visible native placement in the tested controls.

Free connector translation baking retains captions through ODG/FODG reopening, including curved routes and group-child translation. The tested LibreOffice build retains the baked curve and caption but reroutes a curve that still carries a transform attribute. Use `BakeConnectorTransform()` or `TransformChildren()` to materialize a supported translation before native editing when retaining the saved curve matters. In vertical connector controls the tested native build places left-aligned captions far off-page; centered paragraphs retain the caption, while OfficeIMO retains both at the route. Native font and text-color defaults and shared font metrics can differ, and a native application can recalculate a connector route. Projection reports these placement approximations; exact typography and later font-provider replacement require broader native qualification.

Run the [Draw example](../OfficeIMO.Examples/OpenDocument/DrawDocument.cs) with `OfficeIMO.Examples --opendocument` to write ODG, FODG, SVG, and a composed PDF.

### Draw text areas

Use the shared shape properties to edit native text-frame declarations:

```csharp
OdgShape label = page.Shapes.AddTextBox(
    OdfRect.FromCentimeters(2, 2, 8, 3), "Native text frame");
label.TextAreaAlignment = OfficeTextAreaAlignment.Right;
label.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Middle;
label.WrapText = true;
label.TextPadding = OdfInsets.FromCentimeters(0.1, 0.2, 0.1, 0.3);
```

These nullable properties resolve automatic, named, parent and default graphic styles. Assigning `null` removes the local declaration and exposes inheritance. `TextPadding` uses top, right, bottom, left order; each side resolves its explicit declaration or `fo:padding` shorthand at each style level. Missing sides are zero, and no declarations returns `null`. Setting padding validates all four lengths before changing XML, writes explicit sides and removes the local shorthand. Edits isolate shared styles and stay in the shape's owning package part. ODG/FODG saves, page/master cloning and page import retain the declarations. The same properties on `OdpTextBox` edit native ODP/FODP declarations; they do not extend the PowerPoint adapter's rendering profile.

`OfficeTextAreaAlignment.FullWidth` writes native horizontal `justify`. Native vertical `OdfTextAreaVerticalAlignment.Justify` and pixel padding remain editable and preserved, but shared Draw projection reports their text as skipped and strict loss policies reject it. Unknown imported horizontal, vertical or wrapping values remain in XML; typed getters throw `NotSupportedException` without changing the source. Pixel padding has no implicit DPI conversion.

Controlled Draw resaves retain the typed horizontal and vertical alignment, wrapping and four-sided padding declarations across ODG and FODG. The top/middle/bottom wrapped text-box controls retain the sampled word-line breaks; font placement differs by up to 2.55 points. With auto-growth disabled, the tested LibreOfficeDev exporter also wraps boxes declared `no-wrap` while retaining that declaration on resave. Shared projection honors the declared wrapping option and reports approximate native layout; fixed-width native no-wrap appearance remains unqualified.

Rectangle labels and text-box frames also project explicit `left`, `center` and `right` horizontal areas independently of paragraph alignment. Core measures the intrinsic block at render time, including paragraph margins and tab fields, then anchors it inside the padded frame. Shorter hard-break lines align within that block. An omitted or `justify` area uses the full padded frame width. The conversion reports `text-area-alignment` as approximated; it preserves the native declarations and passes `ThrowOnSkippedOrUnsupported` within the supported text profile.

The qualified native profiles distinguish wrapping by shape kind. Text-box frames honor `fo:wrap-option="wrap"`. Ordinary rectangle labels with intrinsic areas remain unwrapped in the tested native exporter despite that declaration; their projection reports `text-area-wrapping` as approximated and retains explicit hard line breaks. Fixed-frame shrinking rectangles honor wrapping in the qualified native profile and shared projection. An oversized intrinsic block can extend outside the shape, including across its left edge when centered or right aligned. The shared Core API retains its own explicit wrapping option.

Controlled native resaves cover all nine area/paragraph left/center/right combinations on rectangles and text boxes, with hard breaks, body tabs, asymmetric frame padding, paragraph margins and long text. Hard-line text and tab-field starts agree with shared PDF within 0.06 points. The wrapped text-box controls retain the same word breaks, but glyph starts differ by up to 3.34 points; native trailing spaces and font layout remain outside exact placement qualification. These controls establish the tested producer profile rather than independent authorship or exact font/pixel fidelity. Non-left indented tab paragraphs in an intrinsic area report `text-area-tab-indent` as unsupported; strict conversion rejects them. Native rectangle first-line controls can place their fields far off-page, and text-box placement differs from the shared indent profile. Other shape kinds, auto-sizing, complex scripts, combined list layouts and additional producers remain outside this qualification and retain applicable loss reports.

### Draw text fitting

Shrink a caption within its saved text-box frame or rectangle bounds:

```csharp
OdgShape caption = page.Shapes.AddTextBox(
    OdfRect.FromCentimeters(2, 2, 8, 3), "A caption that needs to fit its frame");
caption.Paragraphs[0].FontSize = OdfLength.Points(18);
caption.TextFitMode = OdfTextFitMode.ShrinkToFit;
OdfConversionResult<OfficeDrawing> projected = page.ToDrawing(
    OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
```

`TextFitMode` edits the complete native fitting mode: `None`, `Stretch` or `ShrinkToFit`. The nearest graphic style declaring either `draw:fit-to-size` or `style:shrink-to-fit` owns that mode. Setters write both canonical booleans on an isolated style; `null` removes both local declarations and exposes inheritance. ODG/FODG and native ODP/FODP editing retain these modes. Unknown values, legacy fitting tokens and conflicting active declarations remain in XML; the typed getter throws without changing the document.

Shared SVG, raster and PDF layout approximates `ShrinkToFit` on fixed text-box frames and rectangle labels. It keeps frame bounds, padding and paragraph placement. Run sizes scale together, with a one-point minimum for scaled runs. The font scale stops when the largest font reaches six points, or its original size when smaller. Absolute paragraph spacing, indentation, tabs and line heights stay fixed. Native metrics and later font-provider changes can alter fitting. Stretching, fitting combined with active auto-growth, fitted line captions and fitting on other shape kinds retain unsupported mappings; strict projection rejects them.

Controlled native resaves retain the three canonical fitting modes and the nearest-style override rule in Draw, plus shrinking declarations in ODP/FODP. Six fixed-frame caption controls cover rectangles and text boxes with full-width, left and right areas, mixed run sizes, asymmetric padding, wrapping and vertical placement. Shared and native PDF exports retain every sampled caption marker; word breaks, font scale and glyph placement differ. These controls qualify the tested producer profile rather than independent authorship or exact typography.

Use `ToDrawing(renderingProfile, lossPolicy)` or `ToDrawings(renderingProfile, lossPolicy)` to supply an `OfficeRenderingProfile` before projection. Its fonts and shaping settings guide fitting and clipping checks and remain on the returned drawings. ODG-to-PDF conversion uses the final `PdfOptions` font selection for those checks.

Fixed text frames and line/connector captions check shared layout and placed glyph, decoration, background and tab-leader paint for clipping with the conversion's current font resources. `text-clipped` reports an unsupported mapping when frame bounds, paragraph margins, padding, tabs or shared layout limits omit content or cut off text paint, including ordinary text without fitting. `ThrowOnSkippedOrUnsupported` rejects that projection. `ReportOnly` retains the supported preview and the original native text; adjust the frame, margins, padding, line spacing or font size and inspect the report before accepting output. A later font-provider replacement can change the measured result.

PDF clipping checks use the selected conversion metrics. Unembedded standard fonts use Adobe glyph bounds; a reader's substitute can paint outside them. Configure licensed embedded fonts through the PDF adapter's `PdfOptions` when predictable font selection is required. Reader substitution and exact native typography remain outside this qualification.

### Draw text-box sizing

```csharp
OdgShape box = page.Shapes.AddTextBox(
    OdfRect.FromCentimeters(2, 2, 8, 2), "A caption that can grow vertically");
box.AutoGrowHeight = true;
box.AutoGrowWidth = false;
box.WrapText = true;
box.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
box.TextFitMode = OdfTextFitMode.None;
OfficeDrawing drawing = page.ToDrawing(
    OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
```

`AutoGrowHeight` and `AutoGrowWidth` resolve independently through graphic-style inheritance. Setters isolate a shared style; `null` removes the local declaration. Omitted flags remain unspecified because producer defaults differ. Unknown boolean declarations stay in XML, and the typed getter throws without mutation.

`TextBoxMinimumHeight`, `TextBoxMinimumWidth`, `TextBoxMaximumHeight` and `TextBoxMaximumWidth` edit attributes on a frame's direct `draw:text-box`. They are instance constraints, separate from graphic-style creation defaults and saved `svg:width`/`svg:height`. Setters accept finite nonnegative absolute ODF lengths or percentages, preserve lexical values, and reject other containers or invalid assignments before changing XML. `null` removes an instance attribute. Editing retains constraint combinations as authored. ODG/FODG and ODP/FODP round-trip these controls.

Set an instance floor and growth cap using the same unit:

```csharp
box.TextBoxMinimumHeight = OdfLength.Points(24);
box.TextBoxMaximumHeight = OdfLength.Points(96);
box.AutoGrowWidth = true; // Grow width first, then height if wrapping remains.
box.TextBoxMinimumWidth = OdfLength.Points(120);
box.TextBoxMaximumWidth = OdfLength.Points(240);
```

Shared projection resolves absolute instance constraints on ordinary text-box frames when both growth flags are explicitly enabled or disabled and geometry is independent of other frames. A minimum in `cm`, `mm`, `in`, `pt` or `pc` replaces the corresponding saved frame dimension, even when it is smaller or zero. A maximum requires a minimum in the same lexical unit and cannot be smaller; it caps growth rather than setting the frame's size. Graphic-style creation defaults do not resize saved instances. Invalid, orphaned, relative or pixel pairs retain unsupported mappings and leave both saved dimensions in use. Text and frame paint use the resolved dimensions; zero-area frame paint is omitted and a nonempty body reports clipping.

Growth additionally requires horizontal text, top placement and no active fitting. Width growth measures complete unwrapped paragraphs, retaining hard breaks, paragraph insets, labels, tabs and horizontal paint bounds. It reserves horizontal ink and decoration insets before alignment and wrapping, using the same placement in SVG, raster and PDF. Capped no-wrap text retains its paragraph and text-area anchor. Wrapping rechecks newly exposed ink; if bounded layout cannot contain it, strict conversion reports clipping. With both axes enabled, width grows first; an instance maximum can stop width growth and leave wrapping for the height pass. Height growth measures paragraphs and complete text paint at that resolved width, respecting the wrapping option. Both passes use the conversion's font resources.

Each instance minimum replaces its saved dimension; without a minimum, the saved dimension remains the floor. A maximum or a fixed height can leave content clipped or horizontal paint outside the padded frame: `text-clipped` then rejects strict projection. A capped no-wrap body remains unwrapped and can overhang in report-only output. The original ODF geometry remains unchanged.

Chained text-box sources and incoming targets retain `text-chain-flow` unsupported mappings, including empty targets, nested frames, masters and links from other pages; projection does not redistribute their stories. Unspecified growth axes, relative geometry, other anchor or flow modes and fitting combined with growth retain unsupported mappings. Growth keeps page dimensions unchanged and can place content outside its page. Frame clipping and shared resource limits still apply.

Straight connector attachments resolve against constrained and grown frame bounds, including explicit, standard and automatic glue points and transformed group children. Public endpoint getters and editing routes continue to use saved ODF geometry. A cached route that no longer matches the projected attachment points reports loss; projection does not reroute around obstacles.

Controlled native inputs cover omitted and explicit growth flags, instance minimum/maximum heights, style minimum-height defaults, and text append/removal. The tested producer grows an unconstrained 48-point frame to approximately 259.34 points at its fixed 120-point width and retains the grown saved height after text removal. It drops instance constraints and does not enforce the authored limits. Both-axis growth can place the end marker off-page. These observations qualify that producer profile; native constraint retention, exact typography and additional independent producers remain open.

## Edit without flattening the package

Typed objects remain backed by the source XML. A targeted edit rewrites its owning XML part while untouched package entries keep their original bytes.

```csharp
OdtDocument document = OdtDocument.Load("input.odt");
document.Paragraphs[0].Text = "Updated text";
OdfSaveResult result = document.Save("output.odt", new OdfSaveOptions {
    CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
});

IReadOnlyList<string> rewritten = result.Report.RewrittenEntries;
IReadOnlyList<string> lossy = result.Report.LossyEntries;
```

`Serialize`, `ToBytes`, `ToStream`, `SaveCopy`, and `SaveCopyAsync` produce independent outputs without accepting pending edits or changing the source version, signatures, path, or encryption state. Stream saves also remain independent when a source path is attached. Use `Save` or `SaveAsync` with a path to accept changes and associate the document with that destination.

A failed or canceled write keeps pending edits, the source version, and signature and encryption state intact. Existing typed wrappers remain usable after a successful save.

New documents use ODF 1.4. Set `OdfCompatibilityProfile.Odf13` when the output needs the ODF 1.3 schema and compatibility profile.
Distinct first-page master-page headers and footers require ODF 1.4. Saving those stories with the ODF 1.3 profile, or preserving an older source version, fails before writing an invalid package.

## Encrypt and decrypt ODF packages

Password encryption is format-owned and does not require `OfficeIMO.Security`:

```csharp
OdtDocument document = OdtDocument.Load("protected.odt", new OdfLoadOptions {
    Password = password
});

document.AddParagraph("Updated while decrypted in memory.");
document.Save("protected-updated.odt", new OdfSaveOptions {
    Encryption = new OdfEncryptionOptions {
        Password = newPassword
    }
});
```

The password is UTF-8, used only for the current load or save, and is not retained. Input accepts 10,000 through 10,000,000 PBKDF2 iterations per entry and preflights the complete manifest against `OdfLoadOptions.MaxTotalKdfIterations` (10,000,000 by default) before deriving any entry key. Output uses AES-256-CBC, a SHA-256 password start key, per-entry PBKDF2-HMAC-SHA1 with 100,000 iterations by default, and SHA-256/1K checksums. Each encrypted entry receives fresh salt and initialization-vector material.

Encrypted input fails with a classified `OdfEncryptedPackageException` when a password is missing or incorrect, the profile is unsupported, metadata is malformed, or decrypted content exceeds configured limits. Saving an encrypted source without `OdfSaveOptions.Encryption` also fails so protection is not removed accidentally. To write plaintext intentionally, set `EncryptionHandling = OdfEncryptionHandling.Remove`.

## Supported editing surface

| Area | Current support |
| --- | --- |
| Package | Bounded ZIP/XML loading, direct reading of seekable package streams, manifest updates, deterministic output, metadata, atomic path saves, flat XML projection with loss reporting, unknown-entry preservation |
| ODT | Paragraphs, headings, ordered inline text/span/link/image/bookmark syntax, page number/count and date/time fields with cached display text, whitespace controls, common text and paragraph styles, lists, tables, sections, page layout, default/first/left master-page headers and footers, page breaks, images, paragraph insertion/deletion tracking |
| ODS | Sparse repeated rows/cells, typed values, OpenFormula text and cached values, bounded formula evaluation/recalculation, styles and data formats, merges, row/column sizing and visibility, sheet order, typed named ranges, annotations, typed scalar/list validations and messages, links, print ranges |
| ODP | Slide order and visibility, page size, masters/layouts and presentation classes, ordered inline text/run/link syntax, common run styles, lists, rectangles, ellipses, lines, groups, transforms, images and crop, tables, speaker notes, backgrounds, transitions, basic shape animations |
| ODG | Drawing pages and dimensions, shape order, rectangles and rounded corners, ellipses, paths, polygons, polylines, lines, groups, text boxes, images, plain text, font and graphic styles, glue points and attached connectors, affine transforms, screen/print layers, preservation of unmodeled drawing XML |
| Inspection | Annotations, tracked changes, extension namespaces, scripts, event listeners, external links, embedded objects, formulas, validations, transitions, animations, encryption, and signatures |

Unknown XML, vendor extensions, scripts, embedded content, and unsupported drawing features are preserved when their owning part is not replaced. The library never executes scripts, macros, event listeners, embedded objects, or external links. Formula evaluation is a bounded, side-effect-free parser for the documented local subset; it does not execute active content or fetch data.

`OdfCapabilityCatalog.Advanced` provides stable capability IDs and distinguishes editable subsets, preserved content, inspection, and detected-but-unsupported features.

## Content provenance

`OdfDocument.InspectProvenance("input.odt")` reports C2PA and AI-specific IPTC metadata in ODF packages and supported embedded images. `OdfDocument.RemoveProvenance("input.odt", "clean.odt")` performs a bounded package rewrite while preserving the required uncompressed, first `mimetype` entry. Signed-package mutation is blocked unless removal of invalidated ODF signature entries is requested explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`OdfDocument.InspectContentSafety(...)` covers ODT, ODS, and ODP native hidden fields and containers, concealed stored values/formulas, resolved tiny/transparent/low-contrast styles, zero geometry, notes, annotations, alternative descriptions, and Unicode evidence. `OdfDocument.RemoveSelectedContent(...)` removes exact reviewed text segments or exact Unicode ranges inside stored attributes through the preservation-aware package writer. Encrypted-source cleanup and implicit signature invalidation are rejected.

## Explicit boundaries

- Formula evaluation covers arithmetic, comparisons, concatenation, cell/range references, Boolean constants, and common aggregate/math functions. OpenFormula prefix minus binds before powers, and chained powers associate from the left. `OdsFormulaEvaluationOptions` bounds operations, dependencies, ranges, formula length, individual text results, and cumulative text work; concatenation limits are checked before allocation. External data, volatile functions, matrix formulas, and the complete OpenFormula language are not included.
- Typed validation syntax covers explicit lists and scalar whole-number, decimal, and text-length comparisons. Other valid ODF conditions remain preserved text and are reported by conversions that cannot map them exactly.
- Ordered ODT/ODP inline syntax types text, nested spans/runs, and hyperlinks. ODT also types inline images and bookmark markers. Unsupported inline elements remain `Other` nodes and conversion reports their approximation.
- Tracked-change editing covers paragraph insertions and deletions. Arbitrary inline merges and conflict resolution remain preservation-oriented.
- Animation editing covers basic shape-attribute effects and fade-in timing. Advanced timing trees are preserved when untouched.
- Password-encrypted packages using the documented AES-256-CBC profile can be opened and written. Legacy Blowfish and other unsupported profiles fail before content is exposed.
- Changed signed packages fail by default because saving would invalidate signatures. An explicit save option can remove invalidated signature entries.
- The bounded OfficeIMO XML package-manifest signature profile can be created and validated through an explicit `IOfficeSecurityProvider`. Arbitrary producer-specific signature profiles remain inspection or preservation oriented.
- ODS exposes embedded chart names, types, titles, source ranges, frame positions, and supported per-point styles through `OdsSheet.Charts`. `OdsSheet.AddChart` creates column, bar, line, pie, or doughnut charts linked to existing one-dimensional ODS cell ranges of up to 4,096 points; pie uses one series, while the other forms accept up to sixteen. `OdsChartSeries.WithPointStyles` applies solid fills, hatches, and outlines to individual column, bar, pie, and doughnut points. Styled line points require visible symbols, which native authoring does not yet emit; that combination is rejected. Unsupported imported styling remains in package XML; editing imported charts remains outside the current surface. Data pilot tables expose their source and target ranges and field orientations; advanced imported pivot settings remain preservation-oriented.
- Flat XML variants (`.fodt`, `.fods`, `.fodp`, `.fodg`) can be opened and written, including embedded raster images. Exotic embedded objects and package-only features may not project losslessly.
- `OdsSheet.Merge` rejects merges above its default 100,000-cell materialization limit. Use the overload with an explicit lower limit when processing untrusted dimensions.
- Unknown package entries and extension XML are always preserved by package editing. Explicit format conversion and flat XML projection report content they cannot carry through `OdfConversionReport` and `OdfSaveReport.LossyEntries`.

The package targets `netstandard2.0`, `net8.0`, and `net10.0`, plus `net472` on Windows. CI checks generated ODF 1.3 and 1.4 XML against pinned OASIS Relax NG schemas, then opens and resaves the generated packages with the runner's reported LibreOffice version.

Interoperability coverage includes ODT, ODS, and ODP files from LibreOffice and Microsoft Office, plus an externally verified Google Docs ODT export. These files exercise styles, formulas, drawings, embedded content, and preservation of unknown package entries. A separate hash-pinned LibreOffice fixture covers password encryption, including OfficeIMO reading LibreOffice output and LibreOffice reading OfficeIMO output. See the [producer manifest](../OfficeIMO.OpenDocument.Tests/Fixtures/producer-manifest.json) and [encryption manifest](../OfficeIMO.OpenDocument.Tests/Fixtures/Encryption/producer-manifest.json) for exact producer versions, hashes, and evidence.

Draw interoperability covers [LibreOffice FODG and ODG fixtures](../OfficeIMO.OpenDocument.Tests/Fixtures/Drawing/README.md), package/flat round trips, and a generated ODG opened and rendered by LibreOffice. This evidence does not qualify arbitrary Draw layout or every producer.

## Dependency footprint

- **External:** None; no OpenDocument SDK and no LibreOffice process.
- **OfficeIMO:** `OfficeIMO.Core`. ODT/ODS/ODP/ODG parsing, models, preservation, inspection, and writing are first-party.
- **Security:** ODF password encryption/decryption is first-party and dependency-free. Signature carriers are detected and changed signed packages fail safely without a cryptographic dependency. `OdfDocument.SignPackage(...)` and `ValidatePackageSignatures(...)` use an explicit provider for the bounded OfficeIMO XML package-manifest profile; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 3 | 0 | 0 | 0 | 0 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 1 |
| Edit | 2 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 1 |
| Remove | 3 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
