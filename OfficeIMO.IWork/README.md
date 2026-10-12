# OfficeIMO.IWork - Apple iWork reading, native creation and preservation for .NET

`OfficeIMO.IWork` reads modern Apple Pages, Numbers, and Keynote packages, creates native documents and saves captured sources with explicit preservation boundaries, without running iWork or executing embedded content. It owns ZIP, directory-bundle, nested `Index.zip`, Snappy-framed IWA, protobuf-envelope, package-resource, and source-record preservation. Word, Excel, and PowerPoint remain the owners of editable destination documents.

## Reference from a source checkout

For source-based development, reference the bounded reader and the opt-in adapter you need:

```xml
<ItemGroup>
  <ProjectReference Include="../OfficeIMO.IWork/OfficeIMO.IWork.csproj" />
  <ProjectReference Include="../OfficeIMO.Excel.IWork/OfficeIMO.Excel.IWork.csproj" />
</ItemGroup>
```

Use `OfficeIMO.Word.IWork` for Pages or `OfficeIMO.PowerPoint.IWork` for Keynote in place of the Excel adapter. Keep all project references on the same checkout so their coordinated API and package contracts stay aligned.

Numbers reconstruction retains bounded `GCD`/`LCM` references such as `GCD(B2,C2,D2)` and `LCM($B$2,$C$2:$D$2,1)` through saved XLSX and operand recalculation. The [native cross-table reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-cross-table-v14.5-decode.json) covers numeric cells and ranges across tables and sheets, preserving canonical table identities and native caches. Direct single-cell references also accept canonical nonnegative integer text such as `"12"`, within the same numeric bound; the [native text reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-integer-text-reference-v14.5-decode.json) records a text edit to `"18"` and recalculation. [Proven blank references](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-blank-reference-v14.5-decode.json) are editable beside qualified numeric operands: direct blanks are ignored, while range blanks supply zero. Text ranges, all-blank arguments, wider mixed types, formula dependencies and larger integer domains retain their cached values with incomplete-expression diagnostics. The [formula contract](../Docs/officeimo.iwork-support-matrix.md#formula-functions) records the qualified domain; other source limitations still determine conversion acceptance.

## Create a native Keynote presentation

```csharp
using OfficeIMO.IWork;

var presentation = IWorkKeynoteDocument.Create(); // 960 × 540 points
var slide = presentation.AddSlide("E6F2FF");
slide.AddText("Zażółć gęślą jaźń\nA😀B — café", 60, 100, 840, 260,
    fontName: "Arial", fontSizePoints: 40, color: "003366");
slide.AddRichText(new[] {
    new IWorkTextRunData("Bold ", bold: true),
    new IWorkTextRunData("then plain"),
    new IWorkTextRunData(" and underlined", underline: true)
}, 60, 400, 840, 80);
presentation.AddSlide(); // blank white slide
var imageSlide = presentation.AddSlide();
imageSlide.AddImage(File.ReadAllBytes("photo.png"), 60, 100, 480, 320,
    accessibilityDescription: "A photograph of the finished project");
presentation.Save("created.key"); // rejects an existing destination
byte[] package = presentation.SaveBytes();
```

The writer generates its own native object graph, protobuf archives, Snappy frames and deterministic ZIP package. It needs no seed file, OOXML conversion or installed Apple application. Slides have opaque six-digit sRGB backgrounds. `AddText` uses one font, size and opaque color throughout. `AddRichText` copies immutable runs with independent fonts, sizes, colors, bold, italic, single underline and single strikethrough. Plain runs clear earlier decoration. Empty ranges are discarded unless the entire box is blank; `Runs` exposes the retained snapshot. Text boxes use left and top alignment, zero internal padding and no automatic fitting. Frames must lie within the canvas. CR and CRLF normalize to LF; other control characters and invalid UTF-16 are rejected. Fonts are requested by name and are not embedded. Wrapping, fallback glyphs and appearance depend on the receiving application's fonts.

`AddImage` validates and copies a complete PNG or JPEG. The full image fills its frame, stretching when the frame has a different aspect ratio. PNG transparency is retained. Text boxes, images and tables use their combined insertion order for stacking. Optional accessibility descriptions support up to 4096 UTF-16 code units. Crops, masks, rotation and image adjustments are not authored.

`AddTable` snapshots a one-based `IWorkTableData` grid with text, finite numbers, booleans and empty cells:

```csharp
var values = IWorkTableData.Create(2, 3)
    .SetText(1, 1, "Product").SetText(1, 2, "Amount").SetText(1, 3, "Active")
    .SetText(2, 1, "Widget").SetNumber(2, 2, 1200.5).SetBoolean(2, 3, true);
slide.AddTable(values, 60, 120, 840, 300, "Values", "Arial", 24, "003366");
```

Each axis supports 1–256 cells. Tables use equal row heights and column widths, a uniform font/size/color, white cell fills, solid black one-point borders, four-point padding and top alignment. Text beginning with `=` remains literal text. Header/footer regions, merges, formulas, comments and per-cell fonts/fills/borders are not authored. Numeric display formats use the [shared format API](#format-numeric-cells-in-new-tables). `MaximumTables` and `MaximumTableCells` bound aggregate tables and grid positions, including unset cells; table names and set text cells count toward the text budget. The [table reference](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-tables-v15.4.json) records independent decoding, native style presets, all-slide open/export and text/numeric edits that survive native save/reopen. The reader and destination adapters retain the [qualified solid/invisible border profile](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-borders), including borders of unset cells.

`IWorkKeynoteWriteOptions` limits slides, total text boxes and retained character runs, UTF-16 text length, image count, combined encoded/RGBA image bytes and package/intermediate archive size. Image limits count each instance, including repeated payloads; Core image validation also bounds each input. Save operations accept cancellation. Encoding and validation finish before stream copying or file staging. Stream saves leave the caller's stream open and can leave partial data after cancellation or an I/O failure during copying. Path saves use the shared atomic file-commit owner; failed staging preserves the existing destination. Use `overwrite: true` to replace an existing file. Creation instances are not thread safe.

The [native creation evidence](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-v15.4.json) records independent graph decoding and Keynote 15.4 open, PDF export, text edit, save and reopen for a three-slide fixture with Arial, Times New Roman and multilingual text. The [rich-run reference](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-rich-v15.4.json) adds Helvetica Neue, character decorations and resets, mixed font sizes/colors and Unicode ranges. The [image reference](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-created-images-v15.4.json) qualifies PNG alpha, JPEG stretching, stacking and accessibility descriptions, with native PPTX/PDF exports and an image edit/save/reopen check. These qualify their recorded fixtures and font environment. Other native versions, fonts/scripts, image transformations, broader table formatting, transitions, builds and editing loaded packages remain outside this evidence.

## Create a native Pages body document

```csharp
using OfficeIMO.IWork;

var document = IWorkPagesDocument.Create(); // US Letter, 72-point body margins
// Pass width/height and four margins to Create for a different page layout.
document.SetHeader(new IWorkTextRunData("Project report", "Arial", 14, "003366", bold: true), distancePoints: 30);
document.SetFooter("Internal copy", distancePoints: 36);
document.AddParagraph("Project report", fontSizePoints: 24);
document.AddRichParagraph(new[] {
    new IWorkTextRunData("Bold 😀", "Arial", 16, "003366", bold: true),
    new IWorkTextRunData(" then plain", "Arial", 12)
});
document.AddImage(System.IO.File.ReadAllBytes("logo.png"), 120, 80, "Company logo");
document.AddTable(IWorkTableData.Create(2, 2)
    .SetText(1, 1, "Item").SetText(1, 2, "Value")
    .SetText(2, 1, "Widget").SetNumber(2, 2, 1200.5),
    "Values", columnWidthPoints: 150, rowHeightPoints: 28);
document.AddParagraph(""); // editable blank paragraph
document.AddParagraph("Relative line spacing", new IWorkPagesParagraphLayout(1.5f, spaceAfterPoints: 6));
document.AddParagraph("Total\t123.45", new IWorkPagesParagraphLayout(tabStops: new[] {
    IWorkTabStop.Create(180, IWorkTabAlignment.Decimal)
}));
var steps = document.CreateNumberedList(); // Arial, 12 points
steps.AddItem("Prepare", restartAt: 7);
steps.AddItem("Check sources", level: 1);
steps.AddItem("Record reference details", level: 2);
document.AddParagraph("An ordinary paragraph preserves the counters.");
steps.AddItem("Review"); // parent 8; the next child starts at 1
steps.AddRichItem(new[] {
    new IWorkTextRunData("Important", "Arial", 12, "003366", bold: true),
    new IWorkTextRunData(" check", "Arial", 12)
}, level: 1);
document.Save("report.pages");
```

The writer creates its own native document, section, theme, style and text-storage graph. Page dimensions and body margins use points; wider pages carry landscape orientation. Ordinary paragraphs are left-aligned with zero indents. `IWorkPagesParagraphLayout` adds positive relative line spacing, nonnegative before/after spacing and sorted, unique left/center/right/decimal tab stops. Single spacing, zero paragraph spacing and a 36-point default tab interval apply by default. Custom layouts use a period as the decimal separator and author no tab leaders.

Rich paragraphs snapshot immutable `IWorkTextRunData` runs. Plain runs explicitly clear decoration, and empty paragraphs retain their font metrics. Use the plain `AddParagraph(text, layout)` overload for TAB characters; public rich-run constructors reject them. Line breaks require separate paragraphs, and invalid UTF-16 is rejected. The receiving application flows the body over pages using its installed fonts; the writer does not measure pagination or embed fonts.

`SetHeader` and `SetFooter` author one left-aligned paragraph in the single section, repeated on every page. Use a string with font/size/color arguments or one immutable `IWorkTextRunData` for decorations. Distances use points from the top or bottom page edge and must be nonnegative and smaller than the corresponding body margin. Text must contain at least one character and cannot contain tabs or line breaks; native wrapping depends on fonts and available width. `ClearHeader` and `ClearFooter` disable their area. Center/right positions, mixed runs, styled zero-length paragraphs, page-number fields and alternate first/even-page templates are not authored.

`AddTable` snapshots a scalar `IWorkTableData` grid with 1–256 rows and columns. Tables follow paragraphs and other tables in `BodyBlocks` order; each occupies its own body paragraph. Equal positive row and column dimensions must fit the page body. Cells share the requested font/size/color, white fill, black one-point grid, wrapping, top alignment and four-point padding. Formula-looking text remains literal. Names are distinct without regard to case; formulas, merges and per-cell styles are not authored. Numeric display formats use the [shared format API](#format-numeric-cells-in-new-tables).

`AddImage` snapshots a validated PNG or JPEG into its own body paragraph. Finite positive width/height values use points and must fit the page body. The complete image fills that frame, including intentional stretching and PNG transparency. Optional accessibility descriptions retain at most 4096 normalized UTF-16 units. Crops, masks, rotation, floating placement, image links and effects are not authored. Images, paragraphs and tables share `BodyBlocks` insertion order.

`CreateNumberedList` returns a document-bound decimal list with period markers by default. Pass `IWorkDecimalListMarkerStyle.Parentheses` or `IWorkDecimalListMarkerStyle.ClosingParenthesis` to select `(1)` or `1)` for the whole list; `MarkerStyle` exposes that choice. `AddItem` and `AddRichItem` append one immutable body paragraph at any native level from 0 through 8, with an optional positive `restartAt` and paragraph layout. Rich items snapshot `IWorkTextRunData` runs whose font families, opaque colors, bold, italic, underline and strikethrough may vary. Retained runs may use different whole-point sizes. The marker takes the first retained run's size and remains plain black in `FontName`, which also supplies the plain-item font. `FontSizePoints` supplies the default body size and the paragraph-relative indentation. Fractional list sizes, bullets and tiered numbering are unsupported. Nested items require an earlier parent at the preceding level; advancing or restarting an ancestor resets all its descendant counters. Plain paragraphs, tables and images preserve counters. Returning to an earlier group restores its authored counters with explicit native starts; later edits before another group do not propagate through those restored starts. Unused groups do not affect saved bytes. `NumberedLists`, `Items`, `ListLevel`, `ListNumber` and `ListRestartNumber` expose the authored structure. The [rich-list reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-rich-numbering-v15.4.json) records independent decoding, Pages open/edit/save/reopen and native/composed PDF checks, with font and export limits.

Mix whole-point sizes and font families within one item:

```csharp
var pages = IWorkPagesDocument.Create();
var numbered = pages.CreateNumberedList("Arial", 12);
numbered.AddRichItem(new[] {
    new IWorkTextRunData("Serif emphasis", "Times New Roman", 18, bold: true),
    new IWorkTextRunData(" monospace", "Courier New", 8, italic: true),
    new IWorkTextRunData(" plain", "Arial", 14)
});
pages.Save("mixed-size-list.pages");
```

`IWorkPagesWriteOptions` bounds used list groups (`MaximumNumberedLists`), paragraphs including header/footer and list/image/table hosts, retained runs including header/footer and attachment markers, image instances (`MaximumImages`), combined encoded/RGBA image bytes (`MaximumEncodedImageBytes`/`MaximumDecodedImageBytes`), aggregate tables (`MaximumTables`) and dense grid positions (`MaximumTableCells`), UTF-16 body/header/footer/cell/name characters including stored terminators, aggregate custom tab stops (`MaximumTabStops`) and encoded package/intermediate archive bytes. Byte, caller-stream and atomic path saves accept cancellation and use the same package and file-commit owners as Keynote creation. Encoding completes before stream copying or file staging. Stream saves leave the caller's stream open; copy failure or cancellation can leave partial stream output. File staging failure preserves an existing destination.

The [Pages creation reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-v15.4.json) records independent decoding of the complete graph and native Pages 15.4 open, edit, save, close/reopen and PDF/DOCX export. Both PDF pages are inspected. Native Pages/PDF and strict saved OfficeIMO DOCX retain the recorded rich formatting. Apple's DOCX export agrees on text and page settings but changes selected font runs and bold/italic overrides; equivalent DOCX typography is outside that reference.

The [paragraph-layout reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-layout-v15.4.json) adds 1.25/1.5/2 relative spacing, before/after spacing, all four tab alignments, a blank paragraph and formatting resets. These survive native edit/save/reopen and agree with Apple's DOCX paragraph metrics. Strict OfficeIMO DOCX retains the period separator, but its Pages 15.4 reimport places `123.45` about 15 points left of the native source; the other 74 word bounds agree. Native DOCX decimal-tab rendering equivalence remains unqualified. The [inline-table reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-tables-v15.4.json) records independent graph decoding, native open, numeric edit, save/reopen and PDF/DOCX export for mixed paragraphs, two scalar tables and an empty grid. A table-only document also opens natively. Strict OfficeIMO DOCX retains body order, typed text, inherited cell margins and fixed dimensions. Native typography, localized display and receiving-application row-height equivalence remain separate limits. The [inline-image reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-images-v15.4.json) qualifies three PNG/JPEG images around paragraphs and a table, plus an image-only body. Pages 15.4 opens them without repair; an image resize survives native save/reopen and agrees with native DOCX frames. Strict OfficeIMO DOCX and complete one-page PDFs retain image dimensions, descriptions, payload colors and alpha. The reference records native font, emoji and vertical-spacing differences.

The [decimal-list reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-numbering-v15.4.json) adds two levels, parent resets, starts/restarts, group resumption and continuation around a table and image. Pages 15.4 opens the mixed and list-only bodies without repair. Inserting a parent survives save/reopen; native and composed PDFs retain every original and edited counter. Strict saved/reopened OfficeIMO DOCX retains counters and agrees with native DOCX indents. Font, spacing and receiving-application appearance remain separate limits.

The [third-level reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-third-level-v15.4.json) qualifies plain Arial 12 items at levels 0, 1 and 2, including resets after either ancestor advances, explicit restarts and return to an earlier group. A native grandchild insertion survives Pages 15.4 save/close/reopen. Saved OfficeIMO DOCX preserves every counter and renders with the same 74 word bounds as the native page in this reference. Complete composed PDFs retain every item with explicit portable-font substitution. Tiered numbering and wider typography remain separate qualification gates.

The [nine-level reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-nine-level-v15.4.json) qualifies all native decimal levels 0–8, ancestor resets, starts/restarts and resumed scopes in one Arial 12 body document. Nine fixed-length replacements preserve every byte outside their declared text ranges across 13 package entries and 120 IWA messages. Pages 15.4 retains root, middle and deepest-item edits and a new deepest-level item after save/close/reopen. Native exports and strict saved OfficeIMO DOCX retain all 37 original and 38 native-edited counters; its Pages rendering matches all 155 native word bounds in the preserved reference. All nine native/composed/receiving PDF pages are inspected. Composed PDFs retain complete content with explicit font substitution; tiered numbering, mixed fonts/sizes, broader typography, structural edits and cache regeneration remain outside this reference.

The [decimal marker-form reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-decimal-marker-forms-v15.4.json) qualifies `(1)` and `1)` markers at levels 0–8 with uniform Arial 12 text. Each form retains all 37 original and 38 native-edited counters, starts/restarts and resumed scopes. Nine fixed-length text edits per form change only their declared UTF-8 ranges; further native text edits and a deepest-level insertion survive Pages 15.4 save/close/reopen. Native exports and strict saved OfficeIMO DOCX retain the marker forms and indents. Each preserved DOCX renders with all 154 native word bounds unchanged in Pages. All sixteen native/composed/receiving PDF pages are inspected; composed typography retains explicit font substitution. This is an OfficeIMO-authored, Apple-saved body profile, with no broader producer, font, container or pagination claim.

The [mixed-font list reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-mixed-font-lists-v15.4.json) qualifies Arial, Times New Roman and Courier New runs at 12 points, with independent marker fonts and all three decimal marker forms. Pages 15.4 retains all 17 original items, root/deep text edits and an inserted eighteenth item after save/close/reopen. Strict saved OfficeIMO DOCX retains run families, emphasis and counters; its Pages rendering matches all 110 native word bounds. Nine native, composed and receiving PDF pages are inspected. Caller-font PDF output embeds the selected licensed faces; portable output declares standard-family substitution. Apple's DOCX export substitutes Arial Unicode MS and clears selected emphasis, so it is not a typography-equivalence reference. That reference uses a shared size; loaded rich-run editing remains unsupported.

The [mixed-size list reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-mixed-size-lists-v15.4.json) adds whole-point run sizes from 8 to 24 points at marker scale 1. Marker size follows the first retained run; the 12-point paragraph default determines the 18-point indentation step. Pages 15.4 retains root/deep text edits and a new 8-point child after save/close/reopen. Strict saved OfficeIMO DOCX retains all run sizes, families and counters. Its Pages rendering matches all 110 native word bounds; all twelve native, composed and receiving PDF pages are inspected. Fractional mixed sizes, other marker scales and loaded rich-run editing remain unqualified.

Native package identity includes decimal-marker punctuation. Period, paired-parenthesis and closing-parenthesis lists with the same body receive different identifiers; an unused list group leaves saved bytes unchanged. The [identity reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-marker-identity-v15.4.json) checks nine deterministic cases and four fresh Pages 15.4 open/edit/save/export/reopen cycles. The affected creation packages retain every native content-archive byte from their preceding qualified fixtures; only the identifier and matching property-list UUIDs change.

The [macOS arm64 runtime reference](../Docs/Compatibility/iwork-pages-marker-identity-native-aot-macos-arm64-2026-10-11.json) records native .NET 8 and .NET 10 execution of the mixed-font, whole-point mixed-size and marker-identity profiles, including strict saved/reopened DOCX.

The [default header/footer reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-created-headers-v15.4.json) qualifies repeated left text over two pages, including bold/color header styling, an italic Unicode footer and explicit distances. Pages 15.4 accepts edits to both areas after save/close/reopen. Independent decoding and strict saved/reopened OfficeIMO DOCX retain those styles and distances; complete composed PDFs retain every body paragraph and both repeated areas. Native DOCX export agrees on text and distances but changes font and bold/italic declarations. Additional section layouts and loaded-source header/footer editing remain outside this creation model.

## Create a native Numbers spreadsheet

```csharp
using OfficeIMO.IWork;

var spreadsheet = IWorkNumbersDocument.Create();
var sheet = spreadsheet.AddSheet("Budget");
var values = IWorkTableData.Create(2, 3)
    .SetText(1, 1, "Product").SetText(1, 2, "Amount").SetText(1, 3, "Active")
    .SetText(2, 1, "Widget").SetNumber(2, 2, 1200.5).SetBoolean(2, 3, true);
sheet.AddTable(values, "Sales", leftPoints: 40, topPoints: 80,
    columnWidthPoints: 130, rowHeightPoints: 36, fontSizePoints: 14);
spreadsheet.Save("budget.numbers");
```

The writer creates its own document, calculation metadata, sheet/sidebar, table, styles and hidden header/footer graph. It reuses the shared scalar table codec, native package encoder and Core atomic-file owner. Add at least one sheet and one table to each sheet. Sheet names are unique across the document; table names are unique within their sheet, compared without case. Names contain 1–128 UTF-16 code units and no control characters. Added tables snapshot their input data; their exposed collections are read-only.

Each table supports 1–256 positions on either axis, equal positive row/column dimensions and finite positions in points. Cells share a requested font, size and opaque sRGB color, white fills, black one-point borders, wrapping and four-point padding. Text, finite numbers, booleans and unset empty positions remain distinct; text beginning with `=` is literal. Formulas, comments, merges, headers, rich cell text, images and editing loaded sources are unsupported by this creation model. Numeric display formats use the shared API below.

`IWorkNumbersWriteOptions` bounds sheets, combined tables, dense grid positions including unset cells, combined UTF-16 sheet/table names and cell text, and package/intermediate bytes. Encoding and validation finish before copying or atomic path staging. Stream saves write at the current position and leave the caller open; copying failures can leave partial stream data. Path saves reject existing destinations unless `overwrite: true` is supplied and preserve the existing file after failed staging. Save operations accept cancellation. Creation instances are not thread safe.

The [Numbers creation reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-created-v14.5.json) records independent decoding of all 108 objects and native Numbers 14.5 open, numeric edit, save, close/reopen and PDF/XLSX export. Both sheets and both PDF pages are inspected. Saved/reopened OfficeIMO XLSX under explicit partial acceptance retains all three tables and their typed values, including Polish text, emoji, booleans, literal equals-sign text and a completely empty grid. Four-point native region padding remains available through `IWorkTable.GetPadding`; XLSX reports `IWORK_NUMBERS_CELL_PADDING_OMITTED` and requires `AllowPartialEditableReconstruction = true`. Complete-content PDF workflows reject this omission. Apple's empty-table Excel placement differs from the table-relative OfficeIMO output; this reference does not qualify equivalent Excel positioning or appearance. Other versions, fonts, scripts, layouts and runtime architectures require additional qualification.

## Format numeric cells in new tables

All three native creators accept the same immutable format metadata:

```csharp
var values = IWorkTableData.Create(3, 1)
    .SetNumber(1, 1, -1234.5, IWorkNumberFormat.CreateNumber(2, true,
        IWorkNegativeNumberStyle.Parentheses))
    .SetNumber(2, 1, 0.125, IWorkNumberFormat.CreatePercentage(2))
    .SetNumber(3, 1, 12345, IWorkNumberFormat.CreateScientific(4));
// Pass values to a Pages, Numbers or Keynote AddTable call.
```

Decimal and percentage formats support 0–30 fractional digits, optional grouping and minus, red, parentheses or red-and-parentheses negatives. Scientific formats support 0–30 fractional mantissa digits with minus negatives and no grouping. Formatting changes display, not the stored finite double. A later three-argument `SetNumber` clears the explicit format. Tables retain their format snapshots when the input grid changes.

The [native format reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/iwork-created-numeric-formats.json) pairs independent graph/cell decoding with Numbers 14.5, Pages 15.4 and Keynote 15.4 open, format-inspector checks, numeric edits, save/reopen and native editable/PDF exports. All 57 edited numeric cells retain their values and selected format metadata; every native PDF page is inspected. Saved OfficeIMO XLSX retains numeric values and format codes. Saved DOCX/PPTX uses the existing invariant display contract and reports approximation for native locale and appearance. Numbers' inherited four-point cell padding still requires partial XLSX acceptance.

Automatic, currency, fraction, temporal, custom and locale-specific formats are not authored. The receiving app selects separators, fonts, fitting and pagination; extra displayed digits do not increase scalar precision. Source reading supports additional qualified format families independently of this creation subset.

The [macOS arm64 NativeAOT receipt](../Docs/Compatibility/iwork-numeric-formats-native-aot-macos-arm64-2026-10-10.json) records 31 native smoke profiles on .NET 8 and .NET 10, including deterministic numeric-format creation, native-edited readback and saved DOCX/XLSX/PPTX.

## Save a captured native source

```csharp
using OfficeIMO.IWork;

var source = IWorkSourceDocument.Open("source.pages");
var session = source.BeginEdit(); // separate from the immutable source snapshot
IWorkNativeSaveReport report = session.Save("preserved.pages");
// An unchanged save has equal normalized source and output content hashes.
```

An unchanged save writes all captured file-entry payloads into a native ZIP package. Unknown IWA records, auxiliary payloads, IDs, references, version fields, resources and existing previews remain intact, even when conversion-report source-record preservation is disabled. ZIP compression, timestamps, extra fields, comments, directory entries and filesystem metadata are rewritten or omitted. This is file-entry preservation rather than whole-file byte identity. Directory bundles and nested `Index.zip` sources can be saved unchanged; nested index contents stay inside their original entry.

Keynote image mutations move and resize direct, unlocked, unrotated and unmasked images within their slide canvas:

```csharp
var keynote = IWorkSourceDocument.Open("deck.key");
var image = keynote.ReadKeynote().Slides[0].Images[0];
var edit = keynote.BeginEdit();
edit.MoveKeynoteImage(image.SourceIdentity!, leftPoints: 100, topPoints: 130);
edit.ResizeKeynoteImage(image.SourceIdentity!, widthPoints: 420, heightPoints: 168);
var report = edit.Save("edited.key", options: new IWorkNativeSaveOptions {
    AllowUnassessedVisualCaches = true
});
```

Only the requested existing position or size floats change. Image pixels, intrinsic dimensions, descriptions, IDs, references and all other raw fields remain intact. Sizes must be positive and finite; an explicit resize can change the aspect ratio. Moves and resizes compose using the pending frame, which must fit within the canvas after each operation. Missing or ambiguous geometry, unsupported image states, out-of-canvas frames and nested-index mutations are rejected. Repeating an operation uses absolute values; restoring the complete original frame removes that pending edit.

Numbers sessions replace existing numeric or boolean cells using one-based coordinates and the table's source identity:

```csharp
var numbers = IWorkSourceDocument.Open("budget.numbers");
var table = numbers.ReadNumbers().Sheets[0].Tables[0];
var cells = numbers.BeginEdit();
cells.SetNumbersCellNumber(table.SourceIdentity!, row: 2, column: 2, value: 42.125);
cells.SetNumbersCellBoolean(table.SourceIdentity!, row: 2, column: 3, value: false);
cells.Save("edited.numbers", options: new IWorkNativeSaveOptions {
    AllowUnassessedVisualCaches = true
});
```

This profile requires complete modern Numbers content with no selected formulas or unsupported cell features. Numeric cells use Decimal128 storage and accept finite doubles; booleans retain their distinct native type. Active version-five BNC storage and an absent or empty legacy placeholder are required. Only each existing fixed-width scalar changes, preserving styles, formats, unknown bytes, IDs and references. Edits compose across cells and tables. Assigning a cell's original projected value restores its original scalar bytes, including higher-precision source representations. Invalid or cancelled requests leave pending edits unchanged. Insertion, type changes, shared tiles, populated legacy values, double-backed numeric cells and nested-index mutations are rejected.

The [Numbers scalar-edit reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-preserved-scalars-v14.5.json) qualifies six numeric/boolean replacements across two tables. Numbers 14.5 opens both sheets without repair, retains a further numeric edit after save/close/reopen and exports typed XLSX plus a two-page PDF. Independent raw comparisons admit only the requested scalar ranges; all other file-entry payloads, headers and sibling records remain identical. OfficeIMO XLSX retains the typed content with the existing explicit padding loss. Complete composed PDF remains rejected for this padding profile.

The [macOS arm64 NativeAOT receipt](../Docs/Compatibility/iwork-numbers-cell-editing-native-aot-macos-arm64-2026-10-11.json) records 34 native smoke profiles on each of .NET 8 and .NET 10, including scalar editing, native-edited readback, saved XLSX, cancellation and original-byte restoration.

Pages sessions replace a plain body paragraph using its one-based paragraph number and the body text storage's identity:

```csharp
var pages = IWorkSourceDocument.Open("source.pages");
var body = pages.ReadPages().Body;
var paragraphs = pages.BeginEdit();
paragraphs.SetPagesBodyParagraphText(body.SourceIdentity!, paragraph: 1, text: "Edited paragraph layout");
paragraphs.Save("edited.pages", options: new IWorkNativeSaveOptions {
    AllowUnassessedVisualCaches = true
});
```

Replacement text must keep the original paragraph's UTF-16 length and UTF-8 byte length; for example, `Native paragraph layout` can become `Edited paragraph layout`. The source must have complete modern body content in at most one section, without drawables, inline objects or nonempty headers/footers. The selected paragraph must have a uniform run without hyperlinks. Plain paragraphs and qualified non-tiered decimal items at levels 0 through 8 are supported; decimal edits retain their counters and indents. TAB characters are allowed; line breaks and inline-object markers are rejected. Only the requested UTF-8 range changes, retaining delimiters, native style offsets, unknown bytes, IDs and references. Repeated edits compose, invalid or cancelled requests preserve pending changes, and restoring the original text restores its original bytes. Length changes, rich-run replacement, additional list profiles and nested-index mutations require another edit profile.

The [Pages paragraph-edit reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-preserved-paragraphs-v15.4.json) qualifies three replacements in an Apple-saved source. Independent comparison accounts for all 13 file entries and 93 IWA messages. Pages 15.4 opens the result without repair and retains a further native edit after save/close/reopen. Native and strict OfficeIMO DOCX agree on body text, paragraph spacing and custom tabs. Complete composed PDF retains the selected content with explicit font substitution; native typography and cache regeneration remain unqualified.

The [decimal-item edit reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-preserved-decimal-v15.4.json) qualifies four replacements across three levels in an Apple-saved source. Independent comparison accounts for all 13 entries and 98 IWA messages; only the requested text ranges change. Pages 15.4 retains a further list-item edit after save/close/reopen. Native exports and saved OfficeIMO DOCX preserve every counter and indent, and its Pages rendering matches all 74 word bounds in this Arial 12 reference. Complete composed PDF retains the edited items with explicit portable-font substitution. Wider typography, pagination and native cache regeneration remain unqualified.

The [macOS arm64 NativeAOT receipt](../Docs/Compatibility/iwork-pages-paragraph-editing-native-aot-macos-arm64-2026-10-11.json) records 35 passing native smoke profiles on each of .NET 8 and .NET 10, including paragraph replacement, native-edited readback, saved DOCX, cancellation, length rejection and original-byte restoration.

Keynote sessions use the text box's source identity and a one-based paragraph number:

```csharp
var deck = IWorkSourceDocument.Open("source.key");
var box = deck.ReadKeynote().Slides[0].TextBoxes[0];
var textEdit = deck.BeginEdit();
textEdit.SetKeynoteTextBoxParagraphText(box.SourceIdentity!, paragraph: 1,
    text: "Edited PNG — image behind text");
textEdit.Save("edited.key", options: new IWorkNativeSaveOptions {
    AllowUnassessedVisualCaches = true
});
```

The paragraph must retain both its UTF-16 length and UTF-8 byte length; `Native PNG — image behind text` can become `Edited PNG — image behind text`. The source must have complete modern Keynote content. The selected shape must be a direct unlocked freeform text box with unique storage, without a shape hyperlink or inline objects. Its selected paragraph must have a uniform run, no list marker and no hyperlink. Use `box.SourceIdentity`, which identifies the shape; `box.Content.SourceIdentity` identifies its storage and is not the setter argument. Edits retain styles, offsets, geometry, images, unknown bytes, IDs and references. Replacements compose across paragraphs and shapes, rejected/cancelled requests preserve pending changes, and restoring original text restores its original bytes. Placeholder/inherited/shared storage, length changes, rich-run/list replacement, structural edits and nested-index mutations require another profile.

The [Keynote paragraph-edit reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-preserved-paragraphs-v15.4.json) qualifies three replacements in an Apple-saved deck. Independent comparison accounts for all 24 entries and 205 IWA messages, admitting only the three requested UTF-8 ranges. Keynote 15.4 opens the result without repair and retains a further native edit after save/close/reopen. Strict OfficeIMO PPTX and native PPTX agree on ordered text, geometry, image bytes and descriptions across all three slides. Complete composed PDF retains the selected scenes with explicit font substitution; native typography and cache regeneration remain unqualified.

The [macOS arm64 NativeAOT receipt](../Docs/Compatibility/iwork-keynote-paragraph-editing-native-aot-macos-arm64-2026-10-11.json) records 36 passing native smoke profiles on each of .NET 8 and .NET 10, including Keynote paragraph replacement, native-edited readback, strict saved PPTX, cancellation, length rejection and original-byte restoration. PDF conversion has separate managed evidence.

Edited saves require `AllowUnassessedVisualCaches` and report `IWORK_NATIVE_VISUAL_CACHES_UNASSESSED` with `Unassessed` fidelity. The operation retains existing native previews and visual caches without regenerating them or assessing wider application metadata. An unchanged save needs no cache acceptance. The [native preservation reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/native-preservation-macos.json) records independent raw-entry/envelope comparisons, unchanged native open in all three applications, Keynote 15.4 translation and subsequent native edit/save/reopen, and all-slide PDF inspection. It qualifies those fixtures; other application versions and visual-cache behavior require separate evidence.

The [image resizing reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-preserved-resized-v15.4.json) qualifies composed moves and resizes of opaque PNG, alpha PNG and JPEG images, including a subsequent native width edit and save/close/reopen in Keynote 15.4. Independent IWA, PPTX and PDF checks retain exact image resources and final frames; native and composed PDF inspection confirms transparency and stacking. Font metrics differ, and wider image states and native cache regeneration remain outside this profile.

`IWorkNativeSaveOptions` independently bounds the sum of input file-entry bytes and encoded ZIP length, combined decompressed bytes of changed archives, and distinct edited records. `SaveBytesResult` returns a defensive package copy through `GetPackageBytes()` and an operation-level report with normalized content hashes, physical entry/record counts and updated identities/entry paths. Counts do not establish semantic fidelity of opaque content. Encoding completes before writing. Stream saves start at the caller's current position, leave the stream open and can leave partial bytes on copying failure or cancellation. Path saves are atomic, require `overwrite: true` for existing files and reject output inside a captured directory source, including directory aliases. Sessions are not thread safe.

## Read and inspect a source

```csharp
using OfficeIMO.IWork;

IWorkSourceDocument source = IWorkSourceDocument.Open("report.pages");
IWorkPagesProjection pages = source.ReadPages();

Console.WriteLine(source.Kind);                 // Pages
Console.WriteLine(source.ContainerKind);        // ZipPackage, DirectoryBundle, or nested Index.zip
Console.WriteLine(string.Join(", ", source.BuildVersions));
Console.WriteLine(pages.Paragraphs.Count);

IWorkConversionReport report = pages.CreateConversionReport(
    IWorkProjectionKind.EditableReconstruction);
foreach (IWorkArchiveRecord record in report.PreservedRecords) {
    Console.WriteLine($"{record.EntryPath}: {record.MessageType}");
}
foreach (IWorkSourceUnitCount count in report.SourceUnitCounts) {
    Console.WriteLine($"{count.Kind}: {count.ReconstructedCount} reconstructed, {count.OmittedCount} omitted, {count.UnassessedCount} unassessed");
}
foreach (IWorkSourceReferenceIssue issue in report.SourceReferenceIssues) {
    Console.WriteLine($"{issue.Owner.RecordIdentifier}/{issue.FieldPath}[{issue.ReferenceIndex}]: {issue.Kind}, target {issue.TargetIdentifier}");
}
foreach (IWorkSourceDeclarationIssue issue in report.SourceDeclarationIssues) {
    Console.WriteLine($"{issue.Owner.RecordIdentifier}/{issue.FieldPath}: {issue.Kind}, outer values {issue.DeclaredValueCount}");
}
```

Pages body runs expose qualified drawable attachments through `IWorkTextRun.InlineObject`. The attachment keeps its original zero-based UTF-16 `CharacterOffset`, native attachment identity, and referenced drawable identity. Marker-only runs have empty `Text`; use `InlineObject` to distinguish them from ordinary text. Supported zero-offset image and standalone-table placement is described in the [support matrix](../Docs/officeimo.iwork-support-matrix.md#pages-inline-attachments).

Selected character styles expose font-relative `IWorkTextStyle.CharacterTrackingEm` and computed `CharacterSpacingPoints`. The spacing uses the effective inherited font size; an explicit zero clears inherited tracking. Word and PowerPoint map known spacing to their native run properties and report quantization. XLSX has no equivalent tracking property and requires partial acceptance for nonzero tracking. Missing effective font size or destination range violations remain unsupported. The [Keynote tracking reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-image-tracking-v15.4.json) qualifies a 116-point title with -0.02-em tracking, or -2.32 points, against Apple PPTX properties and a bounded native rendering comparison.

Selected modern cell styles expose `IWorkTableCell.Fill`, including supported solid colors and explicit no-fill overrides. Selected styles also expose `Padding` in points, `VerticalAlignment` and `Borders`. A declared empty padding message resets all four sides to zero. Empty cells carrying supported formatting remain in `Cells` with `Kind == Empty` and count toward the materialized-cell budget. Filter by `Kind` when counting value cells. Missing or unsupported selections retain source diagnostics; conditional fills and complete cell styling remain outside this qualified subset.

`table.GetBorders(row, column)` resolves inherited defaults, region visibility and ordered sparse edits into four effective edges for a one-based cell. It uses the shared `OfficeBorderBox`: zero-width sides explicitly hide borders, while null sides are absent or unresolved. Empty source cells remain sparse. `MaximumTableBorderEntries` bounds native layers and runs. The [border contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-borders) defines qualified strokes and regions, destination width limits and independent Keynote/Numbers evidence. The native Numbers border fixture requires explicit partial acceptance because its cell padding has no XLSX equivalent; its border edges remain intact.

Shared tables expose explicit `RowHeights` and `ColumnWidths` as read-only maps keyed by one-based positions, measured in points. `GetRowHeight(row)` and `GetColumnWidth(column)` return an explicit size or the table default. Native zero-size entries use the default. `MaximumTableDimensionEntries` bounds dimension headers and bucket references across the source. Unsupported sizing records remain in source evidence and emit a diagnostic.

Positive native hidden/filtered row or column counts, and malformed count declarations, emit `IWORK_TABLE_VISIBILITY_UNASSESSED` and require explicit partial conversion for editable output. Partial output and Reader can include content hidden in the source. Counts do not identify which positions are hidden and can be zero even when native hidden-state records select content.

Qualified base user-hidden states resolve UUIDs through a complete, unambiguous table position map. `IWorkTable.HiddenRows` and `HiddenColumns` expose sorted one-based positions; cells remain in the model. Excel preserves these positions as hidden rows and columns. Reader includes their content with `IWORK_READER_HIDDEN_TABLE_CONTENT_INCLUDED`. Word and PowerPoint require the partial policy because their table output keeps those positions visible.

Unresolved base selections, summary states, filtering, pivot hiding, collapsed groups, other unqualified extents and unreadable visibility envelopes emit `IWORK_TABLE_HIDDEN_STATES_UNASSESSED`. Empty extents, false selectors and disabled filters retain the editable path; disabled rules are not traversed. `MaximumTableDimensionEntries` also bounds hidden-state declarations, selected UUID maps and filter references. The [independent corpus](../OfficeIMO.TestAssets/Documents/IWorkCorpus/README.md#base-user-hidden-columns) qualifies native hidden columns with empty cells; positive rows and populated hidden cells have synthetic format-boundary proof. Apple export/render equivalence remains unqualified.

Path and stream entry points use the same bounded parser. Stream and byte-array overloads detect the application kind from bounded package content. Pass an expected `IWorkDocumentKind` when the caller already knows the route and wants a mismatch rejected:

```csharp
using FileStream stream = File.OpenRead("budget.numbers");
IWorkSourceDocument source = IWorkSourceDocument.Open(
    stream,
    new IWorkReadOptions {
        MaximumPackageBytes = 64 * 1024 * 1024,
        MaximumArchiveReferenceCount = 1_000_000,
        MaximumMaterializedCells = 1_000_000,
        MaximumTableCatalogEntries = 100_000,
        MaximumImageMetadataEntries = 16_384,
        MaximumFormulaRenderingOperations = 64L * 1024 * 1024
    });

IWorkNumbersProjection workbook = source.ReadNumbers();
```

The verifying form is `IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers, options)`.

`IWorkTable.FillStyles` exposes supported body, header-row, header-column and footer-row fills, plus `BandedBody` for every second body row after the headers. `table.GetFill(row, column)` resolves a one-based position without materializing unstored cells. Selected fills override region defaults, including explicit no-fill; unresolved selected fills stay null. Header and footer rows take precedence over header columns. Banding applies only to body cells. Unsupported or unresolved banding suppresses fill defaults and produces `IWORK_TABLE_FILL_DEFAULTS_UNSUPPORTED`; selected cell fills remain available. The [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-fill-defaults-and-banding) defines the native Numbers export evidence and remaining appearance limits.

Table text formatting is separate from typed values. `IWorkTable.TextStyles` exposes applicable region defaults, `IWorkTableCell.ParagraphStyle` exposes a selected text style, and `table.GetParagraphStyle(row, column)` resolves the supported style for a one-based position, including an unstored empty cell. Selected styles override region defaults; an unresolved selected style stays null and produces `IWORK_TABLE_TEXT_STYLE_UNSUPPORTED`. The [support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md#table-text-defaults-and-selected-styles) defines the qualified fields and limits.

`IWorkKeynoteSlide.HasBackgroundFill` indicates a recovered slide-style background. `BackgroundColor` contains its opaque RGB color; a null color with `HasBackgroundFill = true` denotes explicit no-fill. Selected style inheritance and unsupported-fill diagnostics are described in the [background contract](../Docs/officeimo.iwork-support-matrix.md#keynote-slide-backgrounds).

Keynote `IWorkTextBox.Layout` exposes selected inherited margins in points, `VerticalAlignment` and `ShrinkToFit`; a null layout means those frame properties were not recovered. `IWorkTextParagraph.ListLayout` exposes qualified character-marker indentation in points, text indentation relative to font size and marker scale. These immutable source properties let callers inspect layout without constructing a destination. The [support matrix](../Docs/officeimo.iwork-support-matrix.md#qualified-everyday-subset) defines the bounded PPTX mapping and other destination limits.

## Cancellation

Path, stream, and byte-array `Open` overloads accept a `CancellationToken` after the read options. The token governs loading and all later semantic projections and conversions from that source. Use `source.WithCancellation(newToken)` for an independently cancellable projection or conversion without rereading the package. The view shares loaded bytes and parsed messages and replaces the previous token; `CancellationToken.None` permits reuse after the opening token is cancelled. Caller-owned streams remain open when loading succeeds or is cancelled.

```csharp
using var cancellation = new CancellationTokenSource();
IWorkSourceDocument source = IWorkSourceDocument.Open(
    "budget.numbers", IWorkDocumentKind.Numbers, options: null,
    cancellationToken: cancellation.Token);
IWorkNumbersProjection workbook = source.ReadNumbers();
```

The destination adapters also accept a token after their read and conversion options. Cancellation is cooperative during loading, projection, and destination construction; saving uses the destination owner's separate save API. These APIs do not establish a fixed cancellation latency or memory budget.

The [bounded NativeAOT contract](../Docs/officeimo.iwork-support-matrix.md#bounded-nativeaot-qualification) covers source projections, shared Reader extraction, editable DOCX/XLSX/PPTX conversion and native Keynote creation on macOS arm64, Linux x64 and Windows x64 under .NET 8 and 10. Hash-pinned receipts record their exact source revisions, save/reopen and selected Numbers recalculation; current Numbers conversions require partial acceptance when inherited or selected padding cannot be represented in XLSX; the native Keynote creator emits the same independently decoded, Apple-accepted package on all six host/runtime combinations. The Pages body and Numbers scalar-table creation extensions, unchanged native preservation and bounded Keynote image moves/resizes have actual .NET 8/10 native execution on macOS arm64. The [image resize receipt](../Docs/Compatibility/iwork-keynote-image-resize-native-aot-macos-arm64-2026-10-11.json) records 33 passing native smoke profiles per runtime. The [fixed second-duration profile](../Docs/Compatibility/iwork-duration-seconds-native-aot-macos-arm64-2026-10-10.json) also qualifies saved/reopened native Numbers caches, formula recalculation and integral-second display boundaries on that host. The [fixed millisecond-duration profile](../Docs/Compatibility/iwork-duration-milliseconds-native-aot-macos-arm64-2026-10-10.json) adds saved/reopened native Numbers caches and formula recalculation for three fractional-second ranges on that host. Rendering, complete fidelity and Apple sandbox/device acceptance remain separate qualification requirements.

## Opt in to an Office destination adapter

Install only the adapter for the destination format you need. The Word, Excel, and PowerPoint packages do not depend on iWork.

- [Pages to Word](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Word.IWork/README.md)
- [Numbers to Excel](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Excel.IWork/README.md)
- [Keynote to PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.PowerPoint.IWork/README.md)

Each adapter README owns its conversion API and example. The source reader remains useful on its own for inspection, extraction, and application-owned projection workflows.

This is extended semantic reconstruction rather than plain-text extraction:

- Pages recovers rich paragraphs, source-proven list levels and explicit marker fonts mapped to native Word numbering, page layout, section-specific headers/footers, positioned and sized accessible rich-text boxes, images, rich-text table cells, and merges for editable Word projection.
- Numbers recovers sparse typed cells, rich-text cell runs, supported formulas with cached values, merges, table metadata, and default sizing for editable Excel projection. Each source table receives its own worksheet so table-local formulas and column sizing remain stable; sheet-level text receives a separate worksheet when present. Coordinate-backed single-cell, same-target endpoint ranges and finite rectangular cross-table references resolve unique native identities among selected tables. Their source `Formula` uses quoted sheet/table labels, such as `=SUM('Sheet'::'Table'::$A$1:$B$2)`; the Excel adapter renders the retained expression with actual worksheet names. Table-local and cross-table whole-axis references, including header-named rows and columns, retain coordinate source expressions. The Excel adapter writes the current body extent and reports `IWORK_NUMBERS_TABLE_BODY_RANGE_APPROXIMATED` because named labels and automatic expansion are not preserved. Missing or ambiguous identities and ambiguous body metadata remain incomplete with valid caches retained.
- `IWorkNumbersSheet.Drawables` retains tables and text shapes in their shared source order. The separate `Tables` and `TextBoxes` lists remain available for type-specific access.
- Keynote recovers slide size, order and names, positioned rich text with explicit inline breaks and source-proven list labels, levels and explicit marker fonts, shape/run and presenter-note hyperlinks, notes, images, rich-text table cells, positioned and rotated tables, and merges for editable PowerPoint projection.

Advanced charts, vector effects, animations, non-cell comments/change tracking, masks/crops, and other application-only structures remain available in source records. Reachable unsupported structures produce conversion diagnostics; preserved records alone do not establish a field-level fidelity assessment. Keynote measurements finer than PPTX's integral EMU grid stay editable and emit `IWORK_KEYNOTE_PPTX_PRECISION` when they are quantized to the nearest destination unit.

`IWorkReadOptions` bounds decoded text characters, text items and attribute boundaries, cross-record style inheritance, projected sheets/slides/tables/images, repeated encoded destination-image bytes, merged ranges, source-wide table catalogs, materialized cells, and ArchiveInfo references in addition to the package/IWA byte limits.

All conversion modes use the same bounded semantic source read, so package and projection limits are enforced before the destination representation is chosen. `Auto` prefers editable semantic reconstruction. `EditableOnly` fails when supported editable structure cannot be recovered. `VisualOnly` selects the package's raster preview for the destination and reports `VisualFallback`; it does not erase or bypass the semantic `ReadPages`, `ReadNumbers`, or `ReadKeynote` projection. `RequireCompleteVisualCoverage` defaults to `true`, so incomplete or unknown preview coverage is rejected. Set it to `false` explicitly on a result API to retain and inspect preview output. Value-only adapter APIs require complete editable reconstruction and reject partial or preview results even under permissive options. `AllowPartialEditableReconstruction = true` retains bounded recoverable objects and exposes `IsPartialEditableReconstruction`; source and destination limits still apply.

A raster preview is classified as first-page or composite coverage. An embedded PDF has unknown coverage: its path and structural validity do not establish that it matches the current complete source. Coverage is exposed on `IWorkPreviewAsset`. Embedded PDF inspection accepts bounded classic cross-reference tables and rejects unvalidated cross-reference streams.

`IWorkTextParagraph.ListMarkerKind` distinguishes native image, literal-text and numbered markers. A text marker such as `1.` stays a literal DOCX bullet and an unordered Reader list; it does not become a counter. PPTX retains single-character text bullets and reports unsupported multi-character markers through destination fallback. Image markers remain unassessed, retain declaration evidence at the selected list-style field `11`, and require explicit partial reconstruction for editable output. `ListLabel` describes the initial numbering format. Qualified non-tiered decimal paragraphs at levels 0 through 8 also expose `ListNumber`, the resolved counter, and `ListStartNumber`, an explicit start or restart. Positive paragraph-start declarations establish a counter; zero continues the selected list through ordinary non-list text. An advancing or restarted ancestor resets all descendant counters. The observed restart flag requires a matching positive start; rejected list declarations end the counter scope. The [Pages continuation reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/pages-list-continuation-v15.4.json) records root counters 7–11, child resets, a child restart at 5 and a root restart at 3. Additional scripts, tiered counters and continuation through other numbered or bullet styles remain unqualified.

`IWorkTextParagraph.ListFontName` retains an explicit inherited list-marker font independently of the paragraph and run fonts. A cleared or unspecified font is `null`. DOCX and PPTX adapters preserve explicit marker fonts; Reader reports their omission from Markdown. Rejected child marker-font declarations clear the inherited value and retain source-declaration evidence at list-style field `22` or `23`. Font availability and native rendered glyph equivalence require separate platform qualification.

## Preservation and authoring boundary

Projected documents, sheets, slides, tables, images, and text expose `SourceIdentity` with the native IWA identifier, message type, entry path, and payload position. Adapter reports retain these identities in `SourceUnits` and summarize them by kind in `SourceUnitCounts`, even when source payload preservation is disabled. The inventory counts identified units selected by the semantic projection: text storages are counted once, auxiliary table models and tiles are excluded, and inactive template records are excluded. Explicitly dropped selected units are reported as omitted. Selected unsupported drawables and Numbers sheet references with unsupported native types retain their identities under `UnsupportedObject` when no recognized kind applies. A preview leaves individual unit coverage unassessed. Unresolved references and content paths that the projection has not assessed remain outside this inventory. A reconstructed unit can still contain omitted, approximated, or unassessed fields.

`SourceReferenceIssues` separately records failed declared reference occurrences in assessed content paths: Pages body, stacking-order, decoded floating-canvas, shape text-storage, section, declared header/footer template and header/footer storage references; Numbers sheets, sheet drawables and shape text-storage references; and Keynote show, slide-tree nodes, slides, slide drawables, selected drawable text-storage and presenter-note references. Each issue retains the owning record, protobuf field path, one-based occurrence and readable target identifier. Missing targets, malformed references and readable siblings in a rejected set are distinct. Existing targets whose native type is unsupported by the selected path use `UnexpectedTargetType`. This covers decoded paragraph, list, character and hyperlink attribute references, traversed style parents, and applicable table style, role and selected cell-style references. Character attributes accept both character and paragraph styles. Typed structural links also cover Pages body/shape storage, section and header/footer dependencies, Numbers sheets and shape storage, the Keynote show/slide-node/slide chain, direct and nested drawable text storage, presenter-note wrappers and note storage, selected inline attachment/drawable, table-model, validated tile, sizing-bucket, rich-text catalog/wrapper/storage, selected cell-style and number-format catalog, and qualified root-comment catalog/record/author references. Rejected repeated sets keep their readable siblings classified as unselected; wrong-type targets in an otherwise accepted set retain their physical positions without being counted as missing objects. Unselected catalog values and inapplicable roles remain outside this assessment. Repeated occurrences within a source field remain separate; selecting a shared archive several times does not multiply its evidence. These occurrences do not identify distinct omitted objects, and preview output does not resolve them. Selected tables also assess table-info/model, tile, row/column sizing-bucket, string/formula/rich-text catalog, selected rich-text wrapper and wrapper/storage references. Tile paths retain physical entry positions; entries rejected before their target is read remain unassessed. Declared unresolved or malformed string/formula catalog links prevent complete editable reconstruction while valid cell caches remain recoverable through the explicit partial-conversion policy. Unreferenced templates, unused storage fallbacks and unparseable graph containers remain outside this assessment. `MaximumSourceReferenceIssues` defaults to 100,000 and bounds the cumulative occurrences inspected in distinct fields with failures, including readable siblings, before evidence is materialized. Adapter reports classify this evidence as `Unassessed` through `IWORK_SOURCE_REFERENCES_UNRESOLVED`.

`SourceDeclarationIssues` records unreadable, rejected or unsupported declarations separately from failed references and identified objects. Each item retains its owner, protobuf path, failure kind and known outer field-value count. A whole payload uses `$` and has no outer count; nested repeated entries use one-based positions. These counts do not reveal nested reference totals or omitted content. Coverage includes selected document roots, Pages body/shape/section/header/footer records and stacking/floating graph envelopes, Numbers sheet/plain-text record failures, Keynote show/tree/node/slide/drawable/note record failures, table-info/model/store/tile-list and row-header containers, selected string/formula/rich-text catalog envelopes and entry keys, plain-string values and formula-message envelopes, row/column sizing-bucket payloads and header metadata, selected merge-owner/store/pair/formula declarations, selected rich-text wrapper/storage records, decoded rich-text attribute containers and invalid entry positions, unreadable styles or style-parent envelopes actually traversed for text, and invalid values in supported character and paragraph properties, list label/type/indent vectors, style names and selected hyperlink text. Inactive records and unused rich-text value traversal are excluded. Other unsupported fields and remaining unmaterialized cell storage beyond the selected modern-row evidence remain coverage gaps. Shared paths report once per projection. `MaximumSourceDeclarationIssues` defaults to 100,000 and bounds distinct retained paths; evidence survives disabled source-payload preservation and visual fallback, with category `Unassessed` through `IWORK_SOURCE_DECLARATIONS_UNASSESSED`. Configured protobuf field, repeated-value and depth limits remain fatal; text-budget failures propagate rather than being reported as malformed declarations.

Modern row selection retains `InvalidValue` evidence at tile paths `5[n]/2` for a malformed, ambiguous or inconsistent declared cell count, and `5[n]/6` for a non-empty modern buffer with no selected offsets. A readable count is compared with selected physical cell records, including unformatted empty records that need no destination cell. An absent count remains supported. These paths describe outer field values, not missing-cell totals; no cell coordinates or formula presence are inferred from unselected bytes. Recoverable cells remain available through explicit partial conversion, while `IWORK_TABLE_ROW_STORAGE_UNASSESSED` reports `Unassessed` fidelity. Empty buffers with no selected offsets remain valid. Legacy companion buffers are not treated as unselected modern content when modern storage is present. Shared paths use the existing declaration budget and report once per projection. Selected cells end at the next declared physical record, including bounded offsets outside the table; truncated values and formula caches cannot consume that record's bytes. Populated out-of-table offsets and oversized offset envelopes retain `InvalidSelectionMetadata` at `5[n]/7`. Envelopes exceeding the larger of the declared column count and 256 slots are rejected before slot inspection; readable sibling rows remain recoverable without inferred cells or formulas from the rejected row. Supported tables wider than 256 columns remain readable. Every inspected offset slot, including empty trailing slots, consumes `MaximumTableDimensionEntries` across the source. Invalid offsets beyond the buffer do not discard a preceding complete cell.

`IWorkTableCell.UnsupportedFeatures` identifies selected modern conditional-style, applied-rule, unresolved or unsupported comment and scalar-format fields after the complete cell storage has been decoded. These flagged fields remain unassessed. Qualified root comments are exposed separately through `IWorkTableCell.Comment`. `None` does not establish complete cell fidelity. Cells with these features, including otherwise empty cells, remain materialized under `MaximumMaterializedCells`. `NumericFormat` retains each cell whose selected numeric or currency format cannot be resolved or qualified, including blank cells; catalog-level diagnostics remain available alongside this cell-level evidence. Their row buffer retains `UnsupportedField` declaration evidence at tile path `5[n]/6`, deduplicated under the existing declaration budget; its outer count is not a cell, rule or comment count. `IWORK_TABLE_CELL_FEATURES_UNASSESSED` reports the number of affected selected cells as `Unassessed` and requires the existing partial-conversion policy to retain editable values. Reader retains the warning. Truncated cells keep decode-failure evidence instead of assessed feature presence. Unused feature catalogs, conditional-rule contents, controls, suggestions, formula-error details and import-warning selectors stay outside this inventory.

Selected date/time format `261` exposes `IWorkNumberFormat.DateTimeFormat.SourcePattern` for the five complete patterns `dd/MM/y`, `dd/MM/y HH:mm`, `HH:mm:ss`, `h:mm a` and `d MMM yyyy`. They map to invariant XLSX formats and the shared calendar formatter for Word, PowerPoint and Reader tables. Four-digit calendar years, abbreviated English month names and lowercase day periods follow the independent parser display evidence. Typed dates and original formula caches remain unchanged; XLSX pre-1900 and precision guards still apply. `IWORK_NUMBERS_DATE_DISPLAY_APPROXIMATED` reports locale, calendar, timezone and punctuation limits. Other patterns, true suppression settings, ambiguous declarations or extra metadata retain `DateFormat` evidence and the partial policy. The [date-format evidence](../OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/date-formats.json) checks 89 selected cells in five unchanged sources. The paired Numbers 14.5 export uses localized punctuation and refreshed volatile caches, so it does not establish source/native cache or appearance equivalence.

Fixed abbreviated duration formats expose `IWorkNumberFormat.DurationFormat`. The accepted catalog has exactly one value for format type `268`, style `1`, largest/smallest units and automatic units `false`; missing, duplicate, custom or automatic settings retain unsupported-format evidence.

| Largest / smallest unit | XLSX format | Value contract |
| --- | --- | --- |
| Day / day (`2` / `2`) | `0"d"` | Numeric total day serial |
| Hour / minute (`4` / `8`) | `[h]"h" m"m"` | Total hours and minute component |
| Hour / second (`4` / `16`) | `[h]"h" m"m" s"s"` | Total hours, minute and second components |
| Minute / second (`8` / `16`) | `[m]"m" s"s"` | Total minutes and second component |
| Second / second (`16` / `16`) | `[s]"s"` | Total seconds |
| Hour / millisecond (`4` / `32`) | `[h]"h" m"m" s.000"s" ` | Total hours, minute component and three fractional-second digits |
| Minute / millisecond (`8` / `32`) | `[m]"m" s.000"s" ` | Total minutes and three fractional-second digits |
| Second / millisecond (`16` / `32`) | `[s].000"s" ` | Total seconds with three fractional digits |

Raw source seconds and typed formula caches are retained; XLSX uses numeric day serials. Text tables and Reader use the shared numeric or elapsed formatter. The [native Numbers 14.5 second-range reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-seconds-v14.5.json) pairs 42 selected cells and 21 formulas with Apple XLSX and three inspected PDF pages. Saved/reopened XLSX retains the native format codes and caches, and edited operands recalculate. Integral shared display matches the native PDF. Fractional seconds differ: Numbers rounds `3661.75` to `1h 1m 2s`, while the shared formatter displays `1h 1m 1s`. Native Excel also differs, including hashes for negative hour/second and minute/second formats. `IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED` and text-display diagnostics retain these display limits; the complete composed PDF route retains all rows and formula results across six pages, versus three native positioned-sheet pages. This reference does not establish exact cross-application appearance.

The [native millisecond-range reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-duration-milliseconds-v14.5.json) adds fixed hour/millisecond, minute/millisecond and second/millisecond ranges. All 42 typed values, 21 formulas and native XLSX codes survive save/reopen and operand recalculation. Numbers displays separate components such as `1h 1m 1s 750ms`; its XLSX export and the shared formatter use `1h 1m 1.750s `, including a terminal literal space. This difference retains the duration-display approximation. The complete PDF route retains every row and formula result across six inspected pages. Native Excel appearance for this reference remains unqualified.

Whole-day values match the [independent duration-range evidence](../OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/duration-ranges.json), including negative durations. Fractional days follow destination numeric rounding, which can differ from source truncation. The [hour/minute evidence](../OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/duration-format.json) agrees with the paired Numbers 14.5 XLSX code and cache. Week/day/hour decomposition has no equivalent elapsed XLSX week/day tokens. Other ranges, millisecond-only totals, label styles, automatic selection and locale remain unqualified.

Selected text and Boolean formats resolve the existing format catalog only after complete cell decoding. A unique native text format (`260`) or Boolean format (`1`) with no additional settings preserves the editable default. Explicitly Text-formatted blank cells expose `IWorkNumberFormatKind.Text` and preserve XLSX Text format (`@`) after save/reopen, matching the nine blank Text cells in the Numbers 14.5 native export. Dormant text catalogs do not apply a format; custom text settings remain unassessed. Missing, ambiguous, malformed or other format declarations retain `TextFormat` or `BooleanFormat` evidence and require the existing partial policy. Recoverable text, Boolean, date/time and duration values remain typed; this assessment does not apply other custom format settings. Catalog and row evidence share existing reference/declaration budgets, and parser limits remain fatal. The [independent scalar-format corpus](../OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cell-format-selectors.json) retains 15 representative declarations and selector counts from four unchanged native packages; it does not qualify Apple exports or appearance.


Selected style-value failures use `InvalidValue` with their physical property path, such as `11/3` for font size, `11/7` or `11/26` for an unresolved foreground/background color value, `12/1` for alignment, `1/1` for a style name and `2` for hyperlink text. Counts describe that property’s outer values, including zero for missing required hyperlink text. Ambiguous Boolean values and clear-font/color selectors do not replace qualified inherited values. Rejected list vectors retain qualified parent metadata rather than removing invalid entries and shifting later levels. List evidence uses root fields `11` for label types, `13` for indents and `16` for labels; a packed label-type field counts as one outer value rather than its decoded element count.

Catalog keys are indexed before values are trusted. A malformed value with a readable, distinct key leaves healthy sibling values recoverable. Duplicate keys remain unresolved, including when a duplicate value is malformed; an unreadable entry or key leaves uniqueness unassessed and prevents resolving any catalog value. Formula cells retain valid caches when their expressions cannot be resolved. Catalog evidence uses `3[n]` for an unreadable physical entry, `3[n]/1` for key metadata, `3[n]/3` for invalid plain-string values (`InvalidValue`), and `3[n]/5` for rejected or unreadable formula messages. The catalog-entry budget includes unreadable entries. Rich-text value, wrapper, storage and style traversal stays limited to entries selected by decoded cells.

Sizing-bucket evidence uses `2[n]` for unreadable physical headers, `/1` for invalid or duplicate indexes, `/2` for invalid sizes, and `/3` for unsupported visibility values. `InvalidValue` covers size and visibility failures. An unknown index or unresolved bucket makes overrides on that axis untrusted, including overrides in other selected buckets; the other axis remains recoverable. Invalid values at readable, distinct indexes leave healthy sibling sizes available. Native zero continues to select the table default. Repeated bucket references retire their readable indexes while retaining unrelated overrides, and are reported at model path `4/1/2`; shared physical header paths report once. Explicit partial conversion retains recovered sizes using each destination owner's existing default and geometry rules.

Merged-range evidence uses model paths `47` and `47/2` for rejected or unreadable owner/store messages, `47/2/3[n]` for unreadable physical pairs, and `47/2/3[n]/2` for formula declarations with unsupported bounds or conflicts. Counts describe the outer field at that path. Conflicting rectangles are all retired; disjoint, valid merges remain recoverable through explicit partial conversion. A decoded out-of-bounds rectangle disqualifies only intersecting merges and is never exported as a clipped merge. Exact valid duplicates normalize to one merge; every physical alias retains evidence when that merge conflicts. An unreadable range can conceal any overlap, so cells remain recoverable without applying merges. Configured pair, syntax-node, protobuf and evidence limits remain fatal. Destination guards still protect populated covered cells.

Selected native table tiles retain unreadable payloads as `$`, rejected row sets as `5`, and unreadable or invalid rows as `5[n]` using their physical one-based positions. Tile-list selection failures retain model paths `4/3/1[n]`. Invalid row indices, modern buffer/offset metadata and unsupported legacy storage retain row-level evidence without inventing cell counts. Readable sibling rows survive malformed entries; the table remains incomplete and requires explicit partial reconstruction for editable output. Shared tile paths report once per projection. Rejected envelope/count bounds are checked before nested rows are parsed; configured parser limits remain fatal on content actually parsed.

Pages body, shape, header/footer, Keynote drawable/note and selected table-cell rich text also report failed references in decoded paragraph, list, character-style and hyperlink tables, plus parent references in styles traversed for that text. Each valid-offset decoded attribute entry is assessed, including end-of-text declarations; its path preserves the physical table-entry position. Unused style records are not traversed. Malformed attribute containers and entries with invalid offsets remain unassessed. Numbers shape text is decoded as plain text, so its formatting references remain outside this evidence. Table rich-text catalogs index all declared entries under the catalog budget, preserving structural and duplicate-key checks. Text, styles and wrapper/storage references are resolved when a decoded cell selects a catalog entry; aliases share storage decoding while each cell use consumes its text budget. Unused rich-text entries do not consume text budgets or add reference evidence. Unparseable table containers and unsupported table dependency fields remain unassessed.

`SourceCellIssues` identifies materialized selected cells whose storage or value could not be decoded, with table identity, one-based coordinates and the decoding message. `IWorkTableCell.HasDecodeError` distinguishes these failures from recovered native `#ERROR` markers; it does not assess every field of a successfully decoded cell. The inventory survives visual fallback and disabled record preservation, stays within the source-wide materialized-cell budget and reports `Unassessed` through `IWORK_SOURCE_CELLS_UNDECODED`. Undecoded formula cells also retain their separate formula/cache assessment. Unreadable rows, rejected storage and inactive tables do not establish cell coordinates and add no guessed cell counts. These source assessments do not describe destination cell retention.

Qualified cross-table cell and rectangular-range references remain incomplete until their native table identities can be resolved. Endpoint pairs require the same target identity; other reference node families remain unqualified. Their typed cached values stay available; dropping a table qualifier never produces a complete local expression. A merge qualifier must identify its owning table through the qualified four-word UUID representation; references to another table and ambiguous declarations retain source evidence instead of defining a local merge.

`FormulaCells` identifies source formula declarations by table identity and one-based coordinates. `ExpressionIsAssessed` distinguishes reconstructed or unresolved expressions from undecoded cell contents. Cache status is complete, partial, missing, approximate for a high-precision decimal or generic error marker, or `Unassessed` when the cell cannot be decoded. A supported version-5 header with the formula flag establishes `SourceFormulaIsDeclared`, even if later value decoding fails; the cell keeps its existing error output. Unsupported versions and unreadable headers do not establish formula presence. `FormulaSummary` includes separate unassessed expression and cache counts. These source assessments survive visual fallback and disabled record preservation; they do not establish destination formula retention or cache freshness. The inventory remains bounded by materialized-cell limits and excludes inactive tables.

Every package entry and every decoded IWA payload remains available as defensive bytes on `IWorkSourceDocument`. Import reports expose source payloads through `PreservedRecords` when `PreserveSourceRecords` is enabled. `PreservedRecordCount` counts source payloads regardless of that detail setting. `UnassessedRecordCount` includes consumed and auxiliary records whose field-level fidelity has not been assessed; it is not an omission count. `IWORK_RECORD_FIDELITY_UNASSESSED` uses `OfficeConversionLossKind.Unassessed`, so strict no-loss policies reject it without claiming those records are missing content. The destination DOCX, XLSX, or PPTX contains the supported reconstruction or visual fallback; it is not a lossless iWork package rewrite.

Loaded Pages, Numbers and Keynote packages remain read-only. Native creation uses separate `IWorkKeynoteDocument`, `IWorkPagesDocument` and `IWorkNumbersDocument` models; these do not provide save-back or preservation of edited source packages.

See the [iWork support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.iwork-support-matrix.md) for the version corpus, limits, semantic coverage, and known boundaries.

`IWorkTableCell.NumberFormat` exposes supported modern number, percentage, currency, and scientific semantics: decimal places (`null` for automatic), digit grouping, and negative-value style. Scientific formats support the default minus style without grouping; decimal places describe the mantissa. Currency metadata also exposes the three-letter uppercase source `CurrencyCode` and `UseAccountingStyle`; the code does not imply a symbol or locale. It does not change `Value`, `DisplayText`, or `CachedDisplayText`. The Excel adapter applies this metadata to numeric XLSX cells; the Word and PowerPoint adapters apply supported invariant display text and red format color while reporting approximation for locale, automatic precision and source appearance. Cells whose display cannot be safely formatted retain raw text and report omission. Unsupported selected numeric formats retain source diagnostics and physical catalog-declaration evidence when a declaration can be identified.

Fraction formats expose `Kind == IWorkNumberFormatKind.Fraction` and `FractionAccuracy`: one-, two-, or three-digit denominators, or fixed halves, quarters, eighths, sixteenths, tenths, and hundredths. `DecimalPlaces` is `null` for fractions and does not mean automatic decimal formatting. The qualified subset uses minus negatives without grouping; other controls remain unsupported. The Excel adapter uses mixed-fraction formats and reports `IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED` for rounding, normalization, and spacing differences.

Finite Decimal128 values remain numeric even when their coefficient exceeds fifteen significant digits. Such cells set `NumericValueIsApproximate`, retain the exact normalized source value in `SourceNumberText`, and emit `IWORK_TABLE_NUMERIC_VALUE_APPROXIMATED` with `Approximation` fidelity. Formula caches receive the same metadata and approximate cache assessment. The retained text counts toward projection text limits. Nonfinite, noncanonical, overflowing, and nonzero underflowing values remain decode failures.

## Target frameworks and dependencies

`OfficeIMO.IWork` targets .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2 on Windows. The source reader depends only on `OfficeIMO.Core`; its IWA, Snappy, protobuf-envelope, and package readers are first-party implementations. Destination projection is opt-in through `OfficeIMO.Word.IWork`, `OfficeIMO.Excel.IWork`, or `OfficeIMO.PowerPoint.IWork`.

### Table-cell comments

`IWorkTableCell.Comment` exposes qualified root comments with exact plain `Text`, display `Author`, UTC `CreationDateUtc`, and native comment/author record identities. Comments on empty cells remain materialized and use the existing cell, catalog-entry and text budgets. Catalog, comment and author records are traversed only when selected by a fully decoded cell. Unresolved selectors, replies, invalid timestamps and unsupported record shapes retain `UnsupportedFeatures.Comment` and source diagnostics.

Numbers-to-XLSX conversion writes qualified roots through the Excel threaded-comment owner and retains text, author, timestamp and cell address after reopening. DOCX preserves roots through the Word batch comment owner with cell-scoped paragraph anchors, including empty commented cells. Text or author names requiring normalization use destination fallback. PPTX table conversion requires explicit partial reconstruction to omit cell comments. Reader emits separate comment blocks, chunks and `table.comment` metadata with native identities and source cell anchors; comments outside its table limits retain explicit omission counts. Author colors, collaboration identifiers, native UUID mapping, replies, Apple export equivalence and rendered comment appearance are outside this contract.

### Pages section headers and footers

Pages section projections distinguish stored header/footer templates (`HasFirstPageTemplate`, `HasEvenPageTemplate`, `HasDefaultPageTemplate`) from their selection settings (`DifferentFirstPage`, `DifferentOddAndEvenPages`, `HideFirstPageHeadersAndFooters`). Stored alternate templates may be inactive. `SelectedHeaderContents` and `SelectedFooterContents` expose the semantic inventory selected by those settings. Reader extraction and conversion-report counts use that inventory; it does not establish how many rendered pages use each template.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 0 | 1 | 0 | 0 | 1 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
