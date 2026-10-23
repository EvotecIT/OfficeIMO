# OfficeIMO.Word.Pdf - Word/PDF conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Pdf)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Pdf)

Word-to-PDF conversion and page images share the adapter in `OfficeIMO.Word`. This package adds import of parser-supported PDF logical content into editable Word documents and references the Word and PDF packages. Forward conversion APIs keep the `OfficeIMO.Word.Pdf` namespace; PDF layout, reading, diagnostics and writing stay in `OfficeIMO.Pdf`.

## Legacy DOC to PDF

`LegacyDocPdfConverter` composes bounded binary DOC import with the Word PDF adapter. It blocks known import loss by default and retains the import report separately from PDF rendering diagnostics:

```csharp
using OfficeIMO.Word.Pdf;

using var source = File.OpenRead("archive.doc");
var conversion = LegacyDocPdfConverter.ToPdfDocumentResult(source);
conversion.SaveResult("archive.pdf").RequireSuccess();
foreach (var report in conversion.SourceConversionReports)
    foreach (var finding in report.FidelityDiagnostics)
        Console.WriteLine($"{finding.Code}: {finding.Message}");
```

Pass `lossPolicy: OfficeIMO.OfficeConversionLossPolicy.Allow` only when accepting the reported import reductions. Import errors still block output. The adapter always collects unsupported-content findings, even if the supplied import options disable reporting. Its fidelity is limited by both the legacy importer and the Word PDF renderer; it does not guarantee exact Microsoft Word pagination or rendering of every binary DOC feature. Supported binary DOC SECTIONPAGES fields retain their editable instruction and cached result during import; PDF header/footer fields use calculated section counts during pagination.

Rejected imports throw `OfficeConversionException`; its `Report` retains the available import findings. Catch that exception to inspect a blocked conversion without rendering a PDF.

## Install

```powershell
dotnet add package OfficeIMO.Word.Pdf
```

## Quick start

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Create("report.docx");
document.AddParagraph("PDF export").Style = WordParagraphStyles.Heading1;
document.AddParagraph("This document is exported through OfficeIMO.Pdf.");

document.SaveAsPdf("report.pdf");
```

## Examples

### Export with page and metadata options

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new WordToPdfOptions {
    Orientation = OfficePageOrientation.Portrait,
    Margins = PageMargins.UniformCentimeters(1.5),
    Title = "Customer proposal",
    Author = "Evotec",
    IncludePageNumbers = true,
    PageNumberFormat = "Page {current} of {total}"
};

document.SaveAsPdf("proposal.pdf", options);
```

### Refresh date fields before export

PDF export uses the field results stored in the Word document by default. To refresh dynamic `DATE` and `TIME` fields for the PDF, set `DateTimeFieldUpdateOptions`:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("report.docx");
document.SaveAsPdf("report.pdf", new WordToPdfOptions {
    DateTimeFieldUpdateOptions = new WordFieldUpdateOptions()
});
```

`DATE` and `TIME` use the current local clock during the refresh. Set `CurrentDateTime` on `WordFieldUpdateOptions` when the export needs a fixed value. The refresh updates the in-memory Word document; it does not save the `.docx` file or change other fields. `CREATEDATE` and `SAVEDATE` continue to use their cached results unless explicitly refreshed through the general Word field API.

Fields that display text need a cached result unless PDF export calculates them, such as page counters in headers and footers. When a recognized display field has no stored result, export reports `NativeFieldResultUnavailable` with warning severity and `LossKind = Omission`, so the conversion report's `HasLoss` is `true`. Explicitly stored empty results remain empty. Hidden fields and instructions that produce no display text do not add this warning.

### Word-to-PDF layout

Inline date pickers, drop-down lists and combo boxes retain their position among text in body paragraphs and table-cell paragraphs. Date pickers become editable PDF text fields; list and combo controls become editable PDF choice fields. Cached display runs retain their effective fonts, sizes, colors and ordinary line wrapping. The initial appearance uses that cached text, while the field keeps its logical date value or selected option. A missing date or choice cache can use the current logical display value. Wrapped controls can have several widgets belonging to one logical field. Filling the field replaces its cached appearance without leaving the original text underneath. A control with an effective run highlight or background retains static text and reports `NativeInlineControlBackgroundStatic`: its background must paint beneath surrounding paragraph glyphs, while PDF widgets paint above page content.

Inline checkboxes retain their source position and line metrics. Their current appearance uses the cached glyph; the opposite state uses the declared checked or unchecked glyph and font when those can be represented by the available PDF fonts. Editing a PDF checkbox retains its fixed widget rectangle. The alternate glyph's measured paint, including overhangs and negative bearings, fits inside that rectangle without reflowing surrounding text. Its baseline stays fixed when the ink can fit around that baseline; a glyph extending across a baseline at the rectangle's edge moves inside the rectangle so it remains visible. An unavailable state uses the standard checkbox appearance and reports `NativeCheckBoxGlyphAppearanceUnavailable`. Font substitution can change the glyph's appearance or dimensions. Controls without a usable source run or display value, controls containing complex fields, and controls in paragraphs containing equations retain the existing fallback paths. Checkboxes without cached text retain the standalone editable-field fallback. Form controls in headers and footers remain static content. This mapping does not make every Word content-control kind an editable PDF field.

Table borders follow the Word style and direct cell settings: `nil` suppresses a shared edge while `none` yields to the opposing border. Set `DefaultTableBorders = true` only when you want a fallback grid on otherwise borderless tables.

Collapsed borders and authored cell margins contribute to row spacing and page or column fitting. Ordinary flow tables clear their leading and trailing border paint; continued rows use the border clearance at their new frame, and repeated headers retain their source spacing. Text in vertically merged cells that cross a page boundary can still appear on the earlier page. That fragmentation remains outside the qualified layout mapping.

Set `cell.TextDirection` to `WordTextDirection.TopToBottomRightToLeft` or `WordTextDirection.BottomToTopLeftToRight` to export Latin cell text with a quarter turn. Borders and padding retain their physical page axes, and inline pictures remain upright at their authored size. Automatic turned text rows use paragraph-mark size and line spacing instead of increasing their height with the visible run font; exact row heights retain their source box. Automatic column widths use the visible text's cross-axis line box. A turned cell's text, objects and links are consumed on its first row fragment when a neighboring cell continues onto another page. Flow, column and canvas tables share the resolved text shrink scale.

East Asian vertical glyph layout and exact automatic sizing for cells with hidden end marks or embedded objects remain outside the qualified mapping. Font metrics and source compatibility settings can still affect absolute cell geometry.

Inline pictures in table cells retain their position among the paragraph's text, including multiple pictures in one run. Their authored display dimensions and alternative descriptions pass to the shared PDF layout; exact-height lines can clip them. Mixed picture runs retain visible text, and hidden runs suppress their pictures in body content, tables, headers and footers. Nested-table content is flattened, so preserving its pictures does not preserve nested table frames.

Picture crops, rotation and horizontal or vertical mirroring pass through the shared PDF image style in body paragraphs, table cells, headers and footers. Supported inline pictures in footnote and endnote paragraphs use the same image-style projection. Supported floating pictures retain those transforms at their resolved page anchors. Crop coordinates select the source image before the display transform; the authored layout box continues to determine line and table spacing. Negative crop outsets and collapsed source rectangles report `NativeImageCropApproximated` and fit the complete source to the authored frame. Floating pictures inside notes and unsupported drawing effects remain outside the supported note profile.

Ordinary inline DrawingML pictures in body paragraphs retain their authored display frame, including pictures wider or taller than the text area and image-only paragraphs. Paragraph alignment and surrounding text determine placement; the page boundary clips paint outside the page. A tall picture line can move to a fresh page or column, where a picture taller than the available frame retains its size. Exact-height body lines clip the picture paint to their paragraph segment without applying that picture clip to adjacent text. Pictures in table cells, text boxes, picture content controls, equation paragraphs and running stories retain their separate sizing and fallback paths.

Ordinary underlining includes spaces between words. Word's explicit *underline words only* style continues to leave those spaces clear.

Word positional tabs retain left, center and right alignment relative to margins or paragraph indents, including running headers and footers with page fields. Center and right tabs align the complete field up to the next tab or line break, including spaces and formatting changes, so labels such as `Footer Right` stay aligned as a phrase.

Custom tab stops combine document paragraph defaults, paragraph styles and direct formatting. A more specific declaration replaces the alignment and leader at the same position; a clear suppresses the inherited stop. Body, table-cell, note and running-story paragraphs use the same effective tab resolution.

Explicit paragraph tab stops retain their source positions when a hanging first line starts before the continuation indent. Stops at or before the continuation text frame remain available while they are ahead of the current cursor; following lines skip stops they have already passed. This applies to body, note and table-cell paragraphs. A hanging indent also supplies an implicit stop at the continuation text origin. Older compatibility layout suppresses that stop when `DoNotUseIndentAsNumberingTabStop` (`noTabHangInd`) is enabled; Word 2013-or-later layout ignores the stored flag. Default tab intervals retain their separate behavior.

Word text extraction honors `xml:space`: leading and trailing XML whitespace is significant only when the text node declares `preserve`. PDF export omits insignificant text-node edges. Internal spaces, non-breaking spaces, explicit tabs and line breaks retain their meaning in the Word model. Reading text and exporting PDF leave the source XML unchanged; text authored through the Word APIs preserves supplied spaces. The PDF flow engine can still collapse preserved edge or consecutive spaces in body text and table cells.

Font sizes preserve half-point values, including 10.5 pt. When the run and its styles omit a size, conversion honors the document default. An existing `docDefaults` element without a size uses Word's 10 pt fallback; a document without `docDefaults` uses 12 pt. OfficeIMO-created documents declare an 11 pt default and retain that size.

Each text run resolves its font through its own formatting, character style, paragraph style, table style, and document defaults. A preceding run's explicit font does not change the font of an unformatted sibling. This also applies to source page fields in headers and footers.

Document-wide bold, italic, underline, strikethrough, capitalization, baseline and text-color defaults apply when higher-priority formatting does not declare a value. Direct formatting and style resets, including automatic text color, stop that inheritance. Small capitals retain full-size original capitals and smaller capitals for lowercase text, including styled body text, table cells and running headers or footers. Their effective sizes participate in wrapping and hyperlink placement.

Supplying an empty `PdfOptions` object, metadata, or header/footer presentation settings preserves the document's inherited body font and size. Registering a reusable named font family makes it available to matching text without selecting a different body font. Generated headers and page numbers retain their explicitly selected fonts, including when the body embeds a font with the same standard-font alias. Explicit PDF body-font and default-size assignments retain their precedence. Rendering profiles keep the same distinction between document typography and presentation overrides.

Run character width (`w:w`) and tracking (`w:spacing`) flow through document,
table, paragraph and character styles, with direct formatting and explicit resets
taking precedence. Body text, list markers, table cells, hyperlinks and header/footer
paragraphs, including generated page fields, retain the effective values. Width scaling preserves font height;
tracking remains an advance in page points. These saved run settings do not
implement the separate dynamic table-cell Fit Text policy. Font substitution can
still change exact line endings and glyph appearance.

Run kerning (`w:kern`) inherits through document, paragraph, character, and table styles. A positive minimum font size enables kerning at that size and above; zero disables it. Compatible Latin text runs retain pair kerning across color and link changes in paragraphs, table cells, headers, and footers. Font or size changes, disabled kerning, and text separators end that boundary. Numbering labels retain their unkerned advances, including labels with a numbering-level or paragraph-mark kerning setting. Their measurement and rendering use the same feature settings.

### Preserve document fonts

Balanced conversion embeds subsets of installed document fonts by default. Word, Excel and PowerPoint use the same policy. On macOS, discovery includes fonts supplied by installed Microsoft Office applications. Missing Calibri and Cambria prefer Carlito and Caladea; Arial, Times New Roman and Courier New prefer the corresponding Liberation family. Aptos uses a close installed sans-serif alternative. Further installed alternatives precede the final PDF standard-font fallback. Substitution diagnostics name the requested and selected families; compatible-family selection expresses intent, not proof of identical font-version metrics.

The conversion report's `font-embedding-cost` entries record each emitted font program's original, subset and compressed stream sizes. These are font-program payload bytes; PDF dictionaries and character maps add further overhead. No font files are bundled with this package.

For host-independent conversion, provide a fixed rendering profile or select the portable policy:

```csharp
var options = new WordToPdfOptions {
    ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
};
document.SaveAsPdf("portable.pdf", options).RequireSuccess();
```

Both system-font and document-font embedding must be allowed to discover installed document families. `FontFamily` selects a conversion-wide default. Local-file and remote-resource access remain separate policy settings.

### Page images with PDF pagination

Render the conversion result when PNG, JPEG or SVG pages must share the PDF's line breaks, tables, fonts and pagination:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

var conversion = document.ToPdfDocumentResult();
conversion.Save("document.pdf").RequireSuccess();
conversion.ToImages(new PdfImageExportOptions { TargetDpi = 144 })
    .AsPng()
    .Save("pages");
```

Use `.AsJpeg()` or `.AsSvg()` for those formats, or `conversion.ExportImages(format, options)` for in-memory results. The shared PDF renderer projects the laid-out pages into Drawing scenes and retains conversion and renderer diagnostics. It does not lay out the Word document again or shorten paragraphs with an ellipsis. Managed PDF rendering still has the [PDF renderer's supported-feature limits](../OfficeIMO.Pdf/README.md).

High-resolution nearest-neighbour pictures can exceed the SVG rectangle-vectorization budget. Use PNG or JPEG for those pages until bounded SVG image fallback is supported.

`document.ToPng()`, `ToJpeg()`, `ToSvg()` and `ExportImages(...)` use this same layout and page renderer. Set `WordImageExportOptions.LayoutOptions` when sharing customized conversion settings. The forward Word page adapter and options live in `OfficeIMO.Word`; this package also provides PDF-to-Word import.

SmartArt uses the same drawing projection for PDF and Word images. Supported persisted preset shapes retain node positions; process, cycle and hierarchy diagrams have node-text fallbacks. Header and footer diagrams resolve their parts and cached drawings from the containing story and repeat with its running content. These projections report their approximation because native diagram styles, effects and complete connector geometry are not reconstructed.

Positioned tables in ordinary document flow preserve page, margin, or text anchors, explicit offsets, and text clearances. Following paragraphs use the available space beside the table and return to full width below it. Headings, lists, images, and other structured blocks move below an intersecting table. Positioned tables in multi-column sections retain an approximation warning.

Section columns use the shared PDF flow engine for equal and unequal widths, continuous section transitions and final-page balancing. Paragraph keep, widow and spacing settings remain active inside the column frame. Tables split at row or supported cell-content boundaries, and repeated headers stay with their first body row. Signed table indentation is retained during DOCX/native DOC authoring and left-aligned PDF placement; an explicit zero indent overrides inherited table-style indentation. Centered and right-aligned inline tables ignore the authored left indent.

Inline tables with collapsed borders retain the horizontal placement rules of their Word compatibility mode. Legacy left and right alignment uses the first cell's leading margin to place the grid, including cell margin overrides. Centered tables keep the grid centered. Modern Word layout aligns the outside border edge with the body frame. These translations preserve column widths and continuation indents. Tables using an explicitly configured PDF default style retain that style's presentation placement.

Automatic tables with automatic cell widths size their columns from unwrapped content and cell padding, within the available page or column frame. A stored grid alone does not impose a preferred table width. An automatic-width cell in a later row can grow an earlier absolute column preference to fit its text and inline pictures. A following absolute cell restores wrapping at the authored preference while retaining unbreakable content minimums. An omitted cell width follows the automatic path; `nil` retains the cached grid preference. An authored absolute table width retains its frame even beyond the text area and counts cell spacing once. Within that frame, absolute cell preferences shrink through the space between their preferred widths and content minimums, then grow proportionally when the preferences fit. Automatic cells absorb the space remaining after absolute preferences are satisfied. Percentage table and cell preferences retain their sizing constraints.

Fixed tables retain their authored width even when it extends beyond the page or column text area. Absolute and percentage table widths take precedence over the stored grid; a fixed table with automatic width uses its grid. Percentage widths reserve the effective outer border paint before scaling the text area. For legacy-layout tables with collapsed borders, the percentage reference frame also includes the effective first and last cell margins; modern-layout and separated-border tables use their own frame rules. Table styles, direct table settings and cell overrides supply those margins. An unstyled, left-aligned legacy percentage table uses the first cell margin as its default text indent; a resolved table style or explicit indent retains its own value, including zero. These rules also apply to tables in headers and footers.

A declared fixed width includes outer cell spacing, and very narrow cells retain their margins, border paint and minimum text frame. Merged-cell width and margin constraints apply to the combined span while retaining unequal underlying columns. Fixed tables expand when their effective margins require a larger text frame. Center and right alignment preserve the resulting position across the text-area edges instead of narrowing the columns.

Tables with an absolute cell-spacing declaration retain an outer perimeter and separate cell borders, including table shading in positive gaps and different border colors, widths and supported patterns. An explicit zero retains separate border geometry; omitting spacing uses collapsed borders. The frame follows table alignment, indentation, merged cells, page and column continuation, and repeating headers. Cells in a row share the effective top and bottom margins, and minimum row heights reserve those margins and border paint. Exact row heights retain partially visible text through clipping inside the cell borders; minimum heights allow content to grow. Set `table.StyleDetails.CellSpacing` in twips to author the spacing. An imported explicit automatic or percentage spacing value clears inherited spacing and uses collapsed borders. Floating tables and nested table layout remain subject to the native engine's supported paths.

For unspaced tables, a physical exact row box includes Word's largest effective bottom cell margin. Vertical merge continuations still contribute that margin. Direct cell values override conditional table-style padding, including an explicit zero that clears an inherited margin.

Paragraph border spacing positions the border outside the text frame without adding horizontal text padding. Borders spanning columns or pages keep their side strokes; their top belongs to the first fragment and their bottom to the last. Native DOC compatibility can produce different fragment capacities from modern DOCX. These mappings preserve the supported source settings; font substitution and unsupported layout features can still change pagination.

Line spacing follows document defaults, table styles, paragraph styles and direct formatting, including built-in headings. Automatic spacing uses the effective paragraph and run fonts after substitution or embedding; exact and minimum spacing retain their point units. Rich body and table paragraphs use their own font size during measurement and pagination. A large run on another line or a large paragraph mark does not impose a minimum font size on every rich line. Authored line breaks retain their run formatting, so larger blank lines can expand minimum spacing. Exact spacing keeps a fixed advance; minimum spacing can expand for larger runs or inline elements. An authored line value without a rule uses automatic spacing. A rule without a numeric line value inherits the complete spacing pair. When the effective hierarchy declares no line value, empty paragraphs use the effective font's natural line height in body content, table cells, headers and footers. Font substitution can still change line advances. Exact-height clipping for other objects and exact Word pagination remain limited.

OOXML leaves omitted document paragraph defaults to the application. Declare document or paragraph spacing when reproducible layout across different applications and profiles is required.

These spacing policies apply to `Both` and `Distribute` paragraph alignment. Other justification modes retain their existing fallback alignment.

Justified Word paragraphs expand ordinary spaces using preferred shares from each source space's measured advance, then share the remaining expansion per source space. This applies to paragraph fragments continued on another page or in a footnote. The source `doNotExpandShiftReturn` setting controls expansion at an explicit line break. Distributed paragraphs include explicit-break and final lines. They expand ordinary spaces first, then use the resolved font's available average advance for preferred expansion between shaped grapheme clusters, then share the remaining expansion equally. Missing average metrics leave those cluster boundaries for the final equal share. Combining sequences and shaped ligatures stay intact. Tabs keep their anchors in either mode; only word separators in the final left-tab field expand. Word justification uses the same source-space weights in that field as in an otherwise equivalent indented paragraph. Body, table-cell and supported note paragraphs share these policies.

Imported justified paragraphs can also compress eligible ordinary spaces by up to one quarter of each space's natural measured advance. The allowance follows each source space's font and size and adjusts spacing without scaling glyphs. Whole-word admission includes words split across formatting runs before deciding whether they fit. Nonbreaking spaces and tab advances do not supply this compression allowance.

For word justification in a fixed line frame, line fitting compares expanding the preceding legal word break with compressing the complete next-word candidate, relative to each alternative's source-space budget. A feasible compressed candidate can still wrap at the earlier break. The comparison is limited to whole words that fit within the known fresh-line width. PDF paragraphs without Word spacing enabled, character distribution and dynamically changing line frames retain their existing break selection.

These spacing and break-selection rules do not guarantee the same line membership or pagination as Word. Exact behavior across wider font, compatibility and layout profiles remains a qualification limit, even when source fonts are available. Font substitution introduces a separate source of metric differences.

Paragraphs with exact, automatic or minimum line spacing retain top, center, bottom and baseline font alignment for each text run and numbering label. Automatic and minimum lines include the aligned runs when calculating their font box. Alignment inherits independently of the spacing value through document defaults, paragraph styles and table styles, including conditional table styles. Embedded font device metrics preserve the relative positioning of mixed font sizes. Absolute line positions and accumulated spacing in long paragraphs can still differ from Word. Automatic line spacing in paragraphs mixing Latin and East Asian fonts has a measured remaining difference in line height and following-text placement, even when both fonts are available and the glyphs share a baseline. Native DOC saving carries document-default alignment into the style hierarchy.

Automatic spacing reserves the final line's font box at a page or column boundary while retaining the full distance between successive baselines. Paragraphs, headings, lists and floating-table text flow use this boundary measurement, including kept paragraphs that fit the frame. Exact spacing retains its fixed advance between baselines. Minimum spacing retains its full height requirement.

Section line grids constrain automatic and minimum line boxes and position their baselines within the expanded height. Exact spacing retains its authored height. Grid snapping follows document defaults, table styles, paragraph styles and direct overrides; table paragraphs also honor the `adjustLineHeightInTable` compatibility setting. Before and after paragraph spacing expressed in hundredths of a line uses the active section grid pitch, including paragraphs that disable snapping, or a 12-point unit when no line grid is active. Inherited line-unit spacing takes precedence over point-valued spacing. The first body paragraph retains its spacing before at the document start; normal page breaks suppress that spacing at the next page start. Character-grid placement remains limited.

Numbered and bulleted paragraphs retain their collapsed before/after spacing when grouped into PDF lists. Contextual spacing follows each paragraph's setting and style identity, including lists that change this setting between items. A suppressed trailing space still participates in the next paragraph's spacing calculation. The same calculation applies to body paragraphs and paragraphs inside table cells.

Numbering labels contribute their font ascent to the first line while the body text supplies its descent. Large labels enlarge that line without changing the font metrics of wrapped continuation lines. Multiple line spacing scales the body metrics and retains the label's additional ascent. Automatic spacing below one line also reduces the body ascent used to place the baseline. Empty numbered items, including items containing only spaces or tabs, retain the paragraph mark's font and size. A trailing manual line break retains its final paragraph-mark line in body and table paragraphs. Label font selection respects the document's registered font families, including families assigned to different PDF font slots.

Ordinary text and field results inherit hidden visibility from direct formatting, character styles, paragraph styles and document run defaults, in that order. An explicit visible setting overrides a hidden inherited setting. Paragraph marks inherit paragraph-style and document visibility independently of the text runs. Empty paragraphs retain a blank line when their paragraph mark is visible, including paragraphs whose text runs are all hidden. Their line metrics use the paragraph mark and inherited typography. A hidden mark removes the empty line and its paragraph spacing; an explicit visible mark overrides inherited hidden visibility. An image, shape or inline group rendered in paragraph flow reserves its object height once; a following empty paragraph retains its own blank line. Adjacent text paragraphs in the body, columns, block content controls and individual table cells join across hidden paragraph marks. Joined text uses the first paragraph's alignment and spacing before, the final paragraph's spacing after, and each source run's own inherited typography. Lists, headings, decorated paragraphs, explicit page or column breaks, and object paragraphs retain separate boundaries and emit a warning. Cell joining also retains boundaries when an explicit PDF table style supplies unresolved paragraph typography. Joining in headers and footers remains limited.

A referenced header or footer part with no child content reserves an implicit line using the default paragraph's font, line-height, mark-visibility and spacing-before settings; its spacing after is omitted. An explicit empty paragraph keeps its own spacing. Referenced stories retain their distance from the page edge even when hidden paragraph marks remove their line boxes. A missing story reserves no additional space. These distinctions apply to default, first-page and even-page stories. Mathematical and object-specific visibility inheritance remains outside the qualified ordinary-text mapping. Oversized ordinary stories retain an approximate text fallback; exact native clipping, wrapping and placement remain outside that fallback's qualified mapping.

Ordinary headers and footers can retain non-wrapping DrawingML picture anchors with one picture payload and resolvable page, margin or paragraph positions. Horizontal column anchors require a single-column section. Anchored picture paint does not reserve its full height as flowing story content; the anchor paragraph still contributes its own line. Inline pictures retain their measured line height. Wrapped pictures, unresolved anchors, mixed text-box stories and caller-configured header/footer presentation overrides remain subject to their existing projection and approximation paths.

Empty referenced stories retain their source paragraph geometry when PDF header or footer presentation options are configured. Visible sibling variants and generated page numbers keep their configured typography and physical-margin positioning. A body paragraph can retain its header and place one line beside an expanded footer when that line fits within the physical bottom margin. Imported exact-height body paragraphs also retain their header geometry when the first line extends beyond that margin: each fresh page or column admits one line at its authored baseline. A line taller than the physical column body frame continues on a new physical page; a continuous column section moves past existing page content before admitting such a line. Lines that fit the frame retain ordinary column continuation and widow/orphan rules. Lists, tables, decorated containers and ordinary PDF paragraphs keep their existing flow bounds.

With generated page numbers enabled, text-free first-page and even-page footers remain blank, including childless parts, empty paragraphs, and hidden-only content. A missing default footer uses the normal page-number projection without reserving a source-story frame. Visible sibling PAGE and SECTIONPAGES fields keep their source numbering and do not gain a second generated number.

Later sections inherit omitted default, first-page, and even-page header/footer references from preceding sections. An explicit blank part stops that inheritance. Compatible continuous sections sharing the same effective stories remain on one page. Their SECTIONPAGES fields retain each source section's page count, including pages shared at a continuous boundary. The section at the physical page start owns that page's running fields, including an empty initial section. First-page variants follow the first physical page of a section; odd/even variants and their text or image watermarks follow the displayed page number, including numbering restarts. A missing even-page header suppresses the default header's watermark when the document enables separate odd/even stories.

Missing table margin sides retain Word's standard zero vertical and 5.4-point horizontal padding. Document table styles, direct table margins and individual cell overrides retain their precedence, including explicit zero values; configured PDF defaults supply missing sides when selected. Unnamed tables use the document's default table style. The built-in `TableNormal` and `NormalTable` definitions contribute no authored child formatting, matching Word's treatment of those definitions as style bases. Native DOC saving resolves custom default table styles and materializes effective padding without changing the source XML.

Object-only paragraphs retain direct and inherited spacing before and after their flow objects, collapsing that spacing with adjacent paragraphs. Hidden shapes with a resolved flow height reserve that height without drawing the shape. Positioned-shape placement remains subject to the renderer's supported anchor and wrapping mappings.

### Export to bytes or streams

An otherwise empty authored section retains its section-mark paragraph line, including inherited spacing, font size and mark visibility. A section terminator following ordinary text adds no separate line. Conversion preserves the source XML in both cases.

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("invoice.docx");

byte[] pdfBytes = document.ToPdfBytes();

using var stream = File.Create("invoice.pdf");
document.SaveAsPdf(stream);
```

### Optimize images during PDF export

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("input.docx");
byte[] pdf = document.ToPdfBytes(new WordToPdfOptions {
    PdfOptions = new PdfOptions {
        ImageOptimization = new PdfImageOptimizationOptions {
            Enabled = true,
            Mode = OfficeImageOptimizationMode.DownsampleAndRecompress,
            TargetDpi = 144,
            JpegQuality = 85
        }
    }
});
File.WriteAllBytes("output.pdf", pdf);
```

PDF image optimization is opt-in and leaves Word source media unchanged. `Downsample` uses the final source-image placement, including crop/fit expansion; `Recompress` retains pixels and re-encodes JPEGs; `DownsampleAndRecompress` applies both. The shared managed codecs handle static PNG/JPEG/BMP/GIF/TIFF/WebP input. Multi-frame/page payloads and unsupported formats remain outside static optimization. Candidates that would grow or remove metadata are preserved by default; `AllowMetadataLoss` permits reported loss or intentional stripping. Explicit metadata policies also apply when no pixel reduction is needed. The PDF layout report records optimization decisions and warns when metadata is removed. This export policy is separate from the lossless optimizer for existing PDFs.

### Capture conversion warnings without throwing away the report

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("complex-document.docx");
var options = new WordToPdfOptions {
    DefaultTableBorders = true
};

var result = document.SaveAsPdfResult("complex-document.pdf", options);
if (!result.Succeeded) {
    foreach (string diagnostic in result.Diagnostics) {
        Console.WriteLine(diagnostic);
    }
}

foreach (var warning in result.Warnings) {
    Console.WriteLine($"{warning.Source}: {warning.Message}");
}
```

### Import semantic PDF content into Word

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

PdfDocument pdf = PdfDocument.Load("packet.pdf");
PdfWordConversionResult import = pdf.ToWordDocumentResult(
    new PdfToWordOptions());

using WordDocument word = import.Value;
word.Save("packet.docx");

foreach (var warning in import.Report.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

The semantic import path preserves document metadata, page breaks, headings, paragraphs, lists, detected run color/size/emphasis, logical tables, safe URI hyperlinks, supported internal destination links, embedded images, form-widget placeholders, and conversion diagnostics when those structures are available in the PDF logical model. `UseSharedPageReadingOrder` defaults to `true`, so semantic items follow the core engine's crop-, rotation-, spanning-band-, and column-aware order. Editable output also retains each source page's physical size, uses a 36-point page margin, suppresses extra paragraph spacing in dense source content, preserves detected table-column proportions, and positions supported axis-aligned images relative to the page. These defaults improve business-document reconstruction, but they do not turn semantic import into fixed-layout page recreation or Microsoft Word rendering parity.

Set `PreserveSourcePageSize`, `PreserveCompactSourceSpacing`, or `PreserveImagePlacementPosition` to `false` when normal Word flow is more useful than the corresponding source geometry. `EditablePageMarginPoints` controls the source-sized section margin. Source page sizes require `PreservePageBreaks = true` when more than one PDF page is imported.

For selected pages, Fast versus Structured reconstruction, or custom semantic budgets, pass the canonical read settings through the import options:

```csharp
PdfWordConversionReport report = pdf.SaveAsWord(
    "packet-selection.docx",
    new PdfToWordOptions {
        ReadOptions = new PdfReadOptions {
            PageSelection = PdfPageSelection.Parse("1-3,5"),
            LayoutOptions = new PdfTextLayoutOptions { ForceSingleColumn = true },
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 5 }
        }
    }).RequireSuccess().Report!;
```

`ReadOptions` is used only when the source is an opened `PdfDocument`. When a `PdfDocumentReadResult` is already available, the adapter converts that supplied logical model directly.

For table-only recovery, use the same façade with the explicit profile:

```csharp
pdf.SaveAsWord(
    "statement-tables.docx",
    PdfToWordOptions.CreateTablesOnly());
```

## What it exports

- Paragraphs, headings, rich runs, links, bookmarks, page breaks, lists, and common spacing/indentation settings, including hanging and legal negative left/right indents. Built-in Word heading levels 1–9 use the same level mapping as the table of contents in body and column flow. Set `WordToPdfOptions.PdfOptions.TaggedStructureMode` to `PdfTaggedStructureMode.CatalogMarkers` to retain explicit numeric levels during PDF reading and editable Word import, including skipped levels. Tagged PDF uses standard-compatible role mappings for levels 7–9. Untagged output retains the bookmark hierarchy, whose nesting depth cannot identify skipped numeric levels.
- Word-authored text bullets use portable marker characters. Picture bullets currently use a text bullet in PDF output and report `NativePictureBulletTextFallback` with the source picture-bullet identifier; the embedded marker image is not rendered.
- Word sections, page size, orientation, margins, columns, headers, footers, page numbers, and document background color.
- Tables with common Word table styling, repeated headers, cell fills, borders, alignment, merged cells, and rich text in cells.
- Tables in headers and footers retain their cell widths, borders, paragraph flow and supported text direction. The selected default, first-page or even-page story controls body reservation using the authored distance from the page edge. PAGE, SECTIONPAGES and NUMPAGES fields retain their individual number styles and are calculated during pagination, including fields without cached results. Form controls remain static text in these stories.
- Paragraph-aligned images, selected shapes, text boxes, content controls, simple form controls, footnote/endnote markers, and table-of-contents links where supported by the first-party PDF path.
- DrawingML groups of supported preset shapes retain nested child coordinates and scaling. Non-wrapping groups behind text preserve page- or margin-relative positions and paragraph-relative vertical anchors. Paragraph anchors follow pagination, columns, and floating-table clearance; list paragraphs retain both their marker and the group.

- `InFrontOfText` images with explicit page-relative offsets inside the page bounds, including supported source crops, rotation and mirroring. Images follow the first page of their anchor paragraph or heading, including section columns, and paint over text and other flow content without reserving their height in the document flow. Overlapping foreground images follow `WordImage.ZOrder`; images with equal values retain document order.
- Per-operation conversion warnings through `PdfDocumentConversionResult.Report` or `PdfSaveResult.Report`.

Section gutters reserve space at the left, right, or top of the body frame according to the document settings. Mirrored margins swap the left and right body margins on even visible page numbers, including section numbering restarts. Margin-relative shape groups follow that frame; page-relative groups retain their absolute coordinates. A top gutter uses the same horizontal margins on both page sides, matching Word. An explicit `WordToPdfOptions.Margins` replaces the authored margins, gutter, and mirroring.

Header/footer stories containing tables or SECTIONPAGES fields use the shared bounded PDF flow engine. SECTIONPAGES counts physical pages in the section independently of visible numbering and restarts. Stories must fit on the physical page and leave a positive body frame; page breaks and other unbounded content cannot create pages from a header or footer. Exact text metrics, overflow in very short turned cells, floating table positioning and nested table frames remain subject to the native renderer's limits.

Percentage-width tables in the body, headers and footers use the document's compatibility mode. With collapsed borders, legacy layout includes the effective outer cell margins in the percentage reference frame; Word 2013-or-later layout uses the content frame. Direct and inherited cell margins and table indents retain their precedence. Floating and nested table layout remains subject to the limits above.

Simple DrawingML text boxes with an explicit rectangular preset in a running paragraph retain their page anchors when every square-wrap exclusion clears the surrounding text. This path preserves fractional relative sizing, text insets, solid fill and outline, keeps the surrounding text in one paragraph, and sorts boxes within each paragraph by depth, retaining source order at equal depth, without adding their height to the header or footer reservation. Boxes in different paragraphs qualify only when their depth order agrees with paragraph order. Hidden hosting runs do not contribute boxes to placement or approximation diagnostics. Multiple visible boxes in the same header or footer story require explicit overlap permission on every anchor, including boxes in different paragraphs; mutually exclusive variants are assessed separately. Each peer must satisfy the same simple content, style and geometry guards. Inner paragraph borders or shading, multiple text columns, warped text, DrawingML effects, 3D styling, unresolved theme paint and unsupported outline styling retain the explicitly warned mixed-content route. Rich text must fit the padded frame using the shared canvas font and line-spacing measurements. Foreground boxes qualify only when their visible text, fill and stroke clear a conservative body paint envelope, including upward ink from fixed or proportional lines, negative indents and lower-page overflow. Decorated text, including highlighted or underlined spaces, uses its full inner clipping frame for this check. Foreground placement remains approximate when the body contains unassessed drawings, tables, lists or framed paragraphs, or content whose ink cannot be measured, such as note references, fields or decorations. Non-rectangular, custom or missing DrawingML geometry, selected VML boxes, asymmetric effective mirrored margins, clipped text, foreground body overlap, rotated boxes, intersecting wrapping, unsupported peers and nested boxes retain the bounded mixed-content fallback.

Visible shadow, outline, emboss and imprint effects on text also use the warned route. Inherited combinations of these effects remain approximate when their final paint is uncertain; an explicit direct off value disables the corresponding effect.

Mixed text-box flow retains direct text around peer and nested boxes in source order, including equations, note references, supported sibling drawings, and PAGE and SECTIONPAGES results in running stories. Paragraph image limits apply across its segments, and before/after spacing is applied once. Mixed and object-only text boxes using bounded flow report `NativeMixedTextBoxLayoutApproximated` because anchored positioning, shape geometry, wrapping and pagination are approximated. Keep settings apply to the resulting flow blocks; whole-source-paragraph grouping across boxes remains approximate. Nested flow supports up to sixteen text-box levels; deeper content fails with a clear limit error.

For imported groups with unsupported DrawingML geometry, fixed-position export uses the document's VML fallback when available and reports `NativeShapeGroupVmlFallback`. Supported groups with other wrapping or anchor modes are placed in document flow with `NativeShapeGroupFlowed`; groups that cannot be rendered report `NativeShapeGroupUnsupported`. Arbitrary custom geometry, rotation, flips, foreground stacking, and exact text wrapping around groups remain limited.

Next-page section starts create a new page. Odd/even starts use the continuing page number to insert a blank page when needed, then apply the new section's numbering restart. An odd/even start advances a conflicting restart to the next matching number. Next-page starts with an explicit restart also align the section start with that number's parity. Compatible continuous sections share a page. Embedded page breaks preserve text, run formatting, hyperlinks and explicit bookmark targets on both sides of the break, including consecutive blank pages. Fields spanning paragraphs retain their result visibility through page and column splits; hidden field instructions and their breaks do not create pages. A section mark without body content retains its editable formatting and anchors without adding a blank body line or an extra page. Column layout and changes in page geometry remain subject to the native layout engine's supported paths.

Footnote and endnote reference labels preserve decimal, Roman and letter formats, starting numbers, continuous numbering and section restarts. Word's letter-note sequence continues as `aa`, `bb`, `cc` after `z`. A zero DOCX starting number displays `0` for decimal notes and a blank first label for Roman or letter notes, then continues at `1`, `i` or `a` in the selected case. References retain their source position and styling; note-body reference marks use the same label.

Supported single-column footnotes appear on the page containing their reference, including references in split table rows. Ordinary note paragraphs retain run formatting, indentation, tabs, spacing, justification, inline pictures and links. Paragraph shading and border groups use decorated PDF panels; numbered and bulleted paragraphs retain their list markers. Paragraph, list and panel keep constraints participate in page fitting and continuation. Hidden paragraph marks use the same paragraph-joining rules as body text. Notes reserve space before following body text and continue across pages when needed. Explicit section placement selects page-bottom or beneath-text notes. An omitted section position defaults to page-bottom, independently of the document-level position. Explicit page breaks and new-page section boundaries finish pending note continuations before following content starts. Native controls qualify content, reference order, page partition and placement separately; font substitution can change line wrapping and pagination.

Supported footnote and endnote stories can contain flat inline tables between note paragraphs. Cell formatting, horizontal spans and text-only vertical merges use the shared table layout. Rows can continue across pages, with each logical cell resuming its remaining content independently. Vertically merged cells require flexible body rows that allow page breaks. Header styling is retained, but note-table headers do not repeat on continuation pages; ordinary body tables retain their own header-repetition settings. Conversion reads these note tables without adding missing optional properties to the source document.

Footnote separator selection follows the document's registered story identifiers and declared story types. An explicitly registered empty story stays empty. A missing loaded separator uses an implicit paragraph resolved through the document's styles and fonts without changing the source document. Word 2013-or-later compatibility layout reuses the ordinary separator on continued pages; earlier layout uses the continuation-separator story. Registered continuation notices retain their text, typography and spacing, and appear after notes only when note content continues to another page. Earlier compatibility layout also reserves the notice's measured space on complete and final note pages, including a short note that never continues; those pages do not paint the notice.

Separator paragraphs retain their text, typography, spacing and rule marks. An ordinary separator mark fills up to two inches of the remaining line; a continuation mark fills the remaining line. Their bands use the resolved source font's decoration metrics, including a supported single underline. Missing principal rule metrics report `NativeFootnoteSeparatorMarkOmitted`. Unsupported additional decoration patterns or missing underline metrics report `NativeFootnoteSeparatorDecorationOmitted` while retaining the principal rule.

Columns, running-story references, canvas or floating body peers, changing continuation widths, rotated or repeating-header references, nested or floating note tables and other unsupported note profiles retain the section-flow fallback and report `NativeFootnotePlacementApproximated`. Note-table cells currently contain directly represented paragraphs; row and cell block-control wrappers are outside this profile. Non-text content lost through the fallback reports `NativeFootnoteContentOmitted`. Use the conversion report to enforce your loss policy. Wider font and decoration profiles, superscript metrics and exact paragraph geometry remain qualification limits; the native controls do not establish exact layout for every note profile.

Endnotes in ordinary single-column flow retain rich text, paragraph decorations, lists, keep constraints, inline pictures, links and supported flat inline tables. Their reference marks stay among the source run children. The first fragment follows the selected section-end or document-end placement. Later fragments continue in document flow with the distinct continuation-separator story in both supported compatibility profiles. Registered empty stories remain empty, and unregistered notice stories supply no content. A continuation notice is painted only when note content continues to another page. Unsupported continuation profiles report `NativeEndnoteContinuationApproximated`; unsupported endnote content retains its visible text in section flow and reports `NativeEndnoteContentApproximated`.

Without a selected continuation notice, later endnote fragments can use the full body height while their measured separator occupies otherwise unused bottom margin. This allowance stops at the physical page edge and does not enter an accepted page-footnote region or a selected footer. The first fragment and notice-bearing stories retain their separate space reservations. Exact paragraph-end spacing and legacy notice placement remain native-layout qualification limits.

Placement follows the document-wide setting: section-end notes follow their section, and document-end notes follow the final section, including notes referenced in earlier sections. An omitted document-wide position means document-end; section-level endnote positions are ignored as they are in Word. The existing `AddEndnoteProperties(position: ...)` methods update placement for the whole document, preserving section numbering during a position-only edit. `EndnoteSettings.Position` reports the effective document placement. Native DOC import and save preserve that setting. Exact native endnote section transitions, separators and continuation behavior are not qualified for every document profile. Page-based numbering restart reports `NativeNotePageRestartApproximated`. Other numbering formats use decimal labels and report `NativeNoteNumberFormatUnsupported`.

For PDF export, a trailing manual page break honors `CompatibilitySettings.SplitPageBreakAndParagraphMark` in older Word layout. An enabled setting carries the paragraph mark's line height and after spacing onto the next page. Word 2013 compatibility mode ignores the stored setting. Native DOC import selects the legacy enabled behavior; its DOCX projection keeps the same layout. Text following the break continues in its existing paragraph without an extra mark line.

## What it imports

- Parser-supported PDF metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links, complete image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette PDF image streams into editable `.docx` content when their filters are supported.
- Image fallback placeholders and form-widget placeholders with diagnostics instead of silently dropping unsupported objects.
- Page-range filtered imports through `PdfDocument.Read(new PdfReadOptions { PageSelection = ... })`.
- Active hyperlink reconstruction for absolute `http`, `https`, and `mailto` URI annotations through `PdfToWordOptions.ImportUriLinks` and `PdfToWordOptions.AllowedHyperlinkUriSchemes`.
- Internal PDF destination reconstruction through `PdfToWordOptions.ImportInternalLinks`, mapping supported page and named destinations to Word bookmarks and anchor hyperlinks.
- Native image embedding through `PdfToWordOptions.ImportImages`; complete image files, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware soft-mask-capable simple 8-bit `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware soft-mask-capable `Indexed` palette streams are embedded when their filters are supported. Fully transparent and unplaced resources are suppressed. Images whose clips, placement soft masks, unresolved source transparency masks, or unsupported blend modes could expose hidden source pixels are not embedded as raw pictures; the report records a typed omission, and `PdfToWordOptions.IncludeImagePlaceholders` can retain an editable marker instead. Supported opacity is mapped to native Word picture transparency.
- Per-operation import warnings through `PdfWordConversionResult.Report`. `HasLoss` and `RequireNoLoss()` use typed approximation, omission, and failure evidence even when a diagnostic is informational. The table-only profile reports visible text outside detected tables as `PdfTextContentNotImported` instead of treating an intentionally narrow extraction as lossless.

## Options and diagnostics

Word export preserves source-significant leading and repeated spaces in paragraphs
and table cells. When spaces reach the end of a line, their excess is discarded
before the next word rather than creating additional blank lines. Trailing spaces
do not shift centered or right-aligned text.

The converter selects `PdfTextWhitespaceMode.Preserve` unless the supplied
`PdfOptions` explicitly sets a whitespace policy. Set `TextWhitespaceMode` to
`Collapse` to normalize flow spacing, or to `Preformatted` for literal spacing.
The existing `PreserveTextWhitespace = true` setting selects `Preformatted` and
an explicit `false` selects `Collapse`. Options are cloned before export, so the
converter does not change the caller's settings.

For a PDF whose page appearance matters more than editability, use visual pages:

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word.Pdf;

var pdf = PdfDocument.Load("source.pdf");
var options = PdfToWordOptions.CreateVisualPages();
options.Dpi = 144;
options.ReadOptions = new PdfReadOptions {
    PageSelection = PdfPageSelection.Parse("1-3,5")
};
pdf.SaveAsWord("visual-pages.docx", options);
```

Each selected page becomes an image on a Word section with the source page's physical dimensions. Text, links, and form fields are not editable in this mode. Rendering uses the managed PDF engine and reports its limitations. The default limits are 100 pages, 64 million pixels per page, and 256 MB of encoded page images; Word pages larger than 22 inches in either dimension are rejected. Use the default `EditableContent` mode to reconstruct supported text, tables, and images as Word objects.

Use `WordToPdfOptions` when callers need to override page geometry, metadata, page-number behavior, font family, table-border fallback, profile presets, or text fallback policy. `TextFallbacks` uses the shared `PdfTextFallbackFeatures` enum. The balanced resource default enables installed fonts but denies arbitrary local and remote reads; use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion and `CreateTrustedHost()` only when local or remote resource access is intentional. Profiles do not inject page numbers; set `IncludePageNumbers = true` explicitly when generated numbering is desired. Request `ToPdfDocumentResult()` or `SaveAsPdfResult()` when diagnostics matter; unsupported Word features and preserved header/footer overflow become actionable operation results instead of mutable option state. Available embeddable Word families use shared named PDF resources and are not limited to three compatibility slots. Unavailable or non-embeddable families fall back to a mapped PDF font with an explicit warning.

## Current limits

- Non-wrapping foreground and behind-text images retain supported page, margin, single-column and paragraph-relative anchors, including signed offsets and z-order. Page clipping is reported. Line-relative anchors and unsupported wrap and transform combinations use document flow and report `NativeAnchoredImageFlowed`; tight, square and through wrapping remain approximations.
- This package does not try to be a full Word renderer with perfect Microsoft Word parity or a fixed-layout PDF-to-DOCX recreation engine.
- Editable PDF-to-Word import reconstructs parser-supported logical objects. Complex or unsupported PDF image streams, interactive controls, unresolved destinations, and remote or cross-document navigation actions are not reconstructed as native Word objects. Visual pages preserve rendered appearance as images within the managed renderer's support and configured limits. Open reconstruction work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

Use `OfficeIMO.Pdf` for direct PDF layout and manipulation. PowerShell workflows are available through [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice).

## Related packages

- [OfficeIMO.Word](../OfficeIMO.Word/README.md) - Word document model.
- [OfficeIMO.Pdf](../OfficeIMO.Pdf/README.md) - PDF creation, reading, and manipulation engine.
- [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) - HTML/PDF bridge built on OfficeIMO converters.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages; no browser, native renderer, or commercial PDF SDK.
- **OfficeIMO:** `OfficeIMO.Word`, `OfficeIMO.Pdf`, and `OfficeIMO.Core` own the source model, PDF engine, mapping, and reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 1 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
