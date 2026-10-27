# OfficeIMO.Pdf - First-party PDF engine

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

Create reports, read existing PDFs, fill forms, and combine or split pages from .NET. `OfficeIMO.Pdf` runs in your application without Microsoft Office or a browser. It also supports editing, stamps, redaction, optimization, attachments, and PDF standards checks. Optional packages add certificate-based signing and OCR.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Pdf
```

Install the security provider only when the application creates or cryptographically validates certificate-based PDF
signatures:

```powershell
dotnet add package OfficeIMO.Security
```

Install `OfficeIMO.Pdf.Ocr` plus a provider only when the application recognizes scanned pages or creates searchable PDFs. The base PDF package has no OCR runtime dependency.

## Quick start

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .Paragraph(p => p
            .Text("The team completed the monthly review. ")
            .Bold("All scheduled checks passed.")
            .Text(" The table below lists the next steps."))
        .Table(new[] {
            new[] { "Task", "Owner" },
            new[] { "Send the report", "Operations" },
            new[] { "Review next month", "Service team" }
        })), new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 11
        })
    .Meta(title: "Monthly service review", author: "Operations")
    .Save("service-review.pdf");
```

`document.Content.Heading(level, text)` accepts authored heading levels from 1 through 9 in document, page, column and container flow. The `H1`, `H2` and `H3` shortcuts retain their existing presets. Deeper headings accept the same `PdfHeadingStyle` overrides. Enable tagged output with `new PdfCreateOptions().EnableTaggedPdfCatalogMarkers()` when creating the document to retain explicit numeric levels during semantic reading, including skipped levels. Levels 7–9 use `H7`–`H9` tags mapped to the standard `H6` role. Bookmarks preserve the parent-child hierarchy; without tags, semantic reading uses their nesting depth and cannot recover skipped numeric levels.

## Find your next task

- [Read text and tables](#load-once-and-build-one-semantic-read-result).
- [Extract or split pages](#split-and-extract-pages).
- [Combine and organize PDFs](#merge-reorder-delete-duplicate-move-and-rotate).
- [Fill a form or make its values permanent](#fill-and-flatten-a-pdf-form).
- [Review and remove sensitive content](#review-apply-and-verify-redactions).
- [Create an invoice and matching XML](../OfficeIMO.Invoicing.Pdf/README.md).
- [Render HTML as PDF](../OfficeIMO.Html/PDF.md).

For complete workflows, see the [weekly brief](../OfficeIMO.Examples/Showcase/Workflows/WeeklyBrief.cs),
[equipment inspection](../OfficeIMO.Examples/Showcase/Workflows/EquipmentInspection.cs),
and [branded invoice](../OfficeIMO.Examples/Showcase/Workflows/BrandedInvoice.cs).
They create a report with embedded fonts and tags, an editable form with a permanent
record, and matching invoice XML and PDF. The [examples guide](../OfficeIMO.Examples/README.md#run)
shows how to run each workflow.

Keep the original file when editing. Page and form operations return a new document; signed,
protected or unsupported inputs can restrict which changes are allowed. Read the
operation result before publishing the output.

For a batch job or an uploaded file, set an input limit and pass the job's
cancellation token through each operation:

```csharp
using var stop = new CancellationTokenSource(TimeSpan.FromSeconds(30));
using var input = File.OpenRead("report.pdf");
PdfDocument report = PdfDocument.Load(input, new PdfLoadOptions {
    Limits = new PdfReadLimits { MaxInputBytes = 32L * 1024 * 1024 }
}, stop.Token);
PdfDocument selected = report.Pages.Extract(stop.Token, 1, 2);
selected.Save("selected-pages.pdf", cancellationToken: stop.Token);
```

Input streams are read from their current position and remain open. Cancellation
throws `OperationCanceledException`. File output is staged before publication;
when writing to a caller-owned stream, cancellation can leave a written prefix.
Returned documents carry the selected read policy into subsequent operations;
an explicit per-operation override applies to the result without changing the source.
Filling or flattening forms first checks existing visible appearances against these input limits.
The output budget then allows for the appearance content generated by the operation.
Splitting to several files prepares every part before saving. Cancellation stops
that preparation; once saving begins, the prepared files are published one at a time.

## Output size and images

Page content is Flate-compressed by default, including direct `PdfDocument.Create`
authoring. `PdfCreateOptions.CompressContentStreams = false` keeps page operators plain.
ICC profiles, embedded files and widget appearances are compressed independently;
XMP packets remain uncompressed. Identical graphics states, shadings, functions and
untagged Form resources are shared within the document. Byte-identical page programs
also share storage while retaining each page occurrence. Output keeps classic cross-reference
tables; compact object/xref streams remain available through `Optimization`.

Supported PNG/JPEG ICC profiles are retained as ICCBased image resources. Indexed PNGs
keep their palette and transparency mask. Supported SVG input to `Image(...)` and
`TryPrepareImageBytes(...)` is rendered through the shared managed drawing engine
and normalized to PNG within the image transcode pixel limit. Unsupported SVG
features are rejected rather than silently omitted.

EXIF rotation and mirroring use placement matrices, and layout uses the oriented
dimensions. JPEG DCT payloads, PNG sample depth, profiles, transparency masks and
repeated-image sharing survive orientation changes. Sixteen-bit PNG colour and alpha
samples are retained with PDF 1.5+ output; PDF 1.4 rounds them to eight bits. Visible
signature images use the same image-resource allocator and target-version rules.
Managed display and PDF-to-drawing replay apply the PDF placement exactly once;
image extraction retains the original JPEG bytes.

Attach a `PdfConversionReport` through `ReportDiagnosticsTo(...)` to inspect authored
overflow and clipping in rows, fixed boxes, paragraph frames, rotated and viewport
cells, free-text annotations, and running content. Paint checks include glyph ink,
decorations and inline elements. Standard-font checks use Adobe glyph bounds;
embedded-font checks use the selected outlines. Reports distinguish preserved overflow from clipped
paint. Bounded inspection reports unassessed geometry when its limits are exceeded
or nominal standard-font bounds cannot prove loss through a partial path clip. These checks
do not change layout or text emission.

## Header and footer content

Headers and footers can contain the same bounded paragraph, table, image and drawing flow used by the document body. `Content`, `FirstPageContent` and `EvenPagesContent` select a story for each page and reserve body space from its measured height. Distances and the optional `bodyGap` are in points from the top or bottom page edge. Repeated image payloads are shared across pages and pagination passes. During layout, including pagination stabilization, running content retains at most 128 MiB of distinct encoded image bytes, prepared streams and image masks; exceeding this limit throws `InvalidDataException`. This bounds retained layout assets; transient callback/decoder allocations and opt-in image optimization or PDF/X conversion during serialization use their own policies.

```csharp
var document = PdfDocument.Create(document => document.Page(page => {
    page.Header(header => header.Content(content => content
        .Table(new[] { new[] { "Project", "Status" }, new[] { "Migration", "Active" } }),
        distanceFromEdge: 18, bodyGap: 6));
    page.Footer(footer => footer.Content(context => content => content
        .Paragraph(paragraph => paragraph.Text($"Page {context.PageNumber} of {context.DocumentPages}")),
        distanceFromEdge: 18));
    page.Content(content => content.Paragraph(paragraph => paragraph.Text("Report content")));
}));
document.Save("report.pdf");
```

The page-context overload supplies visible and physical page numbers, section-relative page position, section and document page counts, and the current content width. `SectionPages` counts physical pages in the current page group independently of numbering restarts. `TotalPages` retains the existing `{pages}` meaning: the last visible number in the current numbering sequence. `FormatPageNumber` follows the configured number style or an explicit override. The factory must be deterministic because page totals can require pagination to stabilize. Running content cannot create pages: page breaks, page canvases, deferred content and other unbounded flow are rejected. The measured stories must leave positive space for the document body.

A kept block or indivisible fragment can advance to a later first/even/default running-story frame that has room for it. Each actual frame is measured through normal page and column transitions; future factories are not evaluated eagerly. Static frame cycles end with the normal fitting error. Searching arbitrary page-context factories requires a finite `MaxGeneratedPages` budget, and cancellation and page limits continue to apply.

Running stories are marked as pagination artifacts. Their fonts, images, drawing resources and links use the document's normal PDF resource owners. A later `Text` or `Zones` call replaces rich content for that variant.

## Page-associated footnotes

Define a note with the ordinary flow builder, then associate a styled reference run with that definition:

```csharp
using OfficeIMO.Pdf;

var document = PdfDocument.Create();
PdfFootnoteDefinition note = document.Content.DefineFootnote(
    body => body.Text("1. Additional information about this paragraph."));
document.Content.Paragraph(paragraph => paragraph
    .Text("The report includes a page note")
    .Runs(new[] { PdfTextRun.Superscript("1").WithFootnoteReference(note) }));
document.Save("report.pdf");
```

The reference and the first legal note fragment stay on the same page. Notes reserve space before later body content, can continue across pages, and retain rich paragraph text, inline images and links. Ordinary flow lists and decorated panels can also appear in note definitions. Their markers, borders, backgrounds and fragment boundaries pass through the shared note layout. Paragraph, list and panel keep constraints participate in choosing legal fragments; the fitting and continuation limits below still apply. References in split table rows attach to the page containing the actual reference fragment. Definitions belong to their document; each serialization has independent pagination state.

Eager inline tables can also appear in note definitions. Flat grids support horizontal spans, text-only vertical merges and rows that continue across pages. Each cell retains its own continuation position, so a shorter neighboring cell is not replayed while a longer cell continues. Vertically merged cells require flexible rows that allow page breaks and stay within the table's body-row region. Table headers retain the configured header-repetition policy.

`PdfFootnoteOptions.Placement` defaults to `BottomOfPage`, above any running footer. `BeneathText` places the note below the page's body text. The optional `Separator` and `ContinuationSeparator` callbacks capture ordinary flow content before the first note and continued notes respectively; the engine creates no separator when they are omitted. Visible numbering is authored text, so the reference and note can use the labels and formatting your application needs.

`ContinuationNotice` captures flow to place after the page's notes when note content continues to another page. Its measured height is reserved before accepting note fragments; the notice is painted only when content continues. By default, a complete or final note page reserves no notice space. Set `ReserveContinuationNoticeSpaceOnFinalPage = true` to retain that measured space even for a short note that fits on one page, while keeping the notice hidden there.

The first note admitted to a page selects that page's separator, notice and final-page reservation policy, including when table references are admitted in a different order from their source. Definitions capture the options and callback content when composed. Later changes to those options or retained builders do not change the definition.

Explicit page breaks and new-page section boundaries finish pending note continuations before starting the following content. Automatic pagination can resume body content on the page containing the final note fragment.

Page notes require single-column body flow and statically authored references. `TableDeferred` batches and replayable `Deferred` callbacks cannot introduce note references; use ordinary tables or static `Flow` content for those references. Deferred content without note references can share pages with notes when it follows the same body-layout constraints.

Multi-column or row composition, page canvases, references in running stories, floating or non-flow tables, rotated reference cells, repeating table-header references, recursive notes and unsupported note-flow constraints fail explicitly. An active note continuation cannot cross a change in usable page width. A reference and first note fragment that cannot fit in a fresh page body fail with a fitting error; `MaxGeneratedPages` also bounds continuation output.

Note tables do not support floating placement, captions, rotated cells, recursive references, form fields or viewports. Vertically merged cells do not support inline images, cell data bars or icons, fixed row heights, or rows that prohibit page breaks. Separator and continuation-notice callbacks cannot contain tables.

## Inline rules

Use `PdfInlineRule` for a horizontal band that fills the space remaining on a rich-text line:

```csharp
document.Content.Paragraph(paragraph => paragraph
    .Text("Signature: ")
    .Inline(new PdfInlineRule(height: 0.5, color: PdfColor.Black,
        baselineOffset: -2)));
```

Height, `MaximumWidth` and `BaselineOffset` are in points. The baseline offset positions the bottom of the band relative to the text baseline; a negative value places it below that baseline. A null `MaximumWidth` fills the remaining line, while a positive value caps the width. Each layout resolves the width against its actual paragraph, table-cell or nested frame. The rule adds no intrinsic width during automatic table sizing, and text following a rule wraps normally if the rule consumes the rest of the line. A frame with no remaining width paints no band.

Rules can also be composed inside footnote separator callbacks. For example, `new PdfInlineRule(0.5, PdfColor.Black, maximumWidth: 144)` draws a band up to two inches wide without extending beyond its line frame.

## Fonts and embedding

Text set in a standard font slot (`Helvetica`, `Times` or `Courier`) uses an installed font from the default theme's fallback lists and embeds it as a subset: the body font chain (Aptos, Calibri, Carlito, Helvetica Neue, Arial, Liberation Sans) for `Helvetica`, Times-compatible faces (Times New Roman, Liberation Serif, Tinos) for `Times` and the monospace chain (Consolas, Menlo, Courier New, Liberation Mono) for `Courier`. Office converters back a slot they leave unregistered with a metric-compatible face first (Arial, Liberation Sans; Courier New, Liberation Mono), because such a slot stands for a source family like Arial. A family with real bold and italic faces is preferred, so emphasis survives the fallback. Fonts you register yourself (`UseFontFamily`, `RegisterFontFamily`, `UseOfficeFontFamily`) are never replaced, and sections that use the same face share one subset.

Set `PdfCreateOptions.EmbedDefaultFonts = false` to write the non-embedded standard PDF fonts instead. When no font file from the fallback list is installed, the document still saves with the standard font and reports a `font-not-embedded` warning without loss, because only the appearance depends on the viewer. Slots reserved by a text fallback set stay free for the fallback fonts. Office converters follow their `ResourcePolicy`: when installed fonts are not allowed, they write the standard fonts.

## Literal text to PDF

`PdfPlainTextConverter` renders TXT content without interpreting Markdown or HTML:

```csharp
using OfficeIMO.Pdf;

using OfficeIMO;

using var source = File.OpenRead("report.txt");
using var destination = File.Create("report.pdf");
OfficeConversionReport report = OfficeConverter.Convert(source, OfficeDocumentFormat.Text,
    destination, OfficeDocumentFormat.Pdf,
    new OfficeConversionOptions {
        PlainText = new PdfTextImportOptions { TabSize = 4 },
        MaximumInputBytes = 64L * 1024 * 1024
    });
```

The converter detects UTF-8, UTF-16 and UTF-32 byte order marks and otherwise decodes strict UTF-8. `EncodingName` selects an explicit encoding available in the host; a conflicting BOM or invalid byte sequence fails conversion. Seekable input position is restored and the caller's stream remains open.

Lines, blank lines and repeated spaces are preserved. Tabs advance to source-column stops, form feeds start pages, long lines wrap, and empty input produces one blank page. Defaults are ten-point Courier, 16 million decoded/expanded characters and 10,000 generated pages. `PdfTextImportOptions.PdfOptions` controls fonts and page geometry. Supply embedded font bytes for characters outside the selected font's coverage; unsupported glyphs fail output rather than publishing replacement text.

## Paragraph font size and line spacing

`PdfCreateOptions.TextWhitespaceMode` controls ordinary spaces in flow text.
`Collapse`, the default, keeps one separator between words and discards leading
spaces. `Preserve` retains leading and repeated spaces within a line, counts each
space during justification, and discards excess spacing at an automatic wrap.
`Preformatted` retains literal spacing even when it creates additional blank
lines. Tabs and explicit line breaks keep their separate behavior.

```csharp
var options = new PdfCreateOptions { TextWhitespaceMode = PdfTextWhitespaceMode.Preserve };
var document = PdfDocument.Create(options);
document.Content.Paragraph(p => p.Text("  First   second"));
document.Save("preserved-spacing.pdf");
```

The existing `PreserveTextWhitespace` property remains available. Setting it to
`true` selects `Preformatted`; setting it to `false` selects `Collapse`.

Paragraphs can use a fallback font size independently of the document default:

```csharp
var document = PdfDocument.Create(new PdfCreateOptions { DefaultFontSize = 12 });
document.Content.Paragraph(p => p.Text("First line\nSecond line"), style: new PdfParagraphStyle {
    FontSize = 8,
    LineSpacing = PdfLineSpacing.Exactly(14),
    SpacingAfter = 0
});
document.Save("fixed-spacing.pdf");
```

`PdfLineSpacing.Multiple(1.25)` scales the advance with each line's font size.
`Exactly(14)` keeps a fourteen-point advance even with larger runs or inline
elements; their painted content can extend beyond the line box. `AtLeast(20,
naturalMultiplier: 1.2)` uses twenty points for small text and expands when the
natural text or inline-element height exceeds that minimum. The rule and value
are readable from the immutable spacing object.

`LineSpacing` overrides the existing `LineHeight` multiplier. Heading and list
styles accept the same spacing object. Font size and spacing are snapshotted with
their style and apply consistently to paragraph measurement, pagination and
rendering in flow, columns and nested frames.

Each wrapped line places its baseline using the fonts and sizes it contains.
Larger text and lowered inline elements reserve their combined ascent and descent,
including inside columns, table cells and canvas text boxes.

Wrapping paragraphs retain an identifier's delimiter-separated groups when the
measured group fits the available line. An oversized group uses ordinary long-token
splitting instead of a soft overflow allowance. Wrapping table cells use the same
width admission, including paragraphs inside nested tables.

`PdfParagraphBuilder.LineBreak()` retains the current run style. A larger font
on a blank line can expand proportional or minimum spacing; exact spacing keeps
its fixed advance.

## Paragraph justification

`PdfAlign.Justify` uses `PdfParagraphStyle.JustificationMode` to select expandable boundaries. `Words`, the default, distributes extra width between word separators. `Characters` distributes it uniformly between shaped grapheme clusters, keeping combining sequences and shaped ligatures together. `JustifyExplicitLineBreaks` and `JustifyLastLine` independently select expansion for explicit-break and final lines; both default to `false`.

```csharp
var style = new PdfParagraphStyle {
    JustificationMode = PdfJustificationMode.Characters,
    JustifyExplicitLineBreaks = true,
    JustifyLastLine = true
};
var document = PdfDocument.Create();
document.Compose(layout => layout.Content(content => content
    .Paragraph(paragraph => paragraph
        .Text("Distributed first line").LineBreak().Text("Final line"),
        PdfAlign.Justify, style: style)));
document.Save("distributed-text.pdf");
```

Paragraph fragments retain the source line policy across page, column and table-cell boundaries. Page-note paragraphs retain it across their supported continuation pages. A tab fixes its field anchor in both modes: earlier fields retain their advances, and only word separators after the last tab expand. Character distribution adds positive page-point spacing after layout; it does not compress ordinary spaces, admit additional words to a line or reproduce a native application's nonuniform character-expansion weights. Available fonts and the selected shaping provider still determine the original glyph metrics.

## Paragraph tab stops

Explicit tab positions are finite point coordinates relative to the paragraph's continuation text frame. Zero is that frame's origin; negative positions can be reached from a hanging first line. A line consumes only stops ahead of its current cursor and skips earlier stops after a line break.

```csharp
var style = new PdfParagraphStyle { LeftIndent = 36, FirstLineIndent = -36 };
style.AddTabStop(-18);
PdfDocument.Create()
    .Compose(layout => layout.Content(content => content
        .Paragraph(paragraph => paragraph.Text("1.\tFirst field").LineBreak()
            .Text("Continuation\tLater field"), style: style)))
    .Save("hanging-tabs.pdf");
```

The negative stop is eighteen points before the continuation text origin. Following lines retain that origin and their default tab interval. NaN and infinite positions are rejected.

Bar stops draw stationary vertical rules alongside the paragraph's lines:

```csharp
var style = new PdfParagraphStyle();
PdfTabStop rule = style.AddTabStop(72, PdfTabAlignment.Bar);
rule.BarWidth = 0.5;
rule.BarColor = PdfColor.Black;
```

Each rule spans its line's resolved height, including applied paragraph space before on the first line and space after on the last. Bars keep their paragraph origin through hanging indents, alignment and continuation. They leave glyph positions and positional tab selection unchanged. Paragraph and heading styles capture the rule's position, width and color; bar alignment is available on style declarations, while `PdfTextRun.Tab` accepts text alignments.

## Decimal tab alignment

An explicit decimal tab can select its decimal separator independently of grouping punctuation:

```csharp
var style = new PdfParagraphStyle();
style.AddTabStop(180, PdfTabAlignment.DecimalSeparator).DecimalSeparator = '.';
using var pdf = PdfDocument.Create(new PdfCreateOptions { TextWhitespaceMode = PdfTextWhitespaceMode.Preserve });
pdf.Content.Paragraph(p => p.Text("Total\t1,234.56"), style: style);
```

This aligns the period at 180 points from the paragraph frame. A value without the selected separator is right-aligned at that stop. Leaving `DecimalSeparator` unset retains the first-period-or-comma behavior. The setting survives paragraph-style copies and table layout.

## Glyph width and character spacing

Run width and tracking adjust horizontal advances without changing font height:

```csharp
var condensed = PdfTextRun.Normal("Condensed text", fontSize: 12)
    .WithHorizontalTextScaling(75)
    .WithCharacterSpacing(0.5);

using var pdf = PdfDocument.Create();
pdf.Content.Paragraph(p => p.Runs(new[] { condensed, PdfTextRun.Normal(" Natural text") }));
pdf.Save("run-spacing.pdf");
```

`HorizontalTextScaling` is a positive percentage; 100 retains the font's normal
width. `CharacterSpacing` adds page points after each rendered glyph, independently
of that percentage. Negative spacing condenses advances. The copy methods retain
the source run's other formatting. Paragraph builders expose matching
`HorizontalTextScaling(...)` and `CharacterSpacing(...)` controls; use 100 and zero
to restore natural text for following runs. Measurement, wrapping and painting use
these metrics in rich paragraphs, table cells, page text and positioned text.

## Authoring model

`PdfDocumentBuilder` owns document settings and page boundaries.
`PdfContentBuilder` is the single flow receiver used by document content, page
content, components, logical groups, and row columns. This keeps reusable content
independent of where it is placed.

Composition defaults come from `PdfCreateOptions.Theme`, an `OfficeTheme` that
defaults to `OfficeTheme.Default`. Headings use the theme heading family, type
scale, colour and spacing; tables without a style use the theme table style with a
header rule, hairline rows, columns fitted to their content and right-aligned
numbers; lists, panels and footer text follow the theme as well, and page numbers
read `Page N of M`. Explicit default styles and block styles take precedence.
`Title` and `Subtitle` sit above the heading hierarchy, and a title stays out of
the generated outline:

```csharp
var document = PdfDocument.Create(new PdfCreateOptions { Theme = OfficeTheme.Modern, ShowPageNumbers = true });
document.Content
    .Title("Community workshop handbook")
    .Subtitle("A practical reference for hosts and facilitators.")
    .H1("Before participants arrive")
    .Table(new[] { new[] { "Segment", "Minutes" }, new[] { "Welcome", "10" } });
```

Use `PdfTheme.FromOfficeTheme(theme)` with `ApplyTheme` to copy those styles into
the options and adjust them.

For incremental authoring, use `var document = PdfDocument.Create(options)` and
add blocks through `document.Content`. The same builder supports both ordinary
statements and fluent chains; callbacks remain useful for grouping and page setup.

Image-file pages and image stamps accept at most 128 MiB of encoded data per
image by default. Increase the limit for a trusted source when needed:

```csharp
using var images = new[] { "scan-1.png", "scan-2.png" }.ToPdf(
    new PdfImageImportOptions { MaximumEncodedImageBytes = 256L * 1024 * 1024 });

using var imageStream = System.IO.File.OpenRead("stamp.png");
var stamped = PdfDocument.Load("input.pdf").Stamp.Image(
    imageStream,
    new PdfImageStampOptions { MaximumEncodedImageBytes = 256L * 1024 * 1024 });
```

`PdfImageDocumentSource.FromFile(path, maximumEncodedImageBytes)` applies the
file-read budget when creating a source directly. If that source is passed to
`sources.ToPdf(options)`, set
`PdfImageImportOptions.MaximumEncodedImageBytes` to the intended limit too;
document creation checks every source independently. For edits to existing PDF
images, set `PdfImageEditOptions.MaximumEncodedImageBytes` when adding, replacing,
moving, or transforming a larger trusted image. Stamp streams are read from
their current position. The byte budget applies before image decoding; decoded
image and PDF output limits remain separate.

```csharp
PdfDocument.Create(document => document
    .Settings(options => {
        options.PageSize = PdfPageSizes.A4;
        options.Margins = PdfPageMargins.Uniform(36);
    })
    .Content(content => content
        .H1("Work items")
        .Row(row => row
            .FixedColumn(48, cell => cell.Text("ID"))
            .AutoColumn(cell => cell.Text("Owner"), maximum: 100)
            .RelativeColumn(cell => cell.Text("Description")))));
```

`Element(...)` is the uniform layout and semantics envelope. Its `Content(...)`
callback receives the same `PdfContentBuilder`, so a reusable component or flow
fragment does not need an element-specific API:

```csharp
PdfDocument.Create(document => document
    .Typography(OfficeRenderingProfile.Managed)
    .Settings(options => options.TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers)
    .Content(content => content.Element(element => element
        .Semantic(PdfSemanticRole.Article)
        .Background(PdfColor.FromRgb(248, 250, 252))
        .Border(PdfColor.FromRgb(148, 163, 184))
        .Padding(vertical: 10, horizontal: 12)
        .MaxWidth(420)
        .Align(PdfAlign.Center)
        .KeepTogether()
        .Content(article => article
            .H1("Accessible report")
            .Text("One composition path from API to tagged output.")
            .Row(row => row
                .FixedColumn(64, cell => cell.Text("ID"))
                .RelativeColumn(cell => cell.Text("Description")))))));
```

Decorated elements use a scoped flow frame rather than a second layout engine.
Rows, sections, multi-column flow, layers, deferred content, nested elements, and
explicit page breaks therefore retain their normal pagination behavior inside an
element, while its background and borders continue across page fragments.
`KeepTogether` and `KeepWithNext` are strict promises: content that depends on a
dynamic page callback, automatic multi-column pagination, a deferred table, a
generated table of contents, a canvas, or an explicit page boundary cannot be
preflighted and is rejected when such a promise requires its complete height.
Move that content outside the constrained element or flow instead of relying on
best-effort pagination.

Add `using OfficeIMO.Drawing;` when applying an `OfficeRenderingProfile`.
`Managed` selects the shared dependency-light provider for its
[horizontal static-font Arabic/Hebrew/Latin profile](../OfficeIMO.Core/README.md#managed-multilingual-shaping),
including supported mark attachment and required Arabic ligatures. A declined run
reports approximation loss, so `RequireNoLoss()` rejects its fallback. Supply a
profile containing your embedded font family, fallbacks, language and shaping
provider when the document requires broader scripts or reproducible font selection.

Drawing text preserves numeric font descriptors when the matching faces are registered in
`PdfCreateOptions.UseRenderingProfile(...)`. Measurement and embedded PDF text use the same selected
font program. Foreground text and inherited underline/strikethrough paint use `OfficeColor.A`;
an explicit decoration color uses its own alpha. Foreground transparency preserves logical text,
placement and run backgrounds. Fully transparent glyphs remain searchable in the PDF.

Drawing strokes support native linear and radial gradient shading through their
shared outlines, including caps, joins, dashes, opacity, clipping, and affine transforms.
Line and open-path end markers retain their stroke brush, opacity, transform and
clip. Triangle, stealth, diamond and oval ends paint filled geometry; open-arrow
ends paint an outline.
Running headers and footers report `HeaderFooterPageBoundsClipped` when shape or
marker paint extends beyond the physical page after applying the actual clip.
`HeaderFooterPaintBoundsUnmeasured` records unassessed page clipping when contour
inspection exceeds its work limit.
Linear-light RGB gradients use calibrated PDF RGB shading, preserving the color
field without adding sampled color stops. Explicit print-condition conversion
uses the existing bounded CMYK sampling path with the gradient's interpolation
mode. Alpha masks remain scalar opacity fields.
Radial gradients with an explicit Repeat or Reflect mode use vector function-based
shading, including fields with no finite cycle bound. Color and alpha retain the
same periodic field, within the existing 1,024-stop PDF gradient bound. Explicit
print-condition conversion uses the shared bounded ICC gradient sampler, with up
to 4,096 CMYK samples and independent scalar alpha. Component functions retain
the periodic geometry without rasterizing the PDF field.

The reader renders directly invoked function-based shadings at the requested
image-export resolution. `ToDrawing()` samples these fields at one pixel per PDF
point; use image-export or page-render options for a higher resolution.
`PdfReadLimits.MaxFunctionShadingPixels` bounds aggregate intermediate pixels,
and `MaxFunctionShadingEvaluationWork` bounds calculator work per page.
Function-based shading patterns remain unsupported and produce a render diagnostic.

Gradient fills and strokes preserve color-stop alpha through native transparency masks,
including header and footer shapes. Gradient direction follows the same local coordinates
before and after an affine transform.
Path and polygon fills preserve `OfficeShape.FillRule`, including even-odd holes,
gradient clipping, affine transforms, and header and footer shapes.

Embedded SVG drawing images retain supported vector paths, gradients, clipping and text.
PDF tiling patterns store a reusable cell instead of a copy of every tile. Repeated link
annotations share a generation-wide expansion budget derived from the patterns' existing
`maximumTileCount` limits; nested patterns and running content consume the same budget.
Paint that the SVG reader can import but the PDF vector writer cannot express uses raster fallback and
reports `SvgImageRasterFallback`; incomplete SVG imports fail explicitly.
Canvas gradients interpolate premultiplied channels, including linear-light RGB gradients.
Imported SVG/XPS retains separate colour and alpha interpolation.

Drawing export retains bounded vector tiling patterns, isolated group opacity and
blend modes, and alpha or explicit `PdfDeviceRgb` luminosity soft masks with
transparent backdrops. Mask values remain separate from print color conversion.
For colored luminosity masks, the PDF consumer chooses the DeviceRGB-to-gray conversion.
The default sRGB luminosity standard and nontransparent mask backdrops reject export.
Embedded drawing images preserve their
interpolation setting. When reopening PDFs, the managed reader retains ordinary
isolated Form groups with implicit or DeviceRGB blending space and no knockout.
Invocation opacity applies once after overlapping paths, images, text, and nested
forms are composed. Logical text extraction remains separate from that grouped
paint. Other group blending spaces, non-isolated groups, and knockout semantics
remain outside this reader contract. The [Cairo group fixtures](../OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/Transparency/SOURCE.md)
provide independent-producer coverage for overlapping child alpha and nested groups.

HTML and SVG adapters preserve positioned text without turning its line or
advance measurements into an extra paint clip. Glyphs can extend beyond those
layout boxes; authored CSS overflow and SVG clipping still constrain them.
Explicit drawing text frames and public canvas text boxes retain their clipping
behavior, including clipping link annotations to the visible frame.

`OfficeDrawing.AddVerticalText(...)` draws native positioned glyphs in PDF when
the selected shaping provider supplies vertical advances and complete logical
text coverage, the chosen font is embedded, and an `/ActualText` wrapper owns
extraction. The writer preserves shaped glyph substitutions and X/Y offsets,
marks the original text for extraction, and clips paint to the drawing text
box. Without that font/provider contract it retains searchable stacked text and
reports `vertical-text-stacked-fallback`; `PdfConversionReport.RequireNoLoss()`
rejects that approximation. The [vertical PDF reference gate](../OfficeIMO.Pdf.Tests/Pdf/Fixtures/Vertical/SOURCE.md)
records the tested CJK scope and independent render/extraction checks.

For a complete generated sample, see
[`Pdf.AuthoringModel.cs`](../OfficeIMO.Examples/Pdf/Pdf.AuthoringModel.cs). The
existing report, invoice, label-sheet, ticket, table-style, drawing, dashboard,
and statement examples build on the same public content model.

Rows make width intent explicit: `FixedColumn` uses points, `AutoColumn` measures
content within optional bounds, `RelativeColumn` divides remaining space by
weight, and `PercentColumn` consumes a percentage of the available row width.
Use `Column(PdfColumnWidth, ...)` when the sizing value is created separately.
Mixed rows reject a relative column when fixed, automatic, and percentage columns
leave it no positive width. Automatic sizing measures panel text and padding as
well as ordinary text and fixed-size primitives.
Columns can contain the normal flow primitives, including rich text, lists,
tables, images, drawings, form fields, annotations, and `Panel(...)` groups.
Balanced automatic columns measure nested panels at their child widths and retain
unfinished semantic and static flow groups with their following siblings.
`PdfPanelStyle.LeftInset` and `RightInset` place the whole panel inside its
containing frame, using `OfficeLength` values. The remaining width controls
wrapping, maximum width and alignment on every page or column fragment.
Negative insets extend the frame outward. Padding remains inside that frame.
`PdfPanelStyle.RepeatFragmentDecoration` defaults to `true`: each page or column
fragment repeats its vertical padding and top/bottom borders. Set it to `false`
for a continuous panel whose top belongs to its first fragment and whose bottom
belongs to its final fragment. Side borders and background continue across the
fragments. Lists and table rows retain their legal breaks, including cell keep
and widow rules within the padded content height.

`FragmentBottomInset` reserves space at the bottom of each fragment without
moving the next fragment's first line. Closing padding can occupy that space;
the default inset is zero. Flow images with `ScaleDownToFit` account for the
panel's padding and inset when choosing their size.
Continued images use the space available after the first fragment's top padding.
Final content retains the full closing padding; intermediate fragments keep
their normal capacity. Content that cannot leave closing space is rejected.

Use a side border's `PdfPanelBorder.Offset` to move its stroke independently of
the content frame. Positive point values move it outward and negative values
move it inward. Offsets preserve text wrapping and padding:

```csharp
using var document = PdfDocument.Create();
document.Content.Panel(panel => panel.Paragraph(p => p.Text("Content in a continuous panel.")),
        new PdfPanelStyle {
            RepeatFragmentDecoration = false,
            PaddingY = 4,
            LeftBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = 2 },
            RightBorder = new PdfPanelBorder { Color = PdfColor.Black, Offset = 2 }
        });
```
When columns start partway down a page, kept paragraphs, lists, panels and static
flow groups can move past the partial columns to a full physical page. Headings
and kept blocks retain their following siblings for final-page balancing.
`BalanceKeptParagraphLines` allows a kept paragraph longer than a balanced column
to span columns on one physical page. A kept paragraph that fits the chosen
column stays whole. This option defaults to `false`; ordinary sequential frames
retain their keep-together behavior.
`HonorKeepWithNextWhenBalancing` defaults to `true` and keeps a breakable
paragraph or list's final line group with the following block while balancing.
Set it to `false` for source formats whose older layout modes ignore this rule
in balanced columns. Partial physical pages and ordinary full-height frames
retain their normal keep-with-next rules.
Decorated elements, semantic groups, static components, and static flow with
position capture or keep-together rules work inside rows. Page boundaries,
sections, layers, automatic multi-column layouts, contextual or conditional flow,
and nested rows remain outer layout boundaries.

Percentages are literal in every row. Use relative weights for proportional fill.
`Panel` and `Element` preserve child blocks through the same container layout;
`PanelParagraph` adds one ordinary paragraph inside that container. Keep-together
content must fit a complete frame, including padding. Otherwise allow splitting.

## What it does

- Creates PDFs with page setup, headings, paragraphs, rich text, links, lists, reusable typed and page-aware components, tested report/invoice/label-sheet/ticket recipes, mixed inline images and boxes, dictionary-driven hyphenation, styled multipage containers, balanced block-flow columns, conditional/replayable flow, position capture, sections, generated TOCs, optional-content layers, tables, images, vector drawing, headers, footers, watermarks, metadata, portfolios, and form primitives. Raster inputs accepted by `OfficeIMO.Drawing` normalize once through the shared image owner before PDF embedding. Eight/sixteen-bit Huffman lossless JPEG inputs normalize to PNG before embedding; supported one/three-component eight-bit DCT JPEGs retain their compressed payload. Lossless JPEGs with embedded ICC profiles are rejected when normalization cannot retain their color contract.
- Reads and inspects PDFs through text extraction, logical document objects, page metadata, links, images, attachments, portfolios, outlines, forms, bounded immutable raw-structure views, active-content diagnostics, and security/revision markers.
- Manipulates existing PDFs with page extraction, split, merge, delete, duplicate, move, rotate, metadata editing, stamps, watermarks, and complete-page overlay/underlay while preserving source PDF header versions on shared rewrite paths.
- Renders supported embedded TrueType and OpenType/CFF fonts with stable-glyph subsetting. Generated text applies pair kerning by default, with automatic bounded bidirectional ordering and the shared managed directional-shaping route. Supported static-font Arabic joining, required ligatures and mark positioning use the [managed multilingual profile](../OfficeIMO.Core/README.md#managed-multilingual-shaping); declined complex operations report approximation loss. `UseManagedTextShaping()` selects the shared positioned-glyph provider explicitly. The shared `IOfficeTextShapingProvider` contract remains the extension point for Indic, Thai, and other complex scripts and for full shaping engines.
- Renders embedded TrueType, CFF, Type1C, CIDFontType0C, OpenType/CFF, and Type 1 fonts from their own outlines when reading PDFs, with PDF viewer stroke defaults, one-device-pixel hairlines, filtered image downscaling, anti-aliased clipping, function-based and mesh shadings, and JBIG2 images. Arithmetic JBIG2 dictionaries retain and reuse generic and refinement bitmap contexts, including global dictionaries; direct regions follow the page combination policy and page refinements preserve their replacement operator. The [state fixtures](../OfficeIMO.Pdf.Tests/Pdf/Fixtures/Rendering/jbig2-state.json) record independently decoded pixels. Type 1 and CFF outline projection supports positive uniform font-matrix scales; rotated, sheared, translated, reflected, or unequal-scale font matrices use a mapped substitute with `render.resource.font-cff-substitution`. Unsupported images and knockout groups are reported as `render.resource.image-unsupported` and `render.resource.knockout-group-simplified`.
- Projects authored annotation appearance streams into page images. When a supported free-text, text-markup, shape, line, ink, path, stamp, or caret annotation has no usable normal appearance, the renderer reuses the bounded annotation synthesizer and reports `render.annotation.appearance-synthesized` as an approximation.
- Shares managed CMYK, Lab, XYZ, calibrated-color conversion, bounded sampled, exponential, stitching, and Type 4 calculator color functions, vector tiling fills, standard blend modes, and alpha/luminosity soft masks with `OfficeIMO.Drawing`. Catalog destination output profiles with supported RGB matrix/TRC or ICC mBA transforms soft-proof vector, text, form, pattern, and image colors through the same rendering-intent pipeline. ICC LUT-composed and output-profile-composed shadings remain fail-closed unless their final interpolation can be certified. Pages with explicit transparency retain authored colors and report `render.colorspace.icc-output-intent-transparency-simplified` until output conversion can run after composition. Color-managed DCT/JPEG images use the ICC, `/Decode`, Indexed-palette, and transparency pipeline, while simple DeviceGray/DeviceRGB JPEGs without required color management remain lossless pass-through payloads. DeviceCMYK JPEGs normalize through the PDF image pipeline so that Adobe markers cannot override authored or implicit identity `/Decode` polarity.
- Bounds completed page/effect content and serialized-object retention with separate memory limits, temporary-file spillover, direct large-stream spooling, and chunked final assembly during stream saves. `PdfSaveResult.Serialization` records limits, peak retained bytes, spill decisions, final buffering, and passthrough without claiming forward-only layout. Per-page metadata and the authored block model remain proportional to document size, and `ToBytes()` buffers the final artifact.
- Provides conversion reports, grouped warning summaries, and diagnostics so adapters can expose unsupported or simplified source content honestly.
- Provides reusable conversion proof snapshots for generated PDFs, artifact hashes, required page counts, page sizes, document metadata, outline titles, URI links, form fields, named destinations, page labels, attachments, output intents, optional-content/layer metadata, catalog/viewer metadata, XMP/tagged metadata, text markers, logical readback signals, expected and accepted warning contracts, and post-processing hand-off. Compliance proof records bind external validator name, version, profile, result, warnings, SHA-256, byte length, and validation time to the exact artifact.
- Provides reusable rewrite-preservation proof for page geometry, metadata, navigation, catalog/viewer/action state, optional content, tagged content, security signatures, document versions, and source-structure markers such as incremental updates, xref streams, and object streams.
- Provides a reusable rewrite-preservation matrix for classifying named manipulation scenarios as rewrite-safe, preservation-failed, blocked by safety checks, or operation-failed, including optional-content/layer drift, targeted form-fill preservation, form/tagged/active-content/signature blockers, and fluent `PdfDocument` helpers for normal document rewrite operations.
- Serves as the shared engine for Word, Excel, PowerPoint, OpenDocument, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters.

PNG embedding keeps scanned-page images compact without changing their pixels. See the [image capability matrix](../Docs/officeimo.image-export-capability-matrix.md) for supported sample depths and transparency behavior.

## Existing PDF workflows

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("input.pdf")
    .Pages.Extract("1-2,4")
    .MergeWith("appendix.pdf")
    .UpdateMetadata(title: "Merged report")
    .Stamp.Text("Reviewed")
    .Save("output.pdf");

PdfDocumentReadResult read = PdfDocument.Load("output.pdf").Read();
string text = string.Join('\n', read.Pages.SelectMany(page => page.TextBlocks).Select(block => block.Text));
```

`Load(...)` is the one entry point for byte arrays, files, and streams. It
enforces the same `PdfLoadOptions` limits before buffering, snapshots caller
input once, and reuses one parsed document across read, inspection, preflight,
diagnostic, optimization, signature, and compliance operations.

Lossless optimization through `document.Optimization.Apply(...)` keeps XMP packets uncompressed and unchanged, preserves embedded XML and its associated-file relationships, and retains output intents and ICC profiles. It preserves an absent Info dictionary for PDF/A-4. `MaximumCompression` uses object and cross-reference streams and rejects PDF/A-1; choose `Balanced`, `Archival` or `Web` for that profile. Filtered XMP inputs are rejected before writing output. Optimize before signing: a signed input is refused because a full rewrite would invalidate its signatures. The compliance fixture gate validates representative PDF/A-2b, PDF/A-3b with PDF/UA-1 invoices, PDF/A-4 and PDF/UA-2 before and after each optimizer profile; conformance claims still require validation of the exact output artifact.

### Read and edit named document JavaScript

```csharp
using OfficeIMO.Pdf;

PdfDocument document = PdfDocument.Load("input.pdf");

foreach (PdfJavaScript script in document.JavaScript.List()) {
    Console.WriteLine(script.Name);
}

PdfJavaScriptEditResult edited = document.JavaScript.Edit(scripts => scripts
    .AddOrReplace("Initialize", "this.zoom = 100;")
    .Remove("Obsolete"));

File.WriteAllBytes("output.pdf", edited.ToBytes());
```

Script names use exact, case-sensitive matching. Editing preserves untouched
name-tree entries and action data, then reads the saved artifact back before
returning it. Per-script, script-count, and aggregate-byte limits come from
`PdfLoadOptions.Limits`. Document JavaScript is active content: the default
sanitizer removes it, and full-rewrite edits are blocked for encrypted or signed
inputs rather than weakening their security or revision contracts.

### Preview and select active-content removal

```csharp
PdfDocument incoming = PdfDocument.Load("incoming.pdf");
var policy = new PdfSanitizationOptions {
    ActionKindsToRemove = PdfSanitizationActionKind.JavaScript |
        PdfSanitizationActionKind.Launch |
        PdfSanitizationActionKind.SubmitForm
};

PdfSanitizationReport preview = incoming.InspectSanitization(policy);
Console.WriteLine($"Scripts: {preview.ActionCounts.JavaScript}");
Console.WriteLine($"Launch actions: {preview.ActionCounts.Launch}");

PdfSanitizationResult result = incoming.Sanitize(policy);
File.WriteAllBytes("sanitized.pdf", result.ToBytes());
```

`ActionKindsToRemove` is an exact opt-in selection, so unselected action kinds
remain. Selecting `Uri` removes every URI action and catalog URI base, including
ordinary web links. Leave the property null for the established default policy:
known active-content actions are removed, allowed `http`, `https`, `mailto`, and
`tel` links remain, and URI schemes outside `AllowedUriSchemes` are removed.

### Inspect and sanitize before sharing

```csharp
PdfDocument incoming = PdfDocument.Load("incoming.pdf");
var policy = new PdfSanitizationOptions {
    ContentKindsToRemove = PdfSanitizationContentKind.All,
    ActionKindsToRemove = PdfSanitizationActionKind.All
};

PdfSanitizationReport preview = incoming.InspectSanitization(policy);
Console.WriteLine($"User metadata: {preview.CategoryCounts.UserMetadata}");
Console.WriteLine($"Attachments: {preview.CategoryCounts.EmbeddedFiles}");
Console.WriteLine($"Actions: {preview.CategoryCounts.Actions}");
Console.WriteLine($"Comments and markup: {preview.CategoryCounts.CommentsAndMarkup}");
Console.WriteLine($"Bookmarks: {preview.CategoryCounts.Bookmarks}");
Console.WriteLine($"Layer definitions: {preview.CategoryCounts.OptionalContent}");

PdfSanitizationResult result = incoming.Sanitize(policy);
File.WriteAllBytes("shared.pdf", result.ToBytes());
```

`ContentKindsToRemove` is an exact selection of user-authored Info fields and
XMP, embedded files, actions, comments and markup, bookmarks, and optional-content
definitions. Unselected categories remain. Link and Widget annotations remain
when comments are selected; selecting URI actions can still remove a link's URI.
Producer, creation and modification dates, and trapping status are not classified
as user metadata and remain in the output.

Selecting optional content removes layer definitions and associations. Drawing
and text operators that belonged to a layer remain as ordinary page content, so
they no longer depend on viewer layer state. Use content-safety inspection and
verified redaction when the page content itself must be removed.

### Add interactive fields to an existing PDF

```csharp
PdfAcroFormEditResult edited = PdfDocument.Load("input.pdf").Forms.Edit(form => form
    .Create(new PdfFormFieldCreateOptions {
        Name = "customer.notes",
        Kind = PdfFormFieldCreationKind.Text,
        PageNumber = 1,
        X = 72,
        Y = 560,
        Width = 240,
        Height = 60,
        Style = new PdfFormFieldStyle { IsMultiline = true }
    })
    .Create(new PdfFormFieldCreateOptions {
        Name = "calculate",
        Kind = PdfFormFieldCreationKind.PushButton,
        PageNumber = 1,
        X = 72,
        Y = 520,
        Width = 100,
        Height = 24,
        Caption = "Calculate",
        JavaScript = "this.getField('total').value = 42;"
    }));

File.WriteAllBytes("form.pdf", edited.ToBytes());
```

For a static paper-style form, inspect proposed fields before creating any widgets:

```csharp
PdfDocument source = PdfDocument.Load("static-form.pdf");
PdfStaticFormRecognitionReport proposals = source.Forms.RecognizeStaticLayout();
foreach (PdfStaticFormFieldProposal field in proposals.Proposals)
    Console.WriteLine($"{field.Index}: {field.Label} ({field.Kind}, {field.Confidence:0.00})");

int[] accepted = proposals.Proposals
    .Where(field => field.Confidence >= 0.6)
    .Select(field => field.Index)
    .ToArray();
if (accepted.Length > 0)
    File.WriteAllBytes("fillable-form.pdf", proposals.ApplySelected(accepted).ToBytes());
```

Recognition proposes text fields from empty continuous outlines or writing lines
and check boxes from small square outlines. Dashed outlines and writing lines are
skipped with an `unsupported-outline-dash` diagnostic because their painted
segments cannot be proven from the recognition geometry. Later fill, stroke, or image paint that
intersects a required outline segment is also skipped conservatively, including
partially erased boxes and writing lines. Exact rectangular path fills can prove
opaque covers; arbitrary or compound paths do not prove their entire bounds are
filled. Compound filled-and-stroked paths are not split into independent field
candidates, since their shared fill rule can leave holes. Nearby native text
supplies labels; callers
can also pass bounded positioned OCR text as `PdfStaticFormTextEvidence` without
installing an OCR runtime in `OfficeIMO.Pdf`. Native labels require perceptible
opacity, resolved paint, proven contrast, and full visibility within the visual
page and any rectangular clip. Patterned, dashed-stroke, or combined fill-and-stroke labels,
labels with partial or unproven clipping, and image backdrops need separate
visibility evidence such as positioned
OCR. Native text crossing a page or crop edge still counts as possible field
occupancy within the visible page. Labels partially covered by later paint are excluded.
A label names at most one nearest eligible field; equally plausible assignments
require manual authoring. Detected table text and table regions are excluded from
automatic form evidence. Duplicate labels retain the strongest evidence after
the OCR confidence adjustment, and native text occupancy respects exact
rectangular clips and conservative text-stroke envelopes. Overlapping labels
that support different fields require manual assignment. Unsupported graphics-state paint is excluded
from field evidence. The report includes page-local tab
order suggestions and collision diagnostics. It does not infer radio groups,
choice values, calculations, or form actions from static marks. Applying selected
proposals edits the analyzed PDF snapshot; a new analysis is needed to include
later changes to the source document. The usual PDF mutation and preservation
checks apply.

The same transaction creates text fields, check boxes, combo or list choices,
radio-button groups, push buttons, and empty signature fields. Generated widget
appearances use `PdfFormFieldStyle`; widget JavaScript is returned as inert,
typed `PdfFormWidgetAction` data and is never executed by OfficeIMO.Pdf. Form
edits and fills preserve eligible actions owned by form widgets. Unrelated
catalog, page, outline, or annotation active content remains a rewrite blocker.
Flattening removes actions owned by the fields being flattened, while the
sanitizer removes forbidden actions without discarding the remaining form tree.

For a single health and capability view:

```csharp
PdfAnalysisReport analysis = PdfDocument
    .Load("incoming.pdf")
    .Analyze(PdfComplianceProfile.PdfA2B);

Console.WriteLine($"Pages: {analysis.Info.PageCount}");
Console.WriteLine($"Readable: {analysis.CanRead}");
Console.WriteLine($"Rewrite safe: {analysis.CanRewrite}");
Console.WriteLine($"Healthy: {analysis.IsHealthy}");

foreach (PdfDiagnosticFinding finding in analysis.Diagnostics.Findings) {
    Console.WriteLine($"{finding.Severity}: {finding.Code} — {finding.Message}");
}
```

## Unified document API

The unified API intentionally narrows the public surface around the fluent
`PdfDocument` facade:

- Use `PdfDocument.Load(...)` for existing-document workflows and the instance reading members for extraction.
- Seekable PDF input streams are now consistently read from the beginning and
  restored to their original position. Non-seekable streams are read forward
  from their current position.
- Keep one opened `PdfDocument` and reuse it for `Read`, `Inspect`,
  `Analyze`, compliance, and manipulation work. The source snapshot and canonical
  parse are cached for that document.
- Typed reading rejects input with no readable indirect objects. Lenient parsing
  reports `PdfParseException.Code == "NoIndirectObjects"` for empty, plain-text,
  or unrecoverable header-only input. `Load` retains its lazy snapshot behavior;
  analysis and artifact snapshots retain diagnostic and source-identity
  evidence. Zero-page catalogs and recoverable object fragments remain readable.
- Use `PdfDocument.Analyze(...)` when a workflow needs the combined health,
  rewrite-safety, diagnostics, optimization, signature, repair, and compliance
  view.
- Use `CreateComplianceArtifact(...)` instead of separately rendering bytes and
  passing them back to `AssessComplianceProof(...)`. The returned immutable
  snapshot keeps exact output bytes and matching readiness evidence together,
  including for randomized encrypted output.
- Author new documents through `PdfDocument.Create(pdf => ...)` or append authored
  content through `Compose(...)`. Headings, paragraphs, tables, images, and other
  flow primitives live on `PdfContentBuilder`; they are no longer duplicated on the
  root document.
- Use the fluent `Pages`, `Forms`, `Attachments`, `Bookmarks`, `Annotations`,
  `Stamp`, `Security`, `Redactions`, `Optimization`, `Proof`, and metadata operations instead of the former public
  static engine classes. Those implementation engines are now internal so there
  is one supported route for each operation.
- Native `Save(...)` and `SaveAsync(...)` follow the shared document lifecycle and throw on failure. `PdfSaveOptions.AllowLoss` controls writer-time fidelity rejection. Adapter `SaveAsPdf(...)` returns `OfficeConversionReport`, including source and completed-output diagnostics.

The target-framework support remains `netstandard2.0`, `net8.0`, and
`net10.0`. See the [OfficeIMO migration guide](https://github.com/EvotecIT/OfficeIMO/blob/master/MIGRATION.md)
for the OfficeIMO 3.4 authoring changes and old-to-new API map.

### Low-level text extraction with inline scripts

`PdfReadPage.ExtractText()` keeps adjacent, smaller-font superscripts and
subscripts with a unique base line. For example, the following document extracts
as `m-1 tail` followed by `H2O`:

```csharp
using OfficeIMO.Pdf;

byte[] bytes = PdfDocument.Create(new PdfCreateOptions { DefaultFontSize = 18 })
    .Compose(document => document.Content(content => content
        .Paragraph(p => p.Text("m").Superscript("-1").Text(" tail"))
        .Paragraph(p => p.Text("H").Subscript("2").Text("O"))))
    .ToBytes();

string text = PdfDocument.Load(new MemoryStream(bytes)).Pages[0].ExtractText();
```

The shared layout uses transformed text dimensions for script attachment and
respects explicit breaks in logical fragment order. It leaves ambiguous anchors,
distant runs and ordinary neighboring rows separate. Setting either
`PdfTextLayoutOptions.LineMergeToleranceEm` below `0.6` or
`LineMergeMaxPoints` below `2.5` retains baseline-only grouping. This is a
conservative reading-order rule, not mathematical reconstruction. The separate
semantic understanding pipeline used by `PdfDocument.Read()` retains its own
line grouping.

When a closed `/ActualText` replacement maps completely to visible glyphs on
one forward baseline, low-level line construction uses the complete painted
extent for spacing and script attachment. `PdfTextSpan.Advance`, glyph evidence,
clipping and edit geometry retain their original paint values. Incomplete or
ambiguous mappings, mixed fonts or streams, incompatible transforms or clips,
and backward or multiline paint retain the conservative span extent.

## Examples

### Export PDF pages as images

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

using PdfDocument pdf = PdfDocument.Load("input.pdf");

pdf.Pages[0]
    .ToImage()
    .AtDpi(144)
    .AsThumbnail(800)
    .AsPng()
    .Save("preview.png");

pdf.ToImages()
    .Pages("1-3,last")
    .FitWithin(1600, 1200)
    .WithMaximumRasterPixels(20_000_000)
    .AsWebp()
    .Save("page-images");

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Authored PDF")
        .Paragraph(paragraph => paragraph.Text("The authored model uses the same page renderer."))))
    .ToImages()
    .AsPng()
    .Save("authored-page-images");
```

PNG, JPEG, TIFF, SVG, and WebP use the same `OfficeImageExportResult` contract and Drawing-owned encoders. Pixel-fit limits apply consistently to vector and raster output, and allocation limits are resolved before a raster buffer is created. Unsupported or simplified PDF operators and resources remain visible as typed image diagnostics.

When the PDF reader can prove a text span's code-to-glyph mapping, SVG export preserves its painted glyph contours and keeps the logical source in accessibility labels. Admitted finite zero-advance spans retain the same treatment. That text uses contours instead of relying on the PDF subset as a browser font; the outlines are not selectable SVG text. Unproven mappings and specialized text layouts keep the existing text-fallback limits. `SvgTextAsPaths` remains the opt-in for eligible caller-authored Drawing text.

PDF text rendering retains affine text geometry, including shear combined with rotation or nonuniform scaling. Source clipping stays in page coordinates, and positioned character and word spacing retain their PDF origins.
Glyph outlines transform before text strokes, so the text matrix does not distort stroke widths or dash spacing. Drawings read from PDF pages retain this geometry for raster and SVG export. PDF canvas authoring rejects imported affine glyph text explicitly; writing that drawing into a new PDF needs a qualified glyph-transform writer path.
Dashed outline joins can still differ from native PDF renderers at closed-path seams or when a dash starts exactly at a corner.

Scanned-page images support managed CCITT Group 3/4 decoding and packed 1-, 2-, and 4-bit DeviceGray samples. Opaque JPEG 2000 Gray/RGB images require a caller-supplied `IOfficeRasterImageCodec`; embedded or external masks, channel remapping, and color normalization remain diagnosed limitations. Missing scan decoders and malformed scan data fail raster rendering instead of producing a blank successful page. See the [OCR integration](../OfficeIMO.Pdf.Ocr/README.md#scan-rendering-and-execution-limits) for the decoder limits and scan-to-searchable-PDF workflow.

Convert a source document to `PdfDocument` and use the same paged-image bridge:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;

using var markdown = MarkdownDocument.Parse("# Page images");
using var converted = markdown.ToPdf(null, out OfficeIMO.OfficeConversionReport report);
IReadOnlyList<OfficeImageExportResult> pages = converted
    .ToImages()
    .AsPng()
    .Export();
```

Inspect `report.Diagnostics` for source conversion findings; each page result contains rendering findings. Use `PdfReadPage.ToDrawing()` only when an intermediate `OfficeDrawing` is needed. The returned `OfficeDrawingText.Text` keeps decoded, editable text, including multi-letter ligatures. Raster and SVG rendering use the PDF's painted glyphs where the embedded font requires them, preserving the painted order of shaped Arabic text.

Supply replacement fonts and any shaping provider during projection so glyph visibility uses the final font profile:

```csharp
var fonts = new OfficeFontFaceCollection();
fonts.Add("Courier New", File.ReadAllBytes("fonts/CourierSubstitute.ttf"));
OfficeDrawing drawing = document.Pages[0].ToDrawing(fonts);
```

Use the PDF's resolved font family name when registering a replacement. The overload also accepts `textShapingProvider`, `textShapingLanguage`, and `cancellationToken`. Supplied faces replace matching embedded faces and apply to nested drawings. Adding fonts after `ToDrawing()` cannot restore glyphs already discarded by visibility checks.

### Embedded-font Latin ligatures

Generated PDFs use `PdfTextShapingMode.OpenTypeLigatures` by default for embedded
TrueType and OpenType/CFF fonts. Measurement and painting use the same glyph run;
Unicode mappings retain the source text for extraction, search, and redaction.

Set `PdfCreateOptions.TextShapingMode = PdfTextShapingMode.UnicodeScalar` to retain
separate scalar glyphs and the previous wrapping behavior. Per-run feature
settings such as `OfficeTextFeatureSettings.Default.With("liga", 0)` disable
that optional feature; the language system's required feature remains enabled.
`LatinLigatures` remains available for the presentation-character
substitutions supported by that mode.

With `UnicodeScalar`, flat mixed-direction letter runs use separate directional
text-show operations and ToUnicode mappings to preserve copying and extraction.
Numeric, nested, and other ambiguous scalar contexts retain the logical source
in `ActualText`; extraction order in third-party readers depends on their handling
of bidi text and replacement text.

Pair kerning is on by default for embedded fonts (OpenType GPOS or `kern`
pairs) and for the standard Helvetica and Times fonts (Adobe AFM pairs). Set
`PdfCreateOptions.TextKerning = false` for nominal advances; runs with explicit
feature settings keep their own `kern` value.

For multilingual font shaping, `PdfCreateOptions.TextShapingProvider` accepts the shared `OfficeManagedTextShapingProvider.Instance` as well as an optional provider. The managed horizontal static TrueType/CFF profile applies selected Arabic joining forms and required ligatures, Latin/Hebrew/Arabic mark-to-base, mark-to-ligature and mark-to-mark attachment, contextual substitutions and supported pair adjustments. The automatic directional and marked-Latin routes use the same Core owner. Font-selected script and language features determine support; merely loading a font program does not qualify all of its layout operations.

When default managed OpenType shaping accepts a complete run, `PdfDocument.AnalyzeTextEncoding()` and PDF writing use its glyph coverage. A decomposed accent can use a covered precomposed glyph even when the font lacks the separate combining glyph; extraction retains the original Unicode text. `UnicodeScalar` checks the original scalars and requires their glyphs.

Canonical font-aware normalization retains each complete original Unicode source cluster once. Arabic mark ordering preserves joining controls and CGJ boundaries; retained nonpainting controls have zero advance and verified blank outlines. Native-compatible composition can use a covered composite across a lower-class intervening mark, which differs from UAX #53's adjacent-only recomposition guidance. Supported positioning can legitimately leave a zero-advance mark at its nominal font position when the font supplies no applicable parent anchor; no anchor is invented.

The static multilingual profile declines cursive attachment, device-dependent or variable anchors, feature variations, variation selectors, Indic and Southeast Asian syllable shaping, and true vertical glyph layout. Keep a full OpenType provider such as the optional `OfficeIMO.Drawing.HarfBuzz` adapter when those operations are needed. These limits are separate from the font loader's broader TrueType, CFF/CFF2 and variation support. Inspect conversion diagnostics and use a strict no-loss policy when approximated fallback is unacceptable.

See the [PDF font support contract](../Docs/officeimo.pdf.current-state.md#resources-fonts-and-trust)
for supported substitutions and limits, and the [shared managed shaping contract](../OfficeIMO.Core/README.md#managed-multilingual-shaping)
for source-cluster and provider behavior.

`PdfParagraphBuilder.FeatureSettings(...)` applies the same per-run policy to
subsequent plain, styled, linked and fallback text. Use
`ResetFeatureSettings()` to restore the default policy. Runs supplied through
`Runs(...)` retain their own feature settings.

```csharp
PdfDocument.Create(pdf => pdf.Content(content => content
    .Paragraph(paragraph => paragraph
        .FontFamily("Report Serif")
        .FeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1))
        .Text("AV")
        .ResetFeatureSettings()
        .Text(" Default text."))), options)
    .Save("text-features.pdf");
```

Register the named family in `options` before using it. Feature support depends
on the selected font and shaping provider. Prepared rich header and footer runs
retain their settings through glyph placement, alignment and decoration widths.

### Write a generated PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Page(page => page
        .Header(h => h.AlignCenter().Text("Service report"))
        .Footer(f => f.AlignRight().Text("Page {page} of {pages}"))
        .Content(layout => layout.Item(content => content
            .H1("Service report")
            .Paragraph(p => p
                .Text("Generated ")
                .Bold(DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm 'UTC'"))
                .Text(" with first-party PDF primitives."))
            .Table(new[] {
                new[] { "System", "Status", "Owner" },
                new[] { "Identity", "Green", "Operations" },
                new[] { "Messaging", "Yellow", "Exchange" }
            })))), new PdfCreateOptions {
            PageSize = PdfPageSizes.A4,
            Margins = PdfPageMargins.UniformCentimeters(1.6),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        })
    .Meta(
        title: "Service report",
        author: "OfficeIMO",
        subject: "Generated PDF")
    .Save("service-report.pdf");
```

Metadata, bookmark titles, link descriptions and logical text use PDFDocEncoding or UTF-16. Standard-font paint uses the font encoding separately. Page extraction and rewriting retain original parsed string bytes, including binary image palettes.

Generated headers and footers can combine literal text, visually styled runs, and styled page tokens. The same builder is available for the default, first-page, and even-page variants:

```csharp
PdfDocument.Create(pdf => pdf.Page(page => page
    .Header(header => header
        .Text(text => text
            .Run(PdfTextRun.Bolded("Confidential ", PdfColor.FromRgb(180, 0, 0)))
            .Text("- page ")
            .CurrentPage(PdfTextRun.Italicized(string.Empty))
            .Text(" of ")
            .TotalPages(PdfTextRun.Italicized(string.Empty)))
        .FirstPageText(text => text.Run(PdfTextRun.Bolded("Confidential cover")))
        .EvenPagesText(text => text.Run(PdfTextRun.Underlined("Confidential even page"))))
    .Content(layout => layout.Item(content => content
        .Paragraph(p => p.Text("Generated report body."))))))
    .Save("styled-header.pdf");
```

Styled header/footer runs support fonts, size, color, highlighting, underline, strike, and baseline changes. Use the existing header/footer image and shape methods for visuals; interactive links and inline elements are intentionally kept out of text runs. Authored header/footer content that enters margins or overlaps another zone is preserved instead of rejected; attach a `PdfConversionReport` with `ReportDiagnosticsTo(...)` when the host needs structured overflow or clipping warnings.

Use `DocumentPages()` or `DocumentPages(style)` for a page count across all document
sections. `TotalPages()` retains the count for the current numbering section.
Literal text segments keep braces and page-token-like text unchanged.

### Rich report layout

```csharp
PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Operational summary")
        .Paragraph(p => p
            .Text("Generated ")
            .Bold(DateTime.Today.ToString("yyyy-MM-dd"))
            .Text(" with links, lists, panels, and tables."))
        .Bullets(new[] {
            "No runtime package dependencies", "Word-like document flow", "Reusable PDF primitives for adapters"
        })
        .Panel(panel => panel
            .H2("Review note")
            .Paragraph(p => p.Text("Keep polished report designs in samples; keep reusable primitives in the engine.")))
        .Table(new[] {
            new[] { "Area", "Status" },
            new[] { "Layout", "Ready" },
            new[] { "Reading", "Evolving" }
        })))
    .Save("summary.pdf");
```

### Reusable business recipes

```csharp
var invoice = new PdfInvoiceComponent(
    invoiceNumber: "INV-42",
    issueDate: DateTime.Today,
    seller: new PdfInvoiceParty("Seller Ltd", new[] { "Tax ID 123" }),
    customer: new PdfInvoiceParty("Customer Ltd"),
    lines: new[] { new PdfInvoiceLine("Engineering", 2M, 50M, taxRate: 0.20M) },
    currencyCode: "EUR");

PdfDocument.Create(pdf => pdf.Content(content => content
        .Component(new PdfReportComponent("Delivery summary", "All checks passed."))
        .Component(invoice)))
    .Save("delivery-pack.pdf");
```

These recipes compose normal flow, table, and panel primitives. `IPdfContextComponent`
uses the existing deferred replay path when content must react to the live page number;
it does not introduce another layout engine.

### Fit text within one table cell

Use `PdfTableCell.WithShrinkTextToFit()` to reduce a cell's text size within its available width without reducing its neighbors. Combine it with `WithNoWrap()` to keep each source line together. `PdfTableStyle.MinimumShrinkFontSize` sets the font-size floor, with six points as the default. `PdfTableStyle.ShrinkTextToFit` retains its separate behavior of fitting the whole row.

```csharp
var account = PdfTableCell.TextCell("A long account identifier")
    .WithNoWrap()
    .WithShrinkTextToFit();
```

### Clipped table-cell fragments

Use `PdfTableCell.WithViewport` when a cell is represented by a fragment of a larger cell. The viewport describes the full box, the visible fragment and its offset in one coordinate system. Text wraps and aligns in the full box; the renderer clips it to the fragment. Diagonal borders use the full box too, while ordinary edge borders follow the fragment's configured border style.

```csharp
var lowerHalf = PdfTableCell.TextCell("Bottom aligned")
    .WithViewport(new PdfTableCellViewport(
        contentWidth: 100, contentHeight: 48,
        width: 100, height: 24, offsetY: 24));

var fragmentStyle = PdfTableStyles.Minimal();
fragmentStyle.HeaderRowCount = 0;
fragmentStyle.ColumnWidthPoints = new List<double?> { 100 };
fragmentStyle.FixedRowHeights = new List<double?> { 24 };
fragmentStyle.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom };

PdfDocument.Create().Compose(document => document.Page(page =>
    page.Content(content => content.Item(item =>
        item.Table(new[] { new[] { lowerHalf } }, style: fragmentStyle)))))
    .Save("cell-fragment.pdf");
```

The viewport scales with the rendered cell dimensions and preserves cell links and named destinations. Images, data bars and icons retain their full-cell geometry too and clip to each visible fragment. Rotated and cropped images remain visible when their drawn content crosses the fragment, and their links stay inside it. It works in normal flow, column flow and canvas tables. A viewport's complete visible row span must fit on one page; split larger cells into explicit fragments before rendering. A cell with check boxes or form fields rejects a viewport; place interactive fields separately. Pass `null` to `WithViewport` to return a copy with ordinary cell layout. `PdfCellIcon.HorizontalAlignment` can position an icon independently of the cell text.

### Floating tables

Set `PdfTableStyle.HorizontalOffset` to translate a table after alignment or floating placement. The value is measured in points: positive values move right and negative values move left. Translation preserves the available table width, column widths and vertical flow, including page continuations and tables inside rows or section columns.

Set `PdfTableStyle.Position` to a `PdfTablePosition` to place a table relative to the current text flow, page margins, or page edges. Offsets and text clearances are measured in points; positive vertical offsets move down the page. Paragraphs wrap beside the table and regain their full width below it. Wide inline objects move below the table when the side interval is too narrow. Headings, lists, images, and other structured blocks use space below intersecting floating tables.

Floating placement ignores the table's flow spacing (`SpacingBefore` and `SpacingAfter`). Deferred tables apply their top anchor once across all batches. Vertical `Inside` alignment uses the top edge on odd output pages and the bottom edge on even pages; `Outside` reverses those edges. Center, bottom, inside, and outside alignment require an eager table because a deferred table's total height is not known before its rows are streamed. Floating placement is supported in document flow; tables inside row columns reject `Position`, and Word multi-column sections retain an approximation diagnostic.

### Hyphenation and inline visuals

```csharp
byte[] statusIcon = File.ReadAllBytes("status.png");
var hyphenation = new PdfHyphenationLexicon(new[] {
    "auto-ma-tion",
    "ty-pog-ra-phy",
    "re-port-ing"
});

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph
            .Text("Automation status ")
            .InlineImage(statusIcon, 12, 12, alternativeText: "Healthy")
            .Text(" remains available during long reporting runs."))),
    new PdfCreateOptions().UseTextHyphenationDictionary(hyphenation))
    .Save("inline-status.pdf");
```

Inline elements participate in normal line wrapping. In tagged output, image and box alternative text is carried into the structure tree.

Lines wrap at Unicode line-break opportunities (UAX #14), including after hyphens and slashes and between CJK characters. For pattern-based hyphenation without a word list, use `new PdfCreateOptions().UseAutomaticHyphenation("en-US")`. The embedded patterns cover US English (`en-US`) and reformed German (`de-DE`); other languages are not hyphenated. Automatic hyphenation and `UseTextHyphenationDictionary` share one hyphenation hook, so the last one configured applies.

### Logical text for positioned canvas paint

Use `ActualText` to give positioned paint one logical string. When adjacent
spans need reliable extraction spacing, provide their extraction rectangles in
top-left page coordinates:

```csharp
byte[] logicalText = PdfDocument.Create(pdf => pdf.Content(content => content
    .Canvas(canvas => canvas.ActualText("Total: 42", 10, 28, 80, 12,
        paint => paint.Text("Total: 42", 10, 28, 80, 12))))).ToBytes();
```

The rectangle describes logical text geometry; it does not resize the paint.
Paint and replacement text share an owner, so editing the replacement alone is
rejected when it cannot safely preserve the painted content. The existing
overloads without a rectangle remain available for grouped paint and point
anchors.

Associate a recoverable source with an individual structure element using
`PdfCanvasStructureOptions.AddAssociatedFile`. For a MathML supplement, provide
`application/mathml+xml` and the `Supplement` relationship:

```csharp
var formulaStructure = new PdfCanvasStructureOptions { AlternativeText = "x squared" }
    .AddAssociatedFile(new PdfEmbeddedFile("x-squared.mathml",
        System.Text.Encoding.UTF8.GetBytes(
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>x</mi><mn>2</mn></msup></math>"),
        "application/mathml+xml", PdfAssociatedFileRelationship.Supplement));

byte[] mathematicalPdf = PdfDocument.Create(pdf => pdf.Content(content => content
    .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Formula,
        expression => expression.ActualText("x^(2)",
            paint => paint.Text("x2", 10, 10, 60, 20)), formulaStructure))),
    new PdfCreateOptions().EnableTaggedPdfCatalogMarkers()).ToBytes();
```

These files are attached to the structure element's `/AF` entry and are also
discoverable through `PdfDocument.Attachments`. They are not document-level
associations. Plain output selects PDF 2.0 when a structure association is emitted;
PDF/A-3 groundwork retains its PDF 1.7 version. Tagging must be enabled. File names
are unique across catalog and structure attachments; identical structure files
share a payload, while conflicting descriptions or bytes fail generation. Supply
`ModificationDate` when archival policy requires it. Undated PDF 2.0 associated files omit
the optional stream-parameter dictionary; catalog associations in PDF 2.0 also
require a MIME type. Explicit profile checks include structure-file dates and
the selected output version. This API does not request or prove a
compliance profile.

### Attachment authoring limits

Catalog authoring accepts up to 128 MiB per file through
`PdfCreateOptions.MaxEmbeddedFileBytes` and 256 MiB across the catalog through
`MaxTotalEmbeddedFileBytes` by default. `AddEmbeddedFile` checks these limits
before copying payload bytes into the collection. A rejected add or limit
reduction throws `InvalidDataException` and leaves existing files and settings
unchanged. `ClearEmbeddedFiles()` restores the collection's available capacity.

Structure attachments have the same defaults through
`PdfCanvasStructureOptions.MaxAssociatedFileBytes` and
`MaxTotalAssociatedFileBytes`. Each structure collection enforces its own
admission limits. At save, the document options also apply to every emitted
attachment and to their combined payload total. Catalog files count once each;
identical structure associations that share one emitted file specification count
once across pages. Different files count separately even when their bytes match.

Set explicit larger limits before adding trusted larger inputs. For example,
these settings accept a file up to 256 MiB and a collection up to 512 MiB:

```csharp
var attachmentOptions = new PdfCreateOptions {
    MaxEmbeddedFileBytes = 256L * 1024 * 1024,
    MaxTotalEmbeddedFileBytes = 512L * 1024 * 1024
};

var structureOptions = new PdfCanvasStructureOptions {
    MaxAssociatedFileBytes = 256L * 1024 * 1024,
    MaxTotalAssociatedFileBytes = 512L * 1024 * 1024
};
```

All four limits must be positive. Standalone `PdfEmbeddedFile` descriptions keep
their existing construction behavior; collection admission applies the selected
limits. A combined-limit failure occurs before output bytes are written, so it
preserves an existing destination file or stream. These limits bound admitted
and emitted payload bytes, not defensive snapshot multiplicity or peak process
memory. Use `MaxGeneratedOutputBytes` separately to bound the serialized PDF.
Attachment reading and editing use `PdfLoadOptions.Limits.MaxTotalAttachmentBytes`
as described below.

### Sections, generated navigation, and bounded stream output

```csharp
var options = new PdfCreateOptions {
    PageContentMemoryLimitBytes = 4 * 1024 * 1024,
    ObjectBufferMemoryLimitBytes = 8 * 1024 * 1024
};

PdfDocument.Create(pdf => pdf.Content(content => content
        .TableOfContents()
        .Section("Summary", section => section
            .Element(element => element
                .Background(PdfColor.FromRgb(248, 250, 252))
                .Border(PdfColor.FromRgb(148, 163, 184))
                .Padding(10)
                .KeepTogether()
                .Content(summary => summary
                    .Paragraph(p => p.Text("A styled, keep-together summary.")))))
        .Section("Details", section => section
            .Columns(columns => {
                columns.Paragraph(p => p.Text("First column"));
                columns.ColumnBreak();
                columns.Paragraph(p => p.Text("Second column"));
            }, new PdfMultiColumnOptions { ColumnCount = 2, Gap = 18 }))), options)
    .Save("navigable-report.pdf");

```

Equal-width `Columns` balance their final page by default. Set
`PdfMultiColumnOptions.BalanceTableRowLines = true` to allow a splittable table row
to continue across balanced columns at legal cell paragraph boundaries. Fragment
measurement includes cell padding and repeated headers and preserves fitting
paragraph keep and widow rules. The default balances whole rows; set
`BalanceLastPage = false` to fill column frames sequentially.

Use `ColumnDefinitions` to give sequential columns different widths and gutters:

```csharp
using var document = PdfDocument.Create();
document.Content.Columns(content =>
    content.Paragraph(paragraph => paragraph.Text("Text flows through each column in order.")),
    new PdfMultiColumnOptions {
        BalanceLastPage = false,
        Gap = 18,
        ColumnDefinitions = new[] {
            new PdfFlowColumn(PdfColumnWidth.Fixed(96), gapAfter: 24),
            new PdfFlowColumn(PdfColumnWidth.Percent(25)),
            new PdfFlowColumn(PdfColumnWidth.Relative(1))
        }
    });
document.Save("columns.pdf");
```

Fixed widths and gutters use points. Percentages use the content width after
gutters; relative weights divide the remaining width. `GapAfter` overrides `Gap`
for the following gutter, and the final column has no following gutter. The
definitions determine `ColumnCount`; an explicitly assigned count must match.
Widths must leave positive space for every column. Automatic content-sized widths
are unsupported. Unequal columns fill sequentially; final-page balancing applies
to equal widths.

When columns begin partway down a page, table keep rules and minimum or fixed row
heights can move content to the next full page. Kept tables, unsplittable rows and
complete cell viewports must fit an available physical page. Set
`FinalColumnSpacingAfter` to add space below the last occupied column after
balancing. This space preserves the content's column breaks and stops at the
physical page's bottom margin.

### Load once and build one semantic read result

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("statement.pdf");
PdfDocumentReadResult result = pdf.Read(); // Structured is the default profile.

foreach (PdfLogicalPage page in result.Pages) {
    foreach (PdfLogicalParagraph paragraph in page.Paragraphs) {
        Console.WriteLine(paragraph.Text);
    }

    foreach (PdfLogicalTable table in page.Tables) {
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);
        Console.WriteLine($"Page {page.PageNumber}: {data.Rows.Count} table rows");
    }
}

string markdown = result.ExtractMarkdown();
PdfMetadata metadata = result.Metadata;
PdfDocumentSecurityInfo security = result.Security;
IReadOnlyList<PdfOutlineItem> outlines = result.Outlines;
IReadOnlyList<PdfLogicalLinkAnnotation> links = result.Links;
IReadOnlyList<PdfFormField> formFields = result.FormFields;
IReadOnlyList<PdfAttachmentInfo> attachmentMetadata = result.Attachments;

IReadOnlyList<PdfExtractedImage> images = pdf.Images.Extract();
IReadOnlyList<PdfImagePlacement> placements = pdf.Images.Placements("1-2");
IReadOnlyList<PdfExtractedAttachment> attachments = pdf.Attachments.Extract();
```

For inspection workflows, `pdf.ReadForInspection()` retains readable pages when a selected page has unmappable text or an unsupported content-stream filter:

```csharp
PdfInspectionReadResult inspection = pdf.ReadForInspection();
foreach (PdfPageReadOutcome outcome in inspection.PageOutcomes) {
    Console.WriteLine($"Page {outcome.PageNumber}: {outcome.Status}");
}
string readableText = inspection.Document.Text;
PdfDocumentInfo sourceStructure = inspection.SourceInfo;
PdfRepairReport repairs = inspection.RepairReport;
```

Outcomes preserve caller order and repeated page selections. Failed pages have no logical page; their geometry, links, annotations and document metadata remain in `SourceInfo`. `CompletedAllSelectedPages` describes page completion only. Authentication, permission, parser, resource-limit, cancellation and unexpected stage failures still stop the operation. Ordinary `Read()` keeps its existing fail-fast behavior. Lenient parsing reports recoveries and detected-only issues separately; `PdfLoadOptions.ParsingMode = PdfParsingMode.Strict` rejects structural defects.

For native text and metadata without image assets or secondary form-widget objects, use `ReadForInspection(new PdfLoadOptions { IncludeImages = false, IncludeFormWidgets = false })`. The same options apply to ordinary `Read()`. `IncludesImages` and `IncludesFormWidgets` on the logical result record the requested projections. Native text inside Form XObjects is retained. Source field facts and values remain subject to the canonical parser's form budgets; these flags omit later image and widget construction. Excluding images also omits the optional visual primitive scan, which can materialize Type3 glyph images; `page.HasVisualPrimitiveCounts` reports whether its counts are available.

Each logical page also retains `IncludesImages` and `IncludesFormWidgets`, including after page selection. An empty unrequested collection does not establish source absence. `AnalyzeExtractionScope()` and logical PDF conversions that report content loss require these projections and visual counts; they throw `InvalidOperationException` for a reduced read. Text and table extraction remain available. Use the default full read when you need omission analysis or conversion diagnostics.

Each `page.Analysis.DecodedRuns` item exposes `DecodingEvidence`. It distinguishes ToUnicode, predefined Adobe maps, font encodings, Windows-1252 fallback and ActualText replacement. `UnmappedCodeBytes` counts raw character fallback inside a ToUnicode run; `UsedFallback` describes the returned logical text. ActualText retains the underlying mapping source without claiming that its replacement matches the painted glyphs. A public `PdfUnsupportedTextMappingException` exposes the font resource, base font, encoding and declared-ToUnicode state when extraction cannot map shown codes. These facts describe decoding, not text accuracy, author identity or visual fidelity.

`PdfLoadOptions.PreferToUnicode` selects the Unicode-map path, `UseWinAnsiFallback = false` rejects unresolved mappings instead of emitting fallback characters, and `AdjustKerningFromTJ` controls positioned advances between strings in `TJ` arrays. Load settings and parser budgets can be copied with `Clone()`.

`result.ExtractMarkdown()` uses canonical reading order and emits detected tables once, including tables reconstructed from OCR geometry. `result.ExportStructured(PdfStructuredExportFormat.Json)` preserves positioned `lines` and includes canonical page `text` and detected `tables`. Each table contains its source kind, coordinate space, vertical bounds and original cell `rows`; these additive fields retain the `officeimo.pdf.logical.v1` schema identifier.

For documents with many embedded images, `pdf.Images.Visit(image => { ... }, cancellationToken)`
processes images one at a time. Inside the callback, `image.CopyTo(output, cancellationToken)`
copies its file bytes to a stream without cloning the payload. Check `image.IsImageFile`
before saving it as a standalone image file.

When the application needs only one embedded file, select its metadata from
`pdf.Inspect().Attachments` and call `pdf.Attachments.Extract(selected, maximumDecodedBytes)`.
This bounds the selected decoded payload and avoids decoding the other files.
`pdf.Attachments.Remove(selected)` removes the matching file specification when
more than one attachment uses the same display name.

Attachment edits apply `PdfLoadOptions.Limits.MaxTotalAttachmentBytes` to the
aggregate payload of the target collection. `Add` and `Replace` reject an
oversized target with `PdfReadLimitException` before cloning the new payload or
rewriting the PDF. A rejected operation leaves the edit session unchanged, so
the callback can remove another attachment or choose a smaller replacement.
Load the source with an explicit larger limit when the intended collection
exceeds the default 256 MiB. This payload limit does not represent peak process
memory; defensive snapshots and rewrite validation can retain additional copies.

`PdfDocument.Read(...)` is the only semantic reconstruction entry point. Both
profiles return `PdfDocumentReadResult`; they do not maintain separate logical
models. `Structured` adds document-wide tagged-PDF, outline, repeated-edge,
heading-tier, hierarchy, and continuation evidence. `Fast` keeps the same page
pipeline and result contract while omitting optional document-wide enrichment:

```csharp
PdfDocumentReadResult selected = pdf.Read(new PdfLoadOptions {
    Profile = PdfReadProfile.Fast,
    PageSelection = PdfPageSelection.Parse("1-2,5")
});
```

Every built-in semantic stage observes the supplied cancellation token while it
works, not only between pages. Custom pipeline stages receive the same token and
work-budget context. `PdfUnderstandingPipelineOptions.MaxWorkUnitsPerPage` and
`MaxDocumentWorkUnits` bound geometry comparisons, recursive partitioning, and
document-wide evidence; over-complex input fails with
`PdfReadLimitKind.UnderstandingWork` instead of continuing unbounded work.

Structured reads associate tagged list bodies with their nearest `LI`, including
text inside paragraph and link descendants. A `Lbl` supplies the marker; item
text excludes it. Nested items keep separate bodies and nesting levels, and an
item spanning pages has one fragment per page. Paragraph, heading, and table
classifications remain available. Text exports preserve headings, tables, and
unrelated content on shared lines instead of repeating them inside a flattened
list item.

Each page result exposes `ImagePlacements` and one `ImageRegion` per placement
invocation. Regions retain placement geometry and paint order, associate nearby
caption elements through vertical proximity and horizontal alignment only when
the image has a non-transparent, clipped intersection with the page, and expose
confidence plus stable evidence codes. Structured reads also use tagged-PDF
Figure ownership, including marked content inside Form XObjects, and retain the
owning Figure's alternate text. An unambiguous tagged structure relationship
also binds a direct child caption, or a boundary caption under a parent that has
exactly one Figure, without depending on caption wording or page proximity.
Multi-figure structures remain unassociated unless their ownership is explicit.
Artifact-marked images remain extractable but do not inherit structural Figure
ownership.

```csharp
foreach (PdfLogicalPage page in result.Pages) {
    foreach (PdfUnderstandingImageRegion region in page.Analysis.ImageRegions) {
        Console.WriteLine(
            $"{region.Placement.X}, {region.Placement.Y}: " +
            $"figure={region.IsFigure}, caption={region.Caption?.Region.Text}");
    }
}
```

Set `PdfUnderstandingPipelineOptions.ImageRegionDetection` to replace the
association stage. A custom stage must return exactly one region for each input
placement and may associate captions only from the semantic-stage result. The
built-in stage does not inspect language-specific caption words.
`MaxImageCaptionCandidatesPerPage` bounds viable image-caption edges before
matching and sorting, so adversarial overlap fails with a typed read limit.

### Inspect fonts, image paint, and cross-page continuations

The font inventory reports every unique declared font dictionary and every page
or nested Form XObject resource path that references it. Embedded OpenType and
TrueType programs are parsed through the same bounded font engine used by PDF
generation. The parsed glyph count describes the embedded program; it is not a
count of glyphs painted by the page.

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("annual-report.pdf");
PdfDocumentReadResult result = pdf.Read();
PdfFontInventory fonts = pdf.Resources.Fonts();

foreach (PdfFontInfo font in fonts.Fonts) {
    Console.WriteLine($"{font.FamilyName}: embedded={font.IsEmbedded}, subset={font.IsSubset}");
    if (font.EmbeddedOpenTypeInfo is PdfOpenTypeFontInfo program) {
        Console.WriteLine($"Program glyphs: {program.GlyphCount}; Unicode mappings: {program.UnicodeScalarCount}");
    }
}

IReadOnlyList<PdfImagePlacement> placements = pdf.Images.Placements();
foreach (PdfImagePlacement placement in placements) {
    Console.WriteLine(
        $"Page {placement.PageNumber}: {placement.Width} x {placement.Height}, " +
        $"opacity={placement.Opacity}, blend={placement.EffectiveBlendMode}, clipped={placement.Clip is not null}");
}

IReadOnlyList<PdfLogicalParagraphContinuationGroup> paragraphs =
    result.GetParagraphContinuationGroups(new PdfLogicalParagraphContinuationOptions {
        MinimumConfidence = 0.80
    });
IReadOnlyList<PdfLogicalTableContinuationGroup> tables =
    result.GetTableContinuationGroups(new PdfLogicalTableContinuationOptions {
        MinimumConfidence = 0.80
    });

foreach (PdfLogicalParagraphContinuationGroup paragraph in paragraphs.Where(item => item.SpansPages)) {
    Console.WriteLine($"Paragraph {paragraph.FirstPageNumber}-{paragraph.LastPageNumber}: {paragraph.Confidence:P0}");
}
foreach (PdfLogicalTableContinuationGroup table in tables.Where(item => item.SpansPages)) {
    Console.WriteLine($"Table {table.FirstPageNumber}-{table.LastPageNumber}: {table.TotalRowCount} rows");
}
```

Cross-page recovery is conservative. It joins only adjacent boundary items with
compatible geometry and content evidence, returns confidence and evidence flags,
and leaves the original page-local logical model unchanged. Embedded program bytes
are returned only when `IncludeEmbeddedProgramBytes` is explicitly enabled and
remain subject to both `MaxEmbeddedProgramBytes` per font and
`MaxTotalDecodedFontBytes` across embedded programs and ToUnicode maps. ToUnicode
per-map and aggregate limit outcomes have separate diagnostics and inventory counts,
so resource limits are not reported as malformed maps. Nested Form XObject discovery
deduplicates shared resource contexts by their shallowest reachable depth and is
bounded by `MaxFormResourceTraversals`.
For image paint inspection, `AuthoredBlendMode` is null when the normal PDF default
was not declared and retains an explicit or inherited authored `Normal` value.

The predefined horizontal composite encodings `UniJIS-UCS2-H`, `UniGB-UCS2-H`,
`UniCNS-UCS2-H` and `UniKS-UCS2-H` use bundled Adobe character maps for text
extraction and redaction search when the font has no `ToUnicode` map. The font's
Adobe character collection must match the encoding. Widths are read by the mapped
CID, so precise removal preserves neighboring text positions. Mapping data is
loaded locally and lazily; no external executable or download is required. See
[third-party notices](THIRD-PARTY-NOTICES.md) for source and license information.

Managed previews of unembedded Adobe CJK faces use installed script-capable
substitutes when available. No replacement fonts are bundled. Rendering retains
the font-substitution diagnostic: readable text does not establish identical
outlines, regional glyph forms or appearance. Supply matching fonts during
`ToDrawing()` projection when a controlled rendering profile is required.

An explicit `ToUnicode` map takes precedence and must cover every shown code;
an incomplete advertised map is refused. Other named encodings, including vertical
and half-width variants, require a usable `ToUnicode` map. If shown text lacks a supported mapping,
extraction throws `NotSupportedException` and redaction planning reports an error
instead of interpreting the character codes as WinAnsi text. Supply a PDF with
a supported encoding or an explicit Unicode map before using those text operations. Logical extraction
still uses explicit `ActualText` and excludes artifacts by default. Redaction
review requires mappings for painted text even under `ActualText` or artifacts.
Area-based redaction and image removal also reject a blocked inspection plan,
including unsupported text on an unselected page. Redaction refuses before
painting marks or replacing output files and streams.
Unused font resources do not block extraction; the existing `Identity-H` and `Identity-V` paths retain
their current behavior and editability limits.

Text extraction excludes PDF artifact marked content by default, which is the
logical-text behavior expected for decorative headers, footers, and chart
labels. Opt into visual text when those marked artifacts are part of the
required payload:

```csharp
PdfDocument visualTextPdf = PdfDocument.Load("spreadsheet-export.pdf", new PdfLoadOptions {
    IncludeArtifactText = true
});

PdfDocumentReadResult visual = visualTextPdf.Read();
IReadOnlyList<string> visualTextByPage = visual.Pages
    .Select(page => string.Join('\n', page.TextBlocks.Select(block => block.Text)))
    .ToArray();
```

Image-only and mixed pages use the optional `OfficeIMO.Pdf.Ocr` integration with
any `OfficeIMO.Ocr.IOcrEngine`. This keeps OCR contracts, process execution,
Tesseract, cloud SDKs, and native runtimes outside the base PDF graph. Accepted
words are normalized to cropped, rotated visual page coordinates,
de-duplicated against native text, and projected into the same logical model
used by reverse converters:

```csharp
using OfficeIMO.Ocr;
using OfficeIMO.Pdf.Ocr;

static async Task<PdfDocumentReadResult> ReadWithOcrAsync(
    string path,
    IOcrEngine engine) {
    PdfOcrMergeResult ocr = await PdfDocument.Load(path)
        .ReadWithOcrAsync(engine, new PdfOcrMergeOptions {
            MinimumConfidence = 0.75
        });

    Console.WriteLine(ocr.AcceptedWordCount);
    return ocr.Document;
}
```

`NativeDocument` retains the native-only view for overlap evidence and comparison.
Every parsed text block and inferred table exposes native-or-OCR provenance; OCR text also
retains provider confidence and direct visual bounds. OCR table evidence requires
repeated aligned columns, compact cell content, and stable row rhythm, so ordinary
two-column prose remains separate reading-order content. The work is bounded by
the OCR merge budgets and the shared `ReadOptions` pipeline limits. Word, Excel,
PowerPoint, HTML, RTF, ODT, ODS, and ODP packages consume `Document` directly
through their existing logical-PDF overloads. See the
[`OfficeIMO.Pdf.Ocr` README](../OfficeIMO.Pdf.Ocr/README.md) for installation,
searchable output, providers, and resource limits.

Native and hierarchy-free OCR text use the same direction policy. Automatic
mode selects the page's first strong Unicode direction; set an explicit policy
when a page is ambiguous or its dominant direction is known:

```csharp
PdfDocumentReadResult rtl = PdfDocument.Load("arabic-report.pdf").Read(
    new PdfLoadOptions {
        LayoutOptions = new PdfTextLayoutOptions {
            ReadingDirection = PdfReadingDirection.RightToLeft
        }
    });
```

Table schema remains evidence-based. Tagged header roles or distinct header
typography can establish the first row as a header. Otherwise `SchemaKind` is
`Unknown`, every source row remains data, and `Columns` contains empty names;
the parser does not invent translated or English column labels.

The structured reader handles repeated column gutters, mixed Latin/Hebrew/Arabic
fragments, and dominant quarter-turn text layouts while retaining source-page
geometry. Invisible searchable-text boxes use their selection geometry for line
and table recovery. Reconstructed columns expose full source-page rectangles in
`PdfLogicalTableColumn.VisualBounds`; at 90 or 270 degrees, distinct columns can
share the same X projection in `From` and `To`. Continuation matching uses their
actual progression axis. For OCR providers that merge columns into one line, opt into
[`ReconstructLayout`](../OfficeIMO.Pdf.Ocr/README.md#reconstruct-columns-and-mixed-direction-text).
The [Pango/Cairo layout corpus](../OfficeIMO.TestAssets/MultilingualLayout/README.md)
records exact fixture coverage and recognition limits.

### Rotate pages and change page boxes

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("packet.pdf");
using PdfDocument rotated = source.Pages.Rotate(90, "1,3-5");
rotated.Save("packet-rotated.pdf");

using PdfDocument crop = rotated.Pages.SetPageBox(
    PdfPageBoundaryBox.CropBox, 20, 20, 575, 822, 1);
crop.Save("packet-cropped.pdf");
```

Rotation sets an absolute angle in multiples of 90 degrees, normalized to
0, 90, 180, or 270. A page that inherits rotation receives its own override;
other pages keep their rotation. `Pages.Rotate(...)` provides the same behavior
when you only need the resulting document.

Use `Pages.RotateBy(90, 1, 3)` to turn pages relative to their current effective
rotation. Repeated clockwise turns advance through 90, 180, 270, and 0 degrees.
Each selected page resolves its own inherited angle, and an append-only turn
writes all selected page overrides in one revision. Studio's rotate buttons and
native engine use this relative-turn behavior.

Selectors such as `"odd"`, `"last"`, and `PdfPageSelection` work with
`RotateBy` and `RotateBy`. Use an ordered batch for several corrections:

```csharp
using PdfDocument edits = source.Pages.ApplyAttributes(new[] {
    PdfPageAttributeEdit.RotateBy(90, PdfPageSelection.From(1, 3)),
    PdfPageAttributeEdit.CropMargins(20, 20, 20, 20),
});
edits.Save("packet-edited.pdf");
```

Each edit sees the geometry produced by the preceding edits. The batch validates
and writes once; a failed or cancelled batch leaves the source document unchanged.
`ApplyAttributes` accepts a cancellation token for planning, writing, and
output readback; requested cancellation throws `OperationCanceledException`.

Rotation and page-box setters use a full rewrite when supported. When source
preservation prevents rewriting, they can append a revision containing only the
changed page dictionaries and cross-reference data. This supports cross-reference
streams, object streams, and tagged PDFs while retaining every original byte and
the structure tree. The appended cross-reference section follows the final source
section's format and links to it through `/Prev`. The operation result exposes
the selected route, reasons, and blockers.

`SetPageBox` supports `MediaBox`, `CropBox`, `TrimBox`, `BleedBox`, and
`ArtBox`; convenience methods include `SetCropBox` and `SetMediaBox`. Coordinates
are in PDF page space. Setting a box leaves source content in the file.
`PdfPageAttributeEdit.CropMargins` measures physical points (1/72 inch) from the
displayed left, bottom, right, and top edges. It accounts for inherited rotation,
page-box origins, and `/UserUnit`. `ResetCrop` restores the current media box.
Studio's margin crop and reset use these page-box edits and create one undoable
operation for all selected pages. Hidden content remains recoverable.
`CropAndTranslate` also changes content and annotation geometry, so it requires
a supported full rewrite.

Append-only page attributes retain the bytes covered by earlier approval
signatures. Signature trust and later-change policy remain separate validation
results. Certification/DocMDP restrictions forbid rotation and page-box changes
at every permission level; password permissions and usage rights also apply.
Deleting, reordering, inserting, and duplicating pages use their own preservation
checks. Page attributes preserve existing tags, but changing page geometry does not by itself establish PDF/A or
PDF/UA conformance; validate the saved artifact against the required profile.

Deletion, reordering, and duplication can normalize cross-reference and object
streams through a full rewrite when the remaining document structures can be
preserved. The plan reports `Input.CrossReferenceAndObjectStreamsWillBeNormalized`.
These operations do not use the page-attribute append-only route.

Tagged page organization uses a catalog-preserving full rewrite. `Delete`,
`Reorder`, `Move`, `Duplicate`, `Extract`, and `Split` retain surviving structure
elements and update page ownership, parent-tree entries, IDs, fields and navigation.
Copies receive separate page objects, structure elements, parent-tree keys and IDs,
including across successive copy operations. Marked-content references retain
their ownership whether their optional `/Type /MCR` entry is present or omitted.
`ReorderWithMapping` returns mappings for surviving original objects, including
annotations used by a host's selection state. The operation plan reports
`OrganizePages` and the required structure, form and annotation readback proofs.

`Insert`, `Append` and `Prepend` use `ImportPages` when combining tagged page
graphs. The target catalog and metadata remain, imported structure keys and IDs
are made distinct across successive imports, and role/class mappings are
namespaced. Imported structure roots retain their source language inheritance
and explicit language overrides. Incoming form fields,
layers, document name trees other than destinations, and unknown structure-root
extensions need explicit collision policies and currently fail with a specific
error. XFA, ambiguous ownership and unsupported surviving references to deleted
pages also fail before an output is adopted. Untagged content cannot be inserted
into a declared PDF/UA document without accessible tagging. These operations do
not infer a new semantic reading order for structure groups spanning several
pages. Import retains the target's metadata and color policy; it does not embed
missing fonts or reconcile conformance requirements for arbitrary incoming pages.
Validate accessibility and the declared conformance profile on the saved
artifact. The broad `ModifyPageTree` gate remains for other writers, including
content-changing crop and resize operations.

Use `Pages.NormalizeRotation()` to remove saved quarter-turn rotation while
keeping the pages' displayed appearance. Pass `PdfPageSelection.From(1, 3)` to
normalize selected pages, or use `NormalizeRotationResult` for the execution
plan. Normalization performs a full rewrite of the catalog-rooted graph: it
transforms content, boundary boxes, page pattern matrices, annotation appearances,
destinations, and page-associated tag layout bounds, then sets each selected page's `/Rotate` to
zero. Text and vector content remain editable, and cropped content remains in
the file. `/UserUnit`, tagged content, metadata, and standard AcroForm fields
are retained. Signed, certified, and usage-rights documents keep their rewrite
restrictions. Unsigned signature placeholders are retained, including invisible
widgets. Destination coordinates and fit axes follow the transformed default
page coordinate system; nullable coordinates and zoom values remain nullable.
Visible annotations require readable appearances; unassociated
tag layout bounds and shared page annotations produce explicit errors.

PDF page rotation is clockwise. Rendering, search results, interaction regions,
and visual-coordinate APIs use the same crop-, rotation-, and `/UserUnit`-aware
mapping. View-only rotation in Studio changes presentation and does not change
the saved `/Rotate` value.

### Split and extract pages

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("packet.pdf");

source.Pages.Extract("1-3")
    .Save("cover-and-summary.pdf");

IReadOnlyList<PdfDocument> singlePageDocuments = source.Pages.Split();
for (int index = 0; index < singlePageDocuments.Count; index++) {
    singlePageDocuments[index].Save($"packet-page-{index + 1:000}.pdf");
}

IReadOnlyList<PdfDocument> selectedRanges = source.Pages.Split("1-2,5-6");
selectedRanges[0].Save("packet-front.pdf");
selectedRanges[1].Save("packet-evidence.pdf");
```

For document-relative expressions such as `last,1,last`, resolve the immutable
`PdfPageSelector` against the inspected source page count. The overload
`selector.Resolve(pageCount, maximumPages: 100)` rejects a selection before
retaining more than 100 output pages, including repeats. Host workflows use this
bound before extraction or OCR and keep the source separate from the output.

To place source pages on printable sheets, use the same page import engine for
N-up or duplex booklet output:

```csharp
var nup = source.Pages.ImposeNUp(new PdfNUpOptions(new PdfPageSize(842, 595), columns: 2, rows: 1));
nup.ToDocument().Save("packet-two-up.pdf");

var booklet = source.Pages.ImposeBooklet(new PdfBookletOptions(new PdfPageSize(842, 595)));
booklet.ToDocument().Save("packet-booklet.pdf");
foreach (var placement in booklet.Placements)
    Console.WriteLine($"Source {placement.SourcePageNumber}: sheet side {placement.SheetPageNumber}, column {placement.Column}");
```

Booklet pages are ordered front then back for each physical sheet. Missing
pages are blank-padded to a multiple of four. Set `RightToLeft` for reverse
reading order; printer duplex edge, creep, and bleed settings remain with the
print workflow. Both operations import page
content as vector Form XObjects, but do not carry source annotations, forms,
structure tags, or selected TrimBox, BleedBox, and ArtBox boundaries onto the
new sheets; their visible appearances may also be omitted. Set
`AllowSourceFeatureLoss = true` to accept that loss and inspect
`SourceFeatureLoss` on the result. Set `MaxOutputBytes` on either layout option
to bound retained page content and the finished PDF (256 MiB by default).
Signed sources are rejected by default; set
`SignaturePolicy = PdfImpositionSignaturePolicy.CreateUnsignedDerivative` to
create an unsigned source derivative first. The result reports
`RemovedSignatureCount`, and the original signed PDF remains unchanged.
Layered PDFs are rejected because their visibility settings cannot yet be
preserved on imposed sheets.

For review after page insertion or reordering, align pages and inspect the
changed pairs with retained rendered proof and page-linked classifications:

```csharp
PdfDocument revised = PdfDocument.Load("packet-revised.pdf");
PdfReviewComparisonReport review = source.Proof.CompareReview(revised);
foreach (PdfReviewPageComparison page in review.Pages) {
    Console.WriteLine($"Page {page.ExpectedPageNumber} -> {page.ActualPageNumber}");
    if (page.Visual is not null)
        Console.WriteLine($"  {page.Visual.DifferentPixels} changed pixels");
    foreach (PdfReviewChange change in page.Changes)
        Console.WriteLine($"  {change.Kind}: {change.ExpectedText} -> {change.ActualText}");
}
```

`review.PageAlignment` reports inserted, deleted, moved, and changed page
candidates. Changed rendered pairs retain expected, actual, and highlighted
difference PNGs. Native text is also checked on visually identical pages, so
searchable-layer changes remain visible in the report. Native text and image
placements can be classified as added, removed,
changed, or moved; scan-dominated pages and unsupported differences remain
explicitly uncertain. `PdfReviewComparisonOptions` bounds page pairs, content,
pixels, and output bytes and accepts visual ignore regions. Classification is
review evidence, not a claim about authoring intent or OCR accuracy. An ignored
pixel region suppresses a semantic text block or image only when it fully
contains that element; a mask over part of a line leaves its text change for
review. When setting a custom render scale or background, use the same values
for `PageAlignment` and `Visual` so exact-page alignment remains valid review
evidence.

### Compare selected rendered page sequences

Use independent page selectors to review short ranges from long PDFs without
extracting temporary documents:

```csharp
PdfVisualComparisonReport report = source.Proof.CompareVisual(revised,
    cancellationToken, options: new PdfVisualComparisonOptions {
        ExpectedPages = PdfPageSelector.Parse("120-124,last"),
        ActualPages = PdfPageSelector.Parse("121-125,last"),
        MaxPages = 100
    });
string html = report.ToHtmlGallery("Section review", 96L * 1024 * 1024,
    cancellationToken);
```

Selectors retain caller order and repeated pages. Their first entries form the
first pair, their second entries form the second pair, and any remaining entries
are unmatched. Null selects every page on that side. `ExpectedPageCount` and
`ActualPageCount` are document totals; `ExpectedPageNumbers` and
`ActualPageNumbers` contain the selected source pages. Each paired result retains
both original page numbers. `UnmatchedExpectedPageNumbers` and
`UnmatchedActualPageNumbers` describe selected entries without a partner.

The standalone HTML gallery includes selected scope, totals, pair links, render
limitations and SHA-256 identities of the compared snapshots. This comparison
checks rendered appearances in the supplied order. It does not infer semantic
changes or moved-page alignment; use the separate review APIs above when that
analysis is needed. Selected-page, pixel, retained-image and HTML byte limits
remain enforced. The legacy shared `selection` argument cannot be combined with
independent selectors.

### Merge, reorder, delete, duplicate, move, and rotate

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("packet.pdf")
    .MergeWith("appendix.pdf")
    .Pages.Delete("2,5-6")
    .Pages.Duplicate("1")
    .Pages.Move(insertBeforePageNumber: 3, pageRanges: "7-8")
    .Pages.Rotate(90, "4")
    .UpdateMetadata(title: "Cleaned packet")
    .Save("packet-clean.pdf");
```

Use `PdfPageReorderPlan.Move(pageCount, insertBeforePageNumber, selectedPages)` to
preview a move without rewriting the document. `SourcePageNumbers` gives the full
proposed order, `GetOutputPageNumber(sourcePageNumber)` locates a page after the
move, and `HasChanges` identifies a no-op. `Shift(pageCount, towardStart, selectedPages)`
moves each selected run one position. Apply the proposed order with
`document.Pages.Reorder(plan.SourcePageNumbers.ToArray())`; the document's normal
mutation policy still applies. Plans describe page positions, so regenerate them
when the source document changes. `PdfPageSelection.Parse("2-4,7").Resolve(pageCount)`
resolves ranges with bounds checking and preserves caller order and repeated pages.

Load each source document, then merge the ordered models:

```csharp
using var first = PdfDocument.Load("first.pdf");
using var second = PdfDocument.Load("second.pdf");
using PdfDocument merged = PdfDocument.Merge(new[] { first, second });
merged.Save("merged.pdf");
```

Tagged inputs keep their existing headings, reading order, language and page
relationships. Repeated inputs receive separate structure identities. Plain pages
stay plain: adding them does not create accessibility tags. A primary document
that declares PDF/UA rejects untagged additions. Validate the saved result against
the intended profile and review its reading and navigation behavior.

When combining tagged forms, choose `Forms = PdfMergeStructureMode.Combine` and
`FormFieldCollisions = PdfMergeCollisionMode.RenameIncoming` in `PdfMergePolicy`.
Policies that discard tagged widgets are rejected so their accessibility links
cannot silently point to removed fields.

Tagged merges retain primary metadata and bookmarks. Dropping or combining those
structures is rejected before assembly because their editors do not yet preserve
tags. Primary attachments are retained when incoming sources contain none;
attachment policies that require rewriting files are also rejected explicitly.

Encrypted merge inputs keep independent authentication settings. Owner
authorization is honored automatically. A user password follows the PDF
permission bits unless the caller explicitly opts into ignoring those
restrictions:

```csharp
PdfDocument first = PdfDocument.Load("first.pdf", new PdfLoadOptions {
    Password = "first-owner-password"
});
PdfDocument second = PdfDocument.Load("second.pdf", new PdfLoadOptions {
    Password = "second-user-password",
    PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
});

using PdfDocument merged = PdfDocument.Merge(new[] { first, second },
    new PdfMergeOptions(), out PdfMergeReport report);

File.WriteAllBytes("merged.pdf", merged.ToBytes());
Console.WriteLine(report.OutputHasEncryption); // False
Console.WriteLine(report.Sources[1].PermissionRestrictionsIgnored); // True
```

`IgnoreRestrictions` is an authenticated permission override, not password
recovery. The document must still decrypt with the supplied password; an
unknown or incorrect password remains an error. Full rewrites of signed PDFs
remain blocked because they would invalidate existing signatures.

### Viewing protected PDFs

A user password can allow viewing while restricting extraction. Use the viewing
contract to obtain page geometry and render a page for display:

```csharp
PdfDocument protectedDocument = PdfDocument.Load("protected.pdf", new PdfLoadOptions {
    Password = "document-open-password"
});
PdfDocumentViewInfo view = protectedDocument.InspectForViewing();
PdfPageRenderResult page = protectedDocument.Render.DisplayPage(1, new PdfPageDisplayOptions {
    Scale = 1.5,
    MaximumPixels = 4_000_000,
    MaximumOutputBytes = 16 * 1024 * 1024
});
// Pass page.Bytes (PNG) to the host's image control.
```

`LogicalContent` is null when content extraction is restricted. The display path
returns flattened pixels and rendering diagnostics without exposing editable
drawing objects, source images, fonts, or a logical text model. It still requires
a valid password and observes parsing, pixel, output-size, timeout, and
cancellation limits. `Inspect`, `Read`, drawing/image export, page extraction,
printing, and mutation retain their respective permission checks. Display pixels
are not a DRM boundary: hosts must keep their copy, export, and print actions tied
to the authenticated permissions.

### Production document workflows

Apply one continuous Bates sequence across a batch. Each output includes its
page assignments and rewrite-preservation evidence:

```csharp
var inputs = new[] {
    new PdfBatesDocument(File.ReadAllBytes("volume-a.pdf"), "volume-a.pdf"),
    new PdfBatesDocument(File.ReadAllBytes("volume-b.pdf"), "volume-b.pdf")
};

PdfBatesBatchResult numbered = PdfBatesNumberer.Apply(inputs, new PdfBatesNumberingOptions {
    StartNumber = 1200,
    Prefix = "CASE-",
    MinimumDigits = 6,
    Position = PdfBatesPosition.BottomRight
});

File.WriteAllBytes("volume-a-numbered.pdf", numbered.Documents[0].ToBytes());
Console.WriteLine(numbered.NextNumber);
```

Every label is measured against its target page rectangle. A batch fails
instead of wrapping or clipping an identifier that cannot be rendered in full.

Interleave selected pages with explicit output provenance, including reverse
order for duplex scan backs:

```csharp
var fronts = new PdfInterleaveSource(File.ReadAllBytes("fronts.pdf"), "fronts");
var backs = new PdfInterleaveSource(File.ReadAllBytes("backs.pdf"), "backs") {
    Reverse = true
};

PdfInterleaveResult interleaved = PdfPageInterleaver.Interleave(
    new[] { fronts, backs },
    new PdfInterleaveOptions { RemainderMode = PdfInterleaveRemainderMode.Reject });

File.WriteAllBytes("duplex-scan.pdf", interleaved.ToBytes());
```

Page-addressed form fields and named links follow the selected pages and their
final output ownership; structures that belong only to excluded pages are not
carried into the composed document. Tagged sources are rejected by interleaving;
use `Merge` to combine complete tagged documents in source order.

Production splitting can combine page-count, text-boundary, and target-size
rules. Every part reports its source pages, termination reason, byte size, and
whether one indivisible page exceeded the target. Probe count and cumulative
generated bytes are bounded and reported for the complete operation:

```csharp
using var records = PdfDocument.Load("records.pdf");
IReadOnlyList<PdfProductionSplitPart> parts = records.Pages.SplitForProduction(
    new PdfProductionSplitOptions {
        MaximumPagesPerPart = 100,
        BoundaryText = "START OF RECORD",
        TargetPartSizeBytes = 20L * 1024 * 1024
    });

foreach (PdfProductionSplitPart part in parts) {
    File.WriteAllBytes($"records-{part.PartNumber:000}.pdf", part.ToBytes());
}
```

These operations reuse the existing selector, importer, stamping, mutation
preflight, and preservation engines. They do not bypass encryption,
permissions, signatures, or catalog-structure policy.

### Verified repair artifacts

Lenient parsing can persist only explicitly recovered structural defects into
a normalized artifact. The result must reopen in strict mode and pass the
rewrite-preservation comparison:

```csharp
PdfRepairArtifactResult repaired = PdfRepairArtifact.Create(
    File.ReadAllBytes("recovered-input.pdf"));

if (repaired.IsVerified) {
    File.WriteAllBytes("repaired.pdf", repaired.ToBytes());
}
```

Detected-only ambiguity is rejected by default. Encrypted PDFs, signatures,
certification permissions, and usage rights are not silently removed by the
repair workflow.

### Annotation review threads

Batch annotation edits use `document.Annotations.MoveMany`, `CopyMany`, `RemoveMany`,
and `Arrange`. Pass the indirect object numbers from the current
`document.Annotations.GetForEditing()` metadata or an authorized inspection snapshot.
Each operation expands standard group members and returns one mutation result, with
the same encryption, signature, and object-number mapping rules as individual edits.
Selections can span pages. `MoveMany` and `CopyMany` use PDF default user-space offsets;
`MoveManyVisual` and `CopyManyVisual` use top-left visual offsets and map each page's
crop origin, rotation, and user-unit scale independently in the same mutation.
`Arrange` accepts `Raise`, `Lower`, `BringToFront`, or `SendToBack` from
`PdfAnnotationOrderChange` and preserves the selected annotations' relative order.

`Annotations.GetForEditing()` returns annotation metadata, excluding form widgets,
when annotation mutation is permitted by encryption and certification policy.
`document.Annotations.GetEditingInteractions(pageNumber)` supplies annotation-only visual hit regions
without extracted text, link targets, images, or form fields. Both support annotation
editing when copying page content is restricted; `Inspect()` and `Read()` retain
their content-extraction permission checks.

`document.Forms.GetForInteraction()` reads field values, options, widget geometry,
and whether XFA data is present under content-copy, form-filling, or form-authoring
authorization. `document.GetEditingInteractions(pageNumber)` combines permitted
annotation and form-widget hit regions without extracting page text or images.
These reads do not grant a write operation; use `PlanMutation` to check it.
Encrypted form filling still requires owner authorization or an explicit
`IgnoreRestrictions` policy for its full-rewrite path.

Authorized form-authoring and sanitization operations check preservation without
granting content extraction. Their `PdfRewritePreservationReport` exposes the
proof result and issues; `CanReadDetails` indicates whether the complete
`Original` and `Rewritten` inspection snapshots may be read. Those properties
enforce each input's content-extraction permission.

`document.Annotations.Group(objectNumbers)` groups at least two editable markup
annotations on one page using standard `/IRT` and `/RT /Group` relationships.
Grouping requires PDF 1.6; older ordinary documents receive a catalog version
upgrade. Certified and usage-rights documents that would require that upgrade are
rejected because it changes the catalog beyond an annotation edit.
`Ungroup(objectNumbers)` removes those relationships from the complete selected
groups. `PdfAnnotationGrouping.GetMembers(info.Annotations, objectNumber)` resolves
a saved group without including conversational replies. Links, redaction marks,
and reply annotations cannot become group members. Locked and read-only annotations
are rejected by batch edits. Multi-selection geometry edits move annotations;
resize each annotation separately. Free-text callout transforms remain unsupported.

Copies stay on the source pages, retain supported appearances and styles, and receive
unique `/NM` names. They preserve the copied groups and linked popups, without copying
conversational replies or prior review decisions. Removal retains replies to a removed
parent as standalone comments. Append-only removal needs explicit
`allowResidualDataInAppendOnly: true` because older revisions retain removed data;
it is not a sanitization operation.

Read reply relationships as threads, add a reply, and record a standard review
state through the same annotation mutation policy used by lower-level edits:

```csharp
byte[] source = File.ReadAllBytes("review.pdf");
PdfAnnotationReviewCatalog review = PdfAnnotationReviewCatalog.Read(source);
int rootObjectNumber = review.Threads[0].Root.Annotation.ObjectNumber!.Value;

PdfAnnotationEditResult reply = PdfAnnotationReviewEditor.AddReply(
    source,
    rootObjectNumber,
    "Confirmed against the source record.",
    new PdfAnnotationReplyOptions { Author = "Reviewer", Subject = "Evidence" });

PdfAnnotationReviewCatalog updated = PdfAnnotationReviewCatalog.Read(reply.Bytes);
int replyObjectNumber = updated.Threads[0].Root.Replies[0].Annotation.ObjectNumber!.Value;
PdfAnnotationEditResult accepted = PdfAnnotationReviewEditor.SetState(
    reply.Bytes,
    replyObjectNumber,
    PdfAnnotationReviewState.Accepted);
```

Replies expose raw `/IRT`, `/RT`, `/State`, `/StateModel`, `/Subj`, and `/IT`
metadata alongside the typed standard state. Certified documents use an
append-only revision only when the signature permission model allows the
annotation change. Thread construction has explicit relationship and nesting
limits so hostile reply chains fail closed.

Full annotation rewrites expose `PdfAnnotationEditResult.AnnotationObjectNumberMap` when
the writer can report original-to-output object numbers. Use that map to retain
an annotation selection after changing its text or geometry; `/NM` names are
optional and are not guaranteed to be unique. A null map means no mapping was
reported. Append-only edits retain existing object numbers. For page reordering,
`document.Pages.ReorderWithMapping(2, 1)` returns a `PdfPageRewriteResult` with the
same mapping contract and a `ToDocument()` method for the rewritten PDF.

### Password protection on browser or restricted hosts

Desktop and server applications use platform AES automatically. A host without synchronous platform AES can pass the
managed provider included in `OfficeIMO.Core` explicitly; the same provider handles writing and opening the protected PDF.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Security;

IOfficeAesCryptographyProvider aes = OfficeManagedAesCryptographyProvider.Default;
var encryption = new PdfStandardEncryptionOptions("reader-password") {
    OwnerPassword = "owner-password",
    Algorithm = PdfStandardEncryptionAlgorithm.Aes256,
    AesCryptographyProvider = aes
};

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(p => p.Text("Protected in the current host."))),
        new PdfCreateOptions().SetEncryption(encryption))
    .Save("protected.pdf");

PdfDocument opened = PdfDocument.Load("protected.pdf", new PdfLoadOptions {
    Password = "owner-password",
    AesCryptographyProvider = aes
});
```

Supported page, text, form, metadata and redaction rewrites retain the authenticated
source's passwords, permissions and Standard-security settings. Extraction and
splitting retain that protection too. Merge output uses the primary source's
settings; an unencrypted primary produces an unencrypted output, and the merge
report records the decision. Use `Security.Decrypt(ownerPassword)` when the output
must be unencrypted. Unsupported encryption dictionaries and signed full rewrites
remain blocked.

### Certificate-based PDF signatures

PDF signature discovery, byte-range inspection, mutation blocking, and caller-defined external signing do not require
`OfficeIMO.Security`. For the built-in CMS adapter, install the optional package and pass its provider explicitly:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
using var signer = new PdfCmsExternalSigner(security, signingCertificate);

PdfExternalSignatureCompletion signed = PdfDocument
    .Load("contract.pdf")
    .Security.SignExternal(
        signer,
        new PdfExternalSignatureOptions {
            FieldName = "Approval",
            VisibleAppearance = new PdfVisibleSignatureAppearanceOptions {
                ImageBytes = File.ReadAllBytes("approval-mark.png"),
                ImageFit = OfficeImageFit.Contain,
                ShowText = false,
                ShowBackground = false,
                ShowBorder = false
            }
        });

var cryptography = new PdfCmsSignatureCryptographyProvider(
    security,
    new CmsVerificationOptions());
PdfSignatureValidationReport report = signed.ToDocument().Security.ValidateSignatures(cryptography);
```

Set `SubFilter = PdfExternalSignatureSubFilter.CadesDetached` in the signature
options for PAdES B-B. The built-in signer emits signing-certificate-v2 and omits
CMS signing-time for that subfilter. Both RSA and ECDSA certificates are supported.
For B-T, set `CmsSigningOptions.SignatureTimestampProvider` to an RFC 3161 token
provider, such as `Rfc3161TimestampClient.GetTimestamp`; its transport and trust
policy belong to the caller. Document timestamps use a separate RFC 3161 signer.

Text appearances require an explicit `PdfVisibleSignatureAppearanceOptions.Font`
containing an embedded TrueType font. Image-only appearances do not require it.
`PdfLongTermValidationEnricher.EnrichAsync` can merge caller-supplied certificates,
OCSP responses and CRLs with an optional caller-owned fetcher. DSS/VRI presence
alone does not establish complete trusted B-LT validation.

The PDF package owns byte ranges, incremental updates, signature dictionaries, and preservation policy. The optional
provider owns CMS, timestamps, and certificate trust. A custom `IPdfExternalSigner` or
`IPdfSignatureCryptographyProvider` remains valid without `OfficeIMO.Security`.
Authenticated encrypted signing appends a revision while retaining the original
bytes and encryption. Signature metadata and visible appearances are encrypted;
the detached CMS `/Contents` remains unencrypted as required by the PDF format.
The optional appearance image is visual content only; certificate validation remains the source of signer identity.
Set `ShowText`, `ShowBackground`, and `ShowBorder` to `false` for an image-only appearance. Existing appearances retain
their background and border because both new visibility options default to `true`.

### Review, apply, and verify redactions

Build a source-bound plan, review its areas and matches, then apply that exact plan and retain the evidence report with the output:

```csharp
PdfDocument source = PdfDocument.Load("contract.pdf");
PdfRedactionPlan plan = source.Redactions.Search(
    new PdfRedactionSearchOptions().AddLiteral("Account: 123-45-6789"));

// Present plan.Areas and plan.Matches for approval before applying it.
var verification = new PdfRedactionVerificationOptions {
    RequireCompleteStreamInspection = true,
    CheckManagedRendering = true
}.RequireRemovedText("Account: 123-45-6789");

PdfRedactionApplyResult redacted = source.Redactions.ApplyWithEvidence(
    plan,
    verificationOptions: verification);

redacted.ThrowIfUnverified();
File.WriteAllBytes("contract-redacted.pdf", redacted.Pdf);
Console.WriteLine(redacted.Evidence.Summary);
```

`Evidence.Items` records a verified-absent, residual, or inconclusive outcome for every reviewed match. The report also exposes source/output hashes, residual matches, verification details, and affected page numbers. A UI can pass those page numbers to the existing page renderer for before/after previews without making rendering part of the redaction contract.

Use `PdfRedactionSearchOptions.PageNumbers` to restrict candidate discovery to selected one-based pages. An empty set searches all pages. By default, search marks complete matching logical text blocks, so present the resulting areas for review rather than assuming that only the matched substring will be removed.

Set `TextSelection` to `MatchedGlyphs` to select only the encoded glyphs belonging
to literal or regular-expression occurrences. This uses the same native text flows,
wrapped-line matching, and table boundaries as located text search. `ContentScope`
controls the generated areas independently; choose `TextOnly` to preserve artwork
and images beneath the selected text:

```csharp
var search = new PdfRedactionSearchOptions {
    TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
    ContentScope = PdfRedactionContentScope.TextOnly
}.AddLiteral("Account: 123-45-6789");
PdfRedactionPlan precise = source.Redactions.Search(search);
// Present precise.Areas and any blocking precise.Findings before applying it.
PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(precise);
result.ThrowIfUnverified();
```

Touching native text fragments from the same match and baseline share one reviewed
rectangle. Wrapped lines and separated fragments remain separate areas. When editing
a review reason, use `area.WithLabel(reason)` to retain its exact geometry and
precise removal policy; constructing a new rectangle from its bounds loses that
evidence. After checking the search plan's `IsReviewable` and `Findings`, pass the
selected, relabeled areas to `source.Redactions.Plan(...)` for a fresh impact review.

Matched-glyph selection rejects partial ligatures, ambiguous ActualText mappings,
unsupported clipping or glyph evidence, and areas that intersect unselected text
(including hidden layers and tightly spaced lines whose conservative glyph bounds overlap).
Inspect `IsReviewable` and `Findings`: a blocked selection cannot be applied, even
when other criteria matched safely. Its reviewed areas never fall back to complete
text-object removal. Ordinary `Apply(plan)` also verifies preservation for these
areas before returning output. Regex criteria must select non-empty source text;
logical-kind criteria require complete-block selection. Glyph geometry remains the
reader's positioned width model, rather than an exact glyph-outline guarantee.
Complete logical blocks and `TextAndUnderlay` remain the search defaults.
Evidence for `TextOnly` includes text and annotation removals; preserved image and
vector underlays remain visible in the plan without being counted as removals.

Precise selection retains inherited font state and uses the source text matrix for
scaled or sheared glyph bounds. Rewriting preserves imported numeric font resources
and fractional review geometry. Independently produced TrueType and CFF regression
inputs cover neighboring text, bookmarks and ink preservation; the
[viewer checks](../Build/PdfViewerVerification/README.md) describe independent
readback and rendering. Outlined letters require a reviewed area because they are
vector paths rather than searchable text. Tight spacing can still block a substring
when its review envelope intersects an unselected glyph.

`source.Redactions.ApplyForSharing(plan, sanitizationOptions, verificationOptions: verification)` applies the reviewed redaction, sanitizes with the explicit policy, and verifies the final bytes. It requires successful sanitization, policy-specific preservation, unchanged page content and geometry, and final redaction checks. It does not bypass active-content or protected-document mutation gates. A policy that changes page content, such as flattening optional content, may need to be applied before planning redaction.

Configured removed and retained markers and external validators check the final
sanitized artifact. A marker in policy-selected metadata can remain until
sanitization; it must be absent from the final PDF. The intermediate redaction still
requires evidence for every reviewed area and the configured stream and rendering checks.

Save `PdfRedactionSharingResult.ToBytes()` and serialize its `Summary` when sharing content-free evidence. For redaction without sanitization, use `redacted.Evidence.CreateShareableSummary()`. These summaries include hexadecimal SHA-256 fingerprints and counts, but omit matched text, search criteria, reasons, paths, and detailed diagnostics. The detailed evidence remains suitable for local review and can contain sensitive document content.

Each `PdfRedactionArea` carries two independent policies. `TextOnly` removes selected text while preserving intersecting images and paths; `TextAndUnderlay` also removes supported intersecting underlay content and fails closed when that cannot be done safely. The appearance can remain exact, merge nearby reviewed marks, round widths to a configured quantum, or cover the effective page line. Verification evaluates residue against the selected content scope rather than treating a deliberately preserved underlay as a failure.

Review surfaces can also author standard PDF `/Redact` annotations and plan them later. Quadrilateral geometry is written as `/QuadPoints` and round-trips as exact destructive geometry:

```csharp
PdfDocument annotated = source.Redactions.AddAnnotation(new PdfRedactionAnnotationOptions(
    PdfRedactionRegion.Quadrilateral(1, new[] {
        new PdfRedactionPoint(72, 680),
        new PdfRedactionPoint(260, 680),
        new PdfRedactionPoint(250, 710),
        new PdfRedactionPoint(72, 706)
    })) {
    Contents = "Reviewed for removal"
});

PdfRedactionPlan annotationPlan = annotated.Redactions.PlanAnnotations();
```

Rectangle, quadrilateral, polygon, freehand, and grouped review geometry normalize to the same `PdfRedactionArea` collection used by planning, application, verification, and evidence. Intersection tests retain exact polygon and quadrilateral boundaries for text, annotations, images, hidden text, and pixel rewriting; freehand regions use exact round-capped segments. Destructive granularity still follows the supported PDF object: text is rewritten by glyph, annotations by annotation object, and raster images by pixel. A vector path wholly contained by a reviewed shape can be removed, while a partial intersection fails closed because deleting the complete path would remove paint outside the approved geometry. Raster rewriting selects pixel centers inside the exact shape, so its boundary precision is limited by the source image resolution.

Standard `/Redact` annotations honor every valid `/QuadPoints` quadrilateral instead of collapsing separated marks into the annotation `/Rect`; `/Rect` is the fallback when quad geometry is absent or invalid. Annotation authoring writes exact `/QuadPoints` for quadrilaterals. Polygon and freehand annotations use their bounded `/Rect` because PDF `/Redact` has no equivalent standard path geometry, while direct region planning retains their exact geometry. Annotation authoring defaults to 16 and permits at most 64 normalized areas because the current existing-page annotation writer performs one bounded full rewrite per area.

Text redaction rewrites native text-show operations at glyph granularity. Encoded glyphs outside the reviewed areas retain their original font resource and position; removed glyph advances become `TJ` displacements so adjacent text does not reflow. When a glyph mapping cannot be proven safe, the complete PDF text object is removed instead.

When `verificationOptions` is omitted, `ApplyWithEvidence` requires complete stream inspection and managed-rendering checks by default. Supply explicit options, as above, when the workflow also needs removed/retained markers or an external validator.

Permanent redaction invalidates existing signatures. `document.Security.CreateUnsignedDerivative()` performs an explicit full rewrite, removes signature fields, widgets, permissions, and validation-store references, and returns the unsigned bytes plus the number of detected signatures removed. It requires owner authorization for encrypted input and never overwrites the signed source. A host can then apply and verify redaction before starting a new `SignExternal` workflow for the derivative.

### Stamp and watermark an existing PDF

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("contract.pdf")
    .Stamp.Text("Reviewed", new PdfTextStampOptions {
        X = 72,
        Y = 720,
        FontSize = 18,
        Color = PdfColor.FromRgb(180, 30, 30)
    })
    .Stamp.TextWatermark("CONFIDENTIAL", new PdfTextStampOptions {
        FontSize = 54,
        Color = PdfColor.Gray,
        RotationDegrees = -35
    })
    .Save("contract-reviewed.pdf");
```

An image-backed Stamp annotation keeps the picture in the annotation layer,
which can be useful for a handwritten approval mark on an existing page:

```csharp
using System.IO;
using OfficeIMO.Pdf;

var result = PdfDocument.Load("contract.pdf").Annotations.AddStamp(
    new PdfStampAnnotationOptions {
        PageNumber = 1,
        X = 72, Y = 700, Width = 144, Height = 48,
        ImageBytes = File.ReadAllBytes("approval.png"),
        Contents = "Reviewed"
    });
result.ToDocument().Save("contract-marked.pdf");
```

This is a visual annotation, not a cryptographic PDF signature. The encoded
image input is limited to 128 MiB by default.

Use `Stamp.Watermark` when you want to reopen and revise a watermark later:

```csharp
using System.Linq;
using OfficeIMO.Pdf;

var watermark = new PdfWatermarkOptions {
    Text = "DRAFT",
    RotationDegrees = -30,
    Opacity = 0.3,
    Color = PdfColor.FromRgb(180, 30, 30),
    TargetPages = PdfPageSelector.Parse("1")
};
PdfDocument.Load("contract.pdf").Stamp.Watermark(watermark).Save("contract-draft.pdf");

var reopened = PdfDocument.Load("contract-draft.pdf");
var revision = reopened.Stamp.ReadWatermarks().Single(mark => mark.Id == watermark.Id);
revision.Text = "APPROVED";
revision.RotationDegrees = 0;
revision.TargetPages = null; // Apply the revised watermark to all pages.
reopened.Stamp.Watermark(revision).Save("contract-approved.pdf");
```

Each new `PdfWatermarkOptions` has its own identifier. Reusing that identifier
replaces its previous watermark and removes occurrences outside the new page
selection. Other watermarks keep their own identifiers. `ReadWatermarks` returns
editable settings for watermarks created by this API; it does not infer settings
from arbitrary existing text, images, or older `TextWatermark` stamps.

For an image watermark, set `ImageBytes = System.IO.File.ReadAllBytes("logo.png")`
instead of text. `Width` and `Height` set its size in points. `X` and `Y` use the
visual page's top-left origin; leaving either unset centers that axis. Rotation
is clockwise around the watermark's center, and opacity ranges from 0 to 1.

Import a complete source page above or below selected target pages without
rasterizing it:

```csharp
PdfDocument.Load("contract.pdf")
    .Stamp.OverlayPage("letterhead.pdf", new PdfPageOverlayOptions {
        SourcePageNumber = 1,
        TargetPages = PdfPageSelector.Parse("all,!last"),
        Fit = PdfPageOverlayFit.Contain,
        Opacity = 0.9
    })
    .Save("contract-with-letterhead.pdf");
```

For richer existing-page automation, stamp a general visual canvas instead of
using separate table-, text-, and image-only operations:

```csharp
PdfDocument.Load("contract.pdf")
    .Stamp.Content((canvas, page) => {
        canvas.Text($"Page {page.PageNumber} of {page.PageCount}", 36, 24, 220, 24)
            .Table(new[] {
                new[] { PdfTableCell.TextCell("Status"), PdfTableCell.TextCell("Reviewed") },
                new[] { PdfTableCell.TextCell("Owner"), PdfTableCell.RichTextCell(new[] { PdfTextRun.Bolded("Legal") }) }
            }, 36, 620, page.Width - 72, 90);
    }, new PdfCanvasStampOptions {
        TargetPages = PdfPageSelector.Parse("1,last"),
        Opacity = 0.95
    })
    .Save("contract-with-review-panel.pdf");
```

Canvas stamping is intentionally visual-only. Text, rich tables, images,
shapes, drawings, clipping, and effects are supported. Interactive links and
annotations, named destinations, forms, and document outlines use their
dedicated editors so their behavior is not silently flattened or discarded.

Use a scoped canvas blend mode when foreground artwork must composite with the
existing page. This example adds a Multiply highlight while keeping dark text
legible:

```csharp
using OfficeIMO.Drawing;

var highlight = OfficeShape.Rectangle(180, 18);
highlight.FillColor = OfficeColor.FromRgb(255, 230, 70);
highlight.StrokeWidth = 0;

PdfDocument.Load("contract.pdf")
    .Stamp.Content((canvas, page) => canvas
        .WithBlendMode(OfficeBlendMode.Multiply, blended =>
            blended.Shape(highlight, 72, 140)))
    .Save("contract-highlighted.pdf");
```

`Effect(transform, opacity, blendMode, build)` combines affine transforms,
group opacity, and any standard PDF blend mode. Set `BehindContent = true` on
`PdfCanvasStampOptions` for a true underlay; use the foreground default when
the blend must use existing page artwork as its backdrop.

### Search and edit existing page text

Text editing coordinates use PDF points from the page bottom-left. Inspect a
region when the UI needs the existing text and its detected style, then replace
or move it through the same text-removal and stamping owners used by redaction
and existing-page stamps:

```csharp
PdfDocument document = PdfDocument.Load("contract.pdf");
var region = new PdfPageRegion(pageNumber: 1, x: 72, y: 640, width: 260, height: 28);

PdfRegionText current = document.Text.Inspect(region);
Console.WriteLine($"{current.Text} ({current.SourceFont}, {current.FontSize} pt)");

PdfTextEditResult edited = document.Text.Replace(region, "Approved", new PdfTextEditOptions {
    Color = PdfColor.FromRgb(25, 110, 55),
    RegionWidthPolicy = PdfTextRegionWidthPolicy.ShrinkToFit,
    MinimumFontSize = 8
});

edited.Document.Text.Add(
        new PdfPageRegion(1, 72, 600, 240, 24),
        "Reviewed by Legal",
        new PdfTextEditOptions { Font = PdfStandardFont.HelveticaBold, FontSize = 11 })
    .Document
    .Save("contract-edited.pdf");
```

`Text.Find(...)` supports case and whole-word filters over visible, unclipped
text. A phrase can continue across consecutive visual lines in the same
paragraph, including a line-end hyphen; separate columns and explicit paragraph
breaks are not joined. `PdfTextMatch.Text` is the matched text after search
normalization, so whitespace may be collapsed and a line-end hyphen may be
retained as `-` or omitted. It is not an exact source-text excerpt.
`PdfTextMatch.VisualLineBounds` gives one quad per matched visual line for
highlighting; `VisualBounds` and `X`/`Y`/`Width`/`Height` remain the combined
bounds. `Text.ReplaceAll(...)` preserves unmatched source-span text and keeps
wide same-baseline runs such as columns independent. Edits fail closed
when an atomic PDF text object would require invisible or clipped text to be
recreated without its original rendering state.
Region-based add and replace operations preserve the resolved font size by
default. Select `RejectOverflow` to reject an over-wide line before mutation,
or `ShrinkToFit` to reduce it no lower than `MinimumFontSize`.
When `PdfTextEditOptions.Font` is null, safe horizontal runs reuse embedded
TrueType fonts and Identity-H CID fonts with readable ToUnicode maps. Replacement
text that fits the existing map keeps that font. A missing subset glyph uses the
detected standard substitute when its encoding covers the replacement, and adds
a warning. Set `Font` explicitly to choose another standard face. Text outside
both encodings is rejected before returning an edited document.
Mapped replacement text that requires shaping, bidirectional layout or
combining-mark positioning is also rejected by source-font reuse.
Other source fonts use the detected standard substitute and report a warning.
Clipping, unsupported positioning, shaping and ambiguous ActualText still block
edits whose rendering state cannot be recreated safely.
`PdfTextEditResult.Warnings` reports source-font substitutions that can change
metrics or letterforms.

`PdfReadPage.GetTextSpans()` exposes `PdfTextSpan.TryGetGlyphs(out glyphs)` for
immutable encoded character codes, UTF-16 text ranges, baseline origins and
advances. One glyph can map to several characters. The method returns false with
an empty list when evidence is unavailable or ambiguous, including ActualText
replacements. Baseline origins are not exact outlines or clipping bounds.

Use `Text.ReplaceSelected(find, replacement, matchIndexes, searchOptions, editOptions)`
to apply a reviewed subset of the zero-based occurrences returned by `Text.Find`
with the same query and search options. The document and search contract must
remain unchanged between discovery and mutation. `PdfTextMatch.SourceFont`,
`FontSize`, and `SuggestedFont` expose the detected style for a review interface.

Invisible OCR text stored with PDF text rendering mode 3 is opt-in for both
discovery and mutation. Use `IncludeTextRenderingMode3` to find it, then
`AllowTextRenderingMode3` to authorize an edit that preserves the invisible
rendering mode:

```csharp
PdfDocument scanned = PdfDocument.Load("scanned-contract.pdf");
var ocrSearch = new PdfTextSearchOptions {
    MatchCase = true,
    IncludeTextRenderingMode3 = true
};

PdfTextMatch ocrMatch = scanned.Text.Find("Account number", ocrSearch).Single();
PdfTextEditResult corrected = scanned.Text.Replace(
    ocrMatch,
    "Customer number",
    new PdfTextEditOptions { AllowTextRenderingMode3 = true });

corrected.Document.Save("scanned-contract-corrected.pdf");
```

The OCR opt-ins do not authorize clipping text modes or Type3 font glyph
programs, because those glyph programs can paint visible graphics independently
of the text rendering mode. One edit also cannot combine visible text and
rendering-mode-3 OCR text.

### Find and edit existing page images

Image placement coordinates also use PDF points from the page bottom-left.
Discover placements through the editor, then remove, replace, or move one exact
invocation without deleting overlapping text, paths, annotations, or unrelated
images:

```csharp
PdfDocument document = PdfDocument.Load("contract.pdf");
PdfImagePlacement logo = document.Images.Find(
    new PdfPageRegion(pageNumber: 1, x: 36, y: 720, width: 180, height: 60)).Single();

PdfImageEditResult updated = document.Images.Replace(
    logo,
    File.ReadAllBytes("new-logo.png"),
    new PdfImageEditOptions { Layer = PdfImageEditLayer.AboveExistingContent });

PdfImagePlacement replacement = updated.Document.Images.Find(
    new PdfPageRegion(1, 36, 720, 180, 60)).Single();

updated.Document.Images.Move(replacement, deltaX: 12, deltaY: -8)
    .Document
    .Save("contract-with-new-logo.pdf");
```

`Images.Add(...)` fits a new image to a page region. Replacement and movement
preserve position, size, and rotation when the source transform is portable;
callers explicitly choose whether rewritten content is above or behind existing
page content. The editor fails closed for ambiguous placements and for source
clipping, opacity, skew/reflection, unresolved transparency, image-mask, raw
payload, or inline-image semantics that cannot be reproduced safely. Exact
XObject removal remains available for rotated and skewed placements.

`Images.Transform(placement, deltaX, deltaY, scale, options)` translates an image's
center and scales both dimensions proportionally while retaining its rotation.
The placement must belong to the current document revision. As with movement,
the selected `PdfImageEditLayer` controls where the rewritten placement is painted.

### Fill and flatten a PDF form

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("application-form.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Applicant.Email"] = "adele@example.com",
        ["Approval.Status"] = "Approved"
    })
    .Save("application-form-filled.pdf");
```

Use `Forms.Fill` to retain editable fields, or `Forms.FillAndFlatten` to make the
values part of the page. Both return a new document. Pass
`new PdfFormFillerOptions { CancellationToken = token }` as the second argument
to cancel a job. Text and choice appearances share the same placement rules for
new fields, filled copies and appended revisions; flattening also retains inherited
left, center or right alignment when a source field has no saved appearance.

The [equipment inspection example](../OfficeIMO.Examples/Showcase/Workflows/EquipmentInspection.cs)
saves a blank template, a completed editable copy and a permanent record.

`Forms.FillResult` chooses the revision itself. When a full rewrite would break a
signature, it appends the values as a new revision after the original bytes. This
covers certified forms whose DocMDP level allows fill-in, and Reader-extended forms
such as IRS and other government forms whose only signature is a usage-rights
(`/UR3`) signature granting `/Form /FillIn`. Other changes to those forms stay
blocked. Existing JavaScript and other actions are kept unchanged, and the plan
reports `Input.ActiveContentPreserved`. On a hybrid AcroForm/XFA form the revision
fills the AcroForm fields and reports `Input.XfaDataNotSynchronized`, because
XFA-aware viewers such as Adobe Acrobat may keep showing the unchanged XFA data.
Set `XfaHandling` to remove the XFA packet in the appended revision so XFA-aware
viewers can use the AcroForm values. Removal can invalidate Adobe Reader usage
rights; preserving the original signed bytes does not establish viewer acceptance
of the changed form. Certified forms retain their XFA packet: removing it changes
the form definition and is refused under DocMDP.

```csharp
using OfficeIMO.Pdf;

PdfDocument form = PdfDocument.Load("fw9.pdf");
var values = new Dictionary<string, PdfFormFieldValue> {
    ["topmostSubform[0].Page1[0].f1_01[0]"] = PdfFormFieldValue.From("Adele Vance")
};
PdfMutationPlan plan = form.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
Console.WriteLine($"{plan.ExecutionMode}: {string.Join(", ", plan.Warnings)}");

form.Forms.AppendRevision(values, new PdfIncrementalFormFieldUpdateOptions {
    GenerateAppearanceStreams = true,
    KeepNeedAppearances = false,
    XfaHandling = PdfFormXfaHandling.Remove
}).Save("fw9-filled.pdf");
```

For an interactive editor, `PdfFormFieldValueAssessment.Assess(field, value)` checks
the proposed value against declared field metadata before applying it. It reports
read-only fields, scalar/multiple-value conflicts, text length, radio clearing,
and non-editable choices, including ambiguous display labels. It also assesses a
bounded inert numeric helper profile from field/widget actions: separator notation,
review precision and literal range. Unsupported or extra script code produces an error.
This policy requires a corrected value rather than simulating Acrobat formatting.
`PdfFormField.Actions` exposes actions on separate field dictionaries; widget actions
remain on `PdfFormWidget.Actions`. Empty required fields produce warnings so an unfinished
form can still be saved. Text limits count Unicode scalar values. Assessment does
not run JavaScript or replace document permissions, appearance generation, or
validation of the saved result; Unicode appearances still need suitable fonts.

To make both readable form fields and supported visual annotations static in one
artifact, use the combined operation and retain its separate affected counts:

```csharp
PdfInteractiveContentFlattenResult flattened = PdfDocument
    .Load("application-reviewed.pdf")
    .FlattenInteractiveContent();

Console.WriteLine($"Fields: {flattened.FlattenedFormFieldCount}");
Console.WriteLine($"Annotations: {flattened.FlattenedAnnotationCount}");
flattened.ToDocument().Save("application-static.pdf");
```

Annotation creation and updates accept gray, RGB, or CMYK interior colors,
opacity, border width, solid/dashed/beveled/inset/underline styles, and bounded
dash patterns. Creation generates supported appearances by default. Updates
remove a stale appearance when visual properties change and can explicitly
regenerate it with `RegenerateAppearance`.

### Language and table structure in tagged output

Create the document with `new PdfCreateOptions().EnableTaggedPdfCatalogMarkers()` and
set its default `Language`. Generated table headers have stable document-local
IDs, and data cells reference the applicable row and column headers through
`/Headers`. The ParentTree uses bounded number-tree nodes for larger documents.

For positioned text in another language, set the language on a semantic span:

```csharp
document.Content.Canvas(canvas => canvas.Structure(
    PdfCanvasStructureRole.Paragraph,
    paragraph => paragraph.Structure(
        PdfCanvasStructureRole.Span,
        span => span.Text(new[] { PdfTextRun.Normal("Bonjour") },
            PdfCanvasTextStructureRole.Span, 36, 36, 160, 24),
        new PdfCanvasStructureOptions { Language = "fr-FR" })));
```

Descendants inherit the span language. Set `Language = "en-US"` on a nested span
to return to the document language. PDF 2.0 structure elements declare the PDF 2.0
namespace, including the single Document child of the structure root.

Tagged native lists declare their label style. For custom canvas lists, set
`PdfCanvasStructureOptions.ListNumbering` on the `List` structure, for example
`PdfListNumbering.UpperRoman`. List labels contain inline `Span` text. The
`Ordered`, `Unordered` and `Description` styles require PDF 2.0.

For PDF/UA-2 canvas navigation, add named destinations and outlines inside the
`Structure` container that they target. The writer retains the page destination
and adds the matching structure destination; a target outside tagged content is
rejected when the strict PDF/UA-2 profile requires it.

### Generate and assess validator-backed PDF/A

```csharp
using OfficeIMO.Pdf;

var options = new PdfCreateOptions()
    .UsePdfA(PdfComplianceProfile.PdfA2B)
    .EmbedStandardFont(PdfStandardFont.Helvetica, "SourceSerif4-Regular.otf", "Source Serif 4");

PdfComplianceArtifact artifact = PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("This artifact is ready for external validation."))), options)
    .Meta(title: "Archive copy")
    .CreateComplianceArtifact(PdfComplianceProfile.PdfA2B);

byte[] pdf = artifact.ToBytes();
File.WriteAllBytes("archive.pdf", pdf);

// Create this result from the validator invocation in your build or release lane.
PdfExternalValidationResult validation = PdfExternalValidationResult.PassedForArtifact(
    PdfExternalValidatorKind.VeraPdf,
    "veraPDF",
    "1.30.2",
    "PDF/A-2b validation passed.",
    pdf,
    "PDF/A-2b");

PdfComplianceProofReport proof = artifact.AssessProof(new[] { validation });

PdfDeclaredComplianceClaimsReport declaredClaims = PdfDocument
    .Load(new MemoryStream(pdf))
    .AssessDeclaredComplianceClaims(new[] { validation });

if (!proof.CanClaimConformance || !declaredClaims.CanClaimAllDeclaredConformance) {
    throw new InvalidOperationException(proof.ExternalProofSummary);
}
```

Formal generation gates are available for PDF/A-2a/b/u, PDF/A-3a/b/u, PDF/A-4/4e/4f, PDF/UA-1, PDF/UA-2, PDF/X-1a:2003, PDF/X-4, Factur-X, and ZUGFeRD. `UsePdfA(...)` and `UsePdfUa(...)` enable strict generation and saved-artifact checks. When combined, both declared profiles must pass those checks. `RequireCompliance(...)` selects a strict gate without configuring its prerequisites. For metadata-only preparation, use `ConfigurePdfAGroundwork(...)` or `ConfigurePdfUaGroundwork(...)`. PDF/A-4 stores document metadata in XMP and omits the trailer Info dictionary. PDF/A-4f requires an emitted embedded file. Configured embedded files and emitted structure-associated files, including tagged Formula MathML supplements, count toward that requirement; supplements suppressed in untagged output do not. A conformance claim still requires a passing external result for the same profile, SHA-256, and byte length; validators are build-time tools and are not runtime dependencies of `OfficeIMO.Pdf`.

For an opened PDF, `pdf.AssessCompliance(profile, cancellationToken)` checks readback readiness and observes cancellation during inspection. These checks do not independently verify the PDF/A-4f catalog `EmbeddedFiles` name tree or all PDF/A-2 and base PDF/A-4 attachment restrictions. Attachment recovery alone does not prove those requirements; validate the exact saved artifact externally before claiming conformance.

PDF authoring accepts font faces up to 128 MiB each. `EmbedStandardFont` and `PdfEmbeddedFontFamily.FromFiles` check file size before reading and recheck while reading, so use their path overloads when loading files. Byte-array inputs are checked before OfficeIMO copies them. An oversized face throws `InvalidDataException` without changing the authoring options.

### Generate a fail-closed PDF/X artifact

```csharp
using OfficeIMO.Pdf;

byte[] cmykProfile = File.ReadAllBytes("PSOcoated_v3.icc");
byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");
var options = new PdfCreateOptions()
    .ConfigurePdfX(
        PdfComplianceProfile.PdfX4,
        cmykProfile,
        "FOGRA51",
        PdfTrappingStatus.False)
    .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4");

PdfComplianceArtifact artifact = PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("Generated colors are converted through the CMYK print condition."))),
        options)
    .Meta(title: "Print-ready report")
    .CreateComplianceArtifact(PdfComplianceProfile.PdfX4);

File.WriteAllBytes("print-ready.pdf", artifact.ToBytes());
```

`ConfigurePdfX` requires a caller-selected CMYK output-device profile because the correct profile depends on the printing condition and profile redistribution terms. For example, the [ICC registry entry for PSOcoated_v3](https://registry.color.org/profile-registry/PSOcoated_v3) identifies it as FOGRA51 and permits use, embedding, and exchange while restricting redistribution. OfficeIMO requires the ICC `prtr` device class, validates the header, declared size, CMYK component count, and a supported output transform, and writes a boolean trapping status (`False` by default). It also creates synchronized Info/XMP production dates, PDF/X identification, document and instance UUIDs, version identity, rendition class, and trapping metadata. It converts generated vector, text, and supported raster colors, applies the selected black-preservation policy, writes production page boxes, and inspects the exact PDF for remaining DeviceRGB content, embedded fonts, prohibited references, and profile-specific transparency. PDF/X-1a:2003 also rejects device-independent color spaces and flattens raster alpha before rejecting any remaining transparency. PDF/X-4 retains its standard color-management and transparency allowances, while OfficeIMO deliberately emits a conservative CMYK generated-content subset.

For reproducible builds, replace the generated timestamps and UUIDs with an explicit `PdfXProductionMetadata` value through `SetPdfXProductionMetadata(...)`. Reusing the same value produces stable production metadata; create a new document and instance UUID when the output represents a different resource.

Internal readiness is not a certification. Pass the exact artifact to a qualified PDF/X preflight tool and bind its result with `PdfExternalValidationResult.PassedForArtifact`; `PdfComplianceProofReport.CanClaimConformance` remains false when that exact external evidence is absent or mismatched.

For an imported PDF, inspect print concerns before submitting it to a press:

```csharp
PdfDocument incoming = PdfDocument.Load("incoming.pdf");
PdfProductionPreflightReport report = incoming.Proof.PreflightProduction(
    new PdfProductionPreflightOptions { Profile = PdfProductionPreflightProfile.PdfX4Candidate });
foreach (PdfProductionFinding finding in report.Findings)
    Console.WriteLine($"Page {finding.PageNumber?.ToString() ?? "document"}: {finding.Kind} - {finding.Message}");

foreach (PdfProductionFixupProposal proposal in report.FixupProposals)
    Console.WriteLine($"[{proposal.Index}] Page {proposal.PageNumber} {proposal.Box}: " +
        $"{proposal.Bounds.Left}, {proposal.Bounds.Bottom}, {proposal.Bounds.Right}, {proposal.Bounds.Top} — {proposal.Reason}");

// Enter only indices whose finished size and bleed coordinates you verified.
Console.Write("Approved proposal indices (comma-separated, blank to skip): ");
string? response = Console.ReadLine();
int[] approved = string.IsNullOrWhiteSpace(response)
    ? Array.Empty<int>()
    : response.Split(',').Select(value => int.Parse(value.Trim())).ToArray();
if (approved.Length > 0) {
    PdfProductionFixupResult result = report.ApplySelected(approved);
    result.Document.Save("boxes-reviewed.pdf");
    Console.WriteLine($"Remaining findings: {result.After.Findings.Count}");
}
```

The general-print and PDF/X candidate profiles inspect output intents, page
boxes, reachable font programs, color and transparency, and placed-image
resolution. Findings link to pages and image bounds when the evidence permits.
Printable annotations and reachable tiling patterns currently produce
indeterminate color or image-resolution findings where their painted content
cannot be measured by preflight.
Page-box fixups change metadata only; they never extend artwork or supply an
ICC profile. Each accepted fixup is applied to the inspected PDF snapshot and
the result is reopened for a new engine inspection. A qualified external
validator is still required for a PDF/X conformance claim.

### Supply document font families

Applications can supply a font family independently of the host's installed fonts. Provide every face needed by the document, then register any source-family replacement explicitly:

```csharp
using OfficeIMO.Pdf;

var fonts = new PdfOptions().RegisterNamedFontFamily(
    PdfEmbeddedFontFamily.FromFiles("Carlito",
        "Fonts/Carlito-Regular.ttf",
        "Fonts/Carlito-Bold.ttf",
        "Fonts/Carlito-Italic.ttf",
        "Fonts/Carlito-BoldItalic.ttf"));

fonts.RegisterFontFamilySubstitution("Calibri", "Carlito",
    PdfFontFamilySubstitutionImpact.LayoutSensitive);
```

Pass `fonts` through `WordToPdfOptions.PdfOptions` or `RtfToPdfOptions.PdfOptions`. Registration makes a family available; a substitution selects it for the named source family. Supply the same files and mappings on each host to make that selection independent of system-font discovery. Other document families still need their own registrations or the normal fallback policy.

These files are supplied by the application and are not bundled with the PDF package. Use fonts whose licensing permits the intended embedding and distribution. `LayoutSensitive` records that replacement can affect layout. `Compatible` expresses the caller's compatibility intent; registration does not verify identical glyph advances, line metrics or pagination for the source and replacement font versions. Inspect the conversion report and representative output when choosing a substitute.

### Choose converter-friendly text fallbacks

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("proposal.docx");

var options = new WordPdfExportOptions {
    TextFallbacks = PdfTextFallbackFeatures.Default,
    ResourcePolicy = PdfResourcePolicy.CreateTrustedHost()
}.UseProfile(PdfExportProfile.PrintReady);

OfficeIMO.OfficeConversionReport report = document.SaveAsPdf("proposal.pdf", options);
report.ThrowIfLoss();
```

The Word, Excel, PowerPoint, Markdown, HTML, RTF, OneNote, AsciiDoc, and LaTeX PDF adapters use one `PdfResourcePolicy`; semantic-projection adapters expose it through their nested Markdown PDF options. The balanced default enables installed fonts and bounded data URI/package resources for document fidelity while denying arbitrary local files and remote resolver calls. For Word, Excel and PowerPoint this includes embedding installed document-selected fonts, with installed substitutes preferred when a requested Office family is missing. `font-family-substitution` diagnostics identify replacements, and `font-embedding-cost` records actual font-program stream sizes, compression and subset/full-program status. These byte counts exclude font dictionaries and character maps. Use `PdfResourcePolicy.CreatePortableDeterministic()` for reproducible or untrusted conversion, and `CreateTrustedHost()` only when both source and host are trusted. Profiles never grant resource access.

Use `PdfResourcePolicy.CreateLibraryFontDefault()` to allow library-selected installed families while disabling host fonts selected by the source document. HTML uses this policy by default; callers can explicitly enable `AllowDocumentFontEmbedding` when needed.

The text-capable adapters also expose `TextFallbacks`. `PdfTextFallbackFeatures.Default` enables document, monospace, symbol, emoji, and multilingual groups, including CJK, Arabic, and other non-Latin family candidates; OneNote adds that candidate group unless fallbacks are `None`. Candidate selection does not read installed fonts unless the resource policy allows it.

On Windows, local font discovery includes fonts Office has already cached in its cloud-font folders. Numeric filenames are resolved through family folders and font metadata. The selected fonts follow the same resource and embedding rules.

`PdfCreateOptions.RegisterEmbeddedFontFallbacks` preserves caller-registered embedded fonts in compatibility slots. It fills vacant slots and chooses an available replacement when a fallback segment needs an occupied slot. Replacing the fallback set refreshes slots still owned by the previous set while preserving later caller registrations. It also preserves caller-owned named families, including their bold and italic faces, when a fallback candidate requests the same normalized family name. The fallback uses a separate resource, and runs from the reusable fallback plan resolve that resource through the current `PdfCreateOptions`; callers do not need to rewrite the planned run's family name. A later explicit family registration retains caller ownership. Named fallback plans and explicit named-run mappings preserve literal resource-family names containing commas or backslashes instead of interpreting those names as fallback lists. Unambiguous ordinary names keep the same public run-family value.

`PdfEmbeddedFontFallbackSet.RegisterFonts` installs the supplied slot and named-family mappings directly and can replace those mappings. Coverage scans inspect candidates in priority order; later candidates remain unparsed when an earlier candidate covers the text. A candidate that must be inspected still needs valid font data. Missing-font substitution can change glyph metrics and line breaks even when the substitute covers every character.

### Inspect and remove content provenance

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;

OfficeProvenanceReport report = PdfProvenance.InspectFile("input.pdf");
OfficeProvenanceRemovalResult result = PdfProvenance.RemoveFile("input.pdf", "clean.pdf");
```

The PDF owner recognizes the standards-defined embedded file with media type `application/c2pa` and `/AFRelationship /C2PA_Manifest`. Removal uses a bounded, targeted object-graph rewrite that leaves unrelated attachment associations in place. Malformed indirect-object candidates remain untouched unless the caller disables `RequireStructurallyValidCarrier`; direct file-spec dictionaries are inspected but cannot be removed safely. A signed PDF is never silently rewritten or stripped; handle its signature through an explicit PDF signature workflow first. Optional cryptographic C2PA verification is provided by `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`PdfDocument.InspectContentSafety(...)` reports non-painting text render modes, effective transparency, clipping, tiny/zero/off-canvas geometry, paint-order-resolved low contrast, hidden annotation and form-widget values, text inside layers hidden by the default optional-content configuration, and Unicode evidence from decoded spans. `PdfDocument.RemoveSelectedContent(...)` physically removes exact selected spans and can rewrite reviewed Unicode ranges in ordinary painted or painted low-contrast spans while verifying neighboring text restoration. Hidden layer, annotation, and widget findings remain report-only because their container semantics cannot be safely converted into an exact text edit. Encrypted and signed PDFs are rejected.

### Generate a formal e-invoice carrier

```csharp
using OfficeIMO.Pdf;

byte[] invoiceXml = File.ReadAllBytes("factur-x.xml");
DateTimeOffset invoiceModifiedAt = File.GetLastWriteTimeUtc("factur-x.xml");
byte[] fontBytes = File.ReadAllBytes("SourceSerif4-Regular.otf");

PdfDocument.Create(pdf => pdf.Content(content => content
        .Paragraph(paragraph => paragraph.Text("Invoice preview"))),
    new PdfCreateOptions()
        .UseFacturX(
            invoiceXml,
            relationship: PdfAssociatedFileRelationship.Alternative,
            textFallbacks: PdfTextFallbackFeatures.None)
        .SetEmbeddedFileModificationDate("factur-x.xml", invoiceModifiedAt)
        .EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "Source Serif 4")
        .RequireCompliance(PdfComplianceProfile.FacturX))
    .Save("invoice.pdf");
```

Factur-X/ZUGFeRD hybrid containers support only PDF/A-3a, PDF/A-3b and PDF/A-3u, optionally with PDF/UA-1. PDF/A-4, PDF/A-4e and PDF/A-4f do not satisfy the invoice compliance profile; their general PDF generation and attachment capabilities follow their own profile requirements.

The XML must be a valid CrossIndustryInvoice payload for its declared MINIMUM, BASIC WL, BASIC, EN 16931 or EXTENDED profile. The formal carrier gate checks the PDF/A-3 attachment, metadata, font, Unicode and profile-specific invoice rules before writing; exact-artifact PDF/A and invoice-validator results are still required before claiming conformance.

#### Load and edit existing invoice XML

`PdfCiiInvoiceDocument` is part of `OfficeIMO.Pdf`. It reads existing namespace-100
CII invoice XML and lets you replace an existing identifier or format-102 issue
date while retaining unfamiliar XML content.

```csharp
using OfficeIMO.Pdf;

var invoice = PdfCiiInvoiceDocument.Load("invoice.xml")
    .WithDocumentId("INV-2026-0042")
    .WithIssueDate(new DateTime(2026, 9, 6));

invoice.Save("updated-invoice.xml");
var options = new PdfCreateOptions().UseFacturXDocument(invoice);
```

`Load` also accepts a byte array or readable stream. `Save` accepts a path or
stream; caller-owned streams remain open. `ToBytes()` returns a defensive copy.
An unedited document keeps its original bytes. Edits return a new document with
UTF-8 XML, preserving XML content but not its original lexical formatting.

The model exposes `DocumentId`, `IssueDate`, `TypeCode`, `GuidelineId`, and
`CurrencyCode`. It rejects DTDs, excessive input size or depth, and edits to
signed, missing, ambiguous, or structured fields. Loading is not schema or
business-rule validation. This API does not create complete invoices, calculate
tax, edit PDF pages, or update payment references.

`UseFacturXDocument` snapshots the XML through the existing `UseFacturX(byte[])`
carrier configuration. Omit `conformanceLevel` to derive canonical XMP metadata
from the XML guideline. An explicit conflicting profile, an unknown guideline,
or an ambiguous declaration is rejected. Generation and exact PDF readback also
check XML/XMP agreement when metadata and attachments are supplied separately.
The low-level Factur-X/ZUGFeRD carrier methods accept at most 16 MiB of invoice
XML, matching `PdfCiiInvoiceDocument.MaximumXmlBytes`. File overloads check
this limit before buffering. Split or reduce a larger XML payload before
attaching it; this limit does not establish schema or business-rule validity.
The XMP `version` is `1.0`, not a Factur-X release number.
The application must keep visible invoice content and
XML consistent and validate the resulting invoice/PDF pair for its declared profile.

### Generate a visible invoice and its XML from one snapshot

Use the optional [OfficeIMO.Invoicing.Pdf adapter](../OfficeIMO.Invoicing.Pdf/README.md)
to render a typed invoice and embed its captured CII XML. The adapter owns
`PdfInvoiceDocument`; `OfficeIMO.Pdf` depends only on `OfficeIMO.Core` and keeps
its existing low-level XML carrier and bounded `PdfCiiInvoiceDocument` APIs.

### Page setup, watermarks, and metadata

```csharp
PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Draft report")
        .Paragraph(paragraph => paragraph.Text("This document uses page-level options instead of post-processing."))),
    new PdfCreateOptions {
            PageSize = PdfPageSize.FromCentimeters(21, 29.7).Portrait(),
            Margins = PdfPageMargins.UniformCentimeters(1.5),
            TextWatermark = new PdfTextWatermark("DRAFT") {
                Opacity = 0.12,
                RotationAngle = -35
            }
        })
    .Meta(title: "Draft report", author: "OfficeIMO")
    .Save("draft.pdf");
```

For facing-page layouts, set `PdfCreateOptions.MirrorMargins = true` or use
`page.MirrorMargins()` in a section. The authored left and right margins apply to
odd visible page numbers and swap on even numbers. `PageNumberStart` restarts
that parity. Flow content follows the current page frame, including automatic
continuation through paragraphs, tables, lists, columns, and padded elements.
Canvas coordinates remain absolute.

Use `page.StartOnPageParity(PdfPageParity.Odd)` or `Even` to pad a later section
with a blank page when its physical page index has the wrong parity. Pass
`useContinuingPageNumber: true` to choose parity from the preceding section's
continuing page number; the new section's `PageNumberStart` applies after padding.
The equivalent options are `PageStartParity` and
`UseContinuingPageNumberForStartParity`. The first section starts without padding.
`content.PageBreak(preserveEmptyPage: true)` preserves an empty page when explicit
breaks are consecutive. The parameterless overload skips empty pages.

### Inspect and preflight before rewriting

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("incoming.pdf");
PdfDocumentPreflight preflight = pdf.Analyze().Preflight;

if (!preflight.Can(PdfPreflightCapability.ManipulatePages)) {
    foreach (string diagnostic in preflight.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages)) {
        Console.WriteLine(diagnostic);
    }
}

using PdfDocument result = pdf.Pages.Extract("1-2");
result.Save("incoming-first-pages.pdf");
```

### Inspect before automating

```csharp
PdfDocument pdf = PdfDocument.Load("incoming.pdf");

var inspection = pdf.Inspect();
Console.WriteLine($"Pages: {inspection.PageCount}");
Console.WriteLine($"Links: {inspection.LinkAnnotationCount}");
Console.WriteLine($"Forms: {inspection.FormFields.Count}");
Console.WriteLine($"Active content: {inspection.HasActiveContent}");

foreach (var page in inspection.Pages) {
    Console.WriteLine($"{page.PageNumber}: {page.Width} x {page.Height}");
}

PdfMutationPortfolioReport mutations = pdf.AssessMutations();
PdfRenderCompatibilityReport rendering = pdf.AssessRenderCompatibility();
Console.WriteLine($"Executable mutation families: {mutations.ExecutablePlans.Count}");
Console.WriteLine($"Render capability findings: {rendering.DiagnosticCount}");

foreach (PdfMutationPlan blocked in mutations.BlockedPlans) {
    Console.WriteLine($"{blocked.Operation}: {string.Join(", ", blocked.BlockerCodes)}");
}
```

Pass your job's cancellation token to `pdf.AssessMutations(cancellationToken)`
when assessing large documents or a document collection. The assessment checks
cancellation while collecting requests, during parsing and preflight, and between
individual plans. Each plan retains its own blocker codes; a blocked page edit does
not mean that metadata updates, form filling, or signing are also blocked.

### Convert PDFs through adapter packages

```csharp
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var word = WordDocument.Load("proposal.docx");
word.SaveAsPdf("proposal.pdf");

PdfDocument statement = PdfDocument.Load("bank-statement.pdf");
OfficeIMO.OfficeConversionReport tableReport = statement.SaveAsExcel("bank-statement-tables.xlsx");
Console.WriteLine($"Table reconstruction has loss: {tableReport.HasLoss}");

using var proposal = PdfDocument.Load("proposal.pdf");
proposal.SaveAsHtml("proposal-review.html", new PdfToHtmlOptions {
    Profile = PdfHtmlProfile.PositionedReview,
    IncludeLinkAnnotations = true,
    IncludeFormWidgets = true
});
```

## Conversion adapters

| Package | Role |
| --- | --- |
| [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md) | Maps Word documents into PDF primitives. |
| [OfficeIMO.Excel.Pdf](../OfficeIMO.Excel.Pdf/README.md) | Exports Excel workbooks as PDF. |
| [OfficeIMO.Markdown.Pdf](../OfficeIMO.Markdown.Pdf/README.md) | Exports Markdown documents as PDF. |
| [OfficeIMO.PowerPoint.Pdf](../OfficeIMO.PowerPoint.Pdf/README.md) | Maps PowerPoint slides to PDF and imports PDF pages as visual slides or editable detected tables. |
| [OfficeIMO.Html.Pdf](../OfficeIMO.Html.Pdf/README.md) | Bridges HTML to PDF and PDF to HTML. |
| [OfficeIMO.Rtf.Pdf](../OfficeIMO.Rtf.Pdf/README.md) | Maps semantic RTF into PDF and logical PDF content back to RTF. |
| [OfficeIMO.OneNote.Pdf](../OfficeIMO.OneNote.Pdf/README.md) | Explicitly projects offline OneNote hierarchy into a semantic PDF document with loss diagnostics. |
| [OfficeIMO.AsciiDoc.Pdf](../OfficeIMO.AsciiDoc.Pdf/README.md) | Projects native AsciiDoc through the loss-aware Markdown bridge and combines parser, projection, and PDF diagnostics. |
| [OfficeIMO.Latex.Pdf](../OfficeIMO.Latex.Pdf/README.md) | Projects the bounded LaTeX profile through the loss-aware Markdown bridge without executing TeX. |
| [OfficeIMO.OpenDocument.Odt.Pdf](../OfficeIMO.OpenDocument.Odt.Pdf/README.md) | Provides direct ODT and PDF façades with combined OpenDocument and PDF diagnostics. |
| [OfficeIMO.OpenDocument.Ods.Pdf](../OfficeIMO.OpenDocument.Ods.Pdf/README.md) | Provides direct ODS and PDF façades without Word or PowerPoint dependencies. |
| [OfficeIMO.OpenDocument.Odp.Pdf](../OfficeIMO.OpenDocument.Odp.Pdf/README.md) | Provides direct ODP and PDF façades without Word or Excel dependencies. |

The generated [PDF conversion support matrix](../Docs/officeimo.pdf-conversion-support-matrix.md) records current direct and composed routes from the canonical [`Docs/pdf-conversion-scenarios.json`](../Docs/pdf-conversion-scenarios.json) manifest. `OfficeIMO.Pdf` projects the dependency-free `OfficeDocumentModel` through `PdfProjectionOptions`; `OfficeIMO.Reader.Pdf` keeps a thin compatibility bridge for existing `OfficeDocumentReadResult` workflows. Email, EPUB, and Visio are advertised as direct conversion only when their route-specific artifact gates are proven. Open PDF work is tracked in the repository [roadmap](../Docs/ROADMAP.md).

## Related packages and docs

- `OfficeIMO.Pdf` provides first-party PDF parsing, layout, writing, rendering, password security, and signature structure. Optional CMS, DER, and X.509 services come from an explicitly supplied `OfficeIMO.Security` provider.
- Source-format adapters map their document models onto the neutral `OfficeDocumentModel`; PDF projection remains owned by this package.
- Explicit `Location.LogicalOrder` takes precedence over local paragraph indexes. With `PdfProjectionPagePolicy.ContinuousFlow`, it also takes precedence over physical page order. Automatic page headings are omitted with `MODEL_SOURCE_PAGE_LABELS_OMITTED` when they would interrupt that order.
- Page and document table views are reconciled by source scope and occurrence count. Distinct equal tables remain in the output. An unlocated aggregate that matches several page scopes is retained with `MODEL_TABLE_CORRELATION_UNASSESSED` rather than assigned to an inferred page.
- Neutral-model raster assets retain their aspect ratio and scale down to the configured page content area. Source diagnostics keep their severity and an explicit `lossKind` attribute (`None`, `Approximation`, `Omission` or `Failure`) in the PDF report; absent or invalid categories use the severity default, and errors remain failures.
- See the [PDF current-state guide](../Docs/officeimo.pdf.current-state.md) for the detailed capability inventory and known limits.

## Repository validation

The repository keeps the public contract, target frameworks, package dependency
shape, performance budgets, compliance proof, and rendered output under
separate gates:

```powershell
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net8.0
dotnet test OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj -c Release -f net10.0
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net8.0 -- --verify-budgets
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0 -- --verify-budgets
dotnet run --project OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0 -- --verify-timing-budgets
Build/Test-PdfQualityCorpus.ps1 -Configuration Release -Framework net8.0
Build/Test-RealWorldCorpusContract.ps1
Build/Export-PdfComplianceProof.ps1 -Configuration Release -Framework net8.0
Build/Export-PdfVisualReviewGallery.ps1 -Configuration Release -Framework net8.0
```

The strict PDF quality scorecard runs every public inspection and semantic stage,
the complete mutation portfolio, managed rendering, and declared
compliance claim gating against hash-pinned Open Preservation Foundation and
veraPDF fixtures. The real-world corpus lane applies the same deep PDF stages to
a deterministic sample of a checksum-pinned public GovDocs archive in isolated
processes. The opt-in PDF Performance Evidence workflow
uses a deterministic 60-page mixed corpus and checks cold and cached analysis,
SVG rendering, PNG rendering, output integrity, absolute allocation and heap
budgets, generous elapsed-time ceilings, and cached allocation savings. Relative
cached speedup is opt-in through `--verify-timing-budgets` for controlled
benchmark hosts. Ordinary correctness CI runs the PDF tests without these
host-dependent performance budgets. The budget runner stays outside the normal
solution and can also be run directly with the commands above.

Pixel baselines are strict when the installed Poppler major/minor version
matches the recorded renderer. A different renderer version still runs semantic
and page-count checks in ordinary local runs. Required-rasterizer and CI visual
gates fail on a version mismatch; release investigations can deliberately opt
into a cross-version comparison.

## Contents for authored bookmarks

Use `PdfTableOfContentsOptions.Entries` when a document already has bookmarks.
Entries keep their authored order and hierarchy; the PDF layout engine resolves
their page numbers after pagination, including content that wraps onto more pages.
Omit `Entries` to keep the generated section-based contents.

```csharp
byte[] bytes = PdfDocument.Create()
    .Compose(document => document.Content(content => content
    .TableOfContents(new PdfTableOfContentsOptions {
        Entries = new[] { new PdfTableOfContentsEntry("Findings", "findings") }
    })
    .PageBreak()
    .Bookmark("findings")
    .H1("Findings")
    .Paragraph(paragraph => paragraph.Text("Evidence and recommendations."))))
    .ToBytes();
```

## Current state

The PDF engine is useful and broad, but it is still evolving. It has strong first-party coverage for common generated business documents, reusable Unicode line breaking and Latin ligatures, a bounded managed static-font profile for Arabic joining/required ligatures and Arabic/Hebrew/Latin mark positioning, plus an optional HarfBuzz adapter for broader GSUB/GPOS shaping, and bounded Type 3 rendering within the documented capability contract. It also supports authored and bounded-synthesized annotation appearances in page images, conservative read/manipulation workflows, password security, optional provider-backed certificate signing/validation, standards-compliant Fast Web View output, and bounded-payload stream saves with runtime serialization evidence. See the [image export capability matrix](../Docs/officeimo.image-export-capability-matrix.md) for the exact Type 3 rendering coverage and current limitations.

Vertical `OfficeDrawing` text retains its logical string through PDF `/ActualText`, but painting still uses stacked horizontal glyphs and reports `vertical-text-stacked-fallback` as an approximation. Strict no-loss profiles reject that result. Native vertical glyph output remains open until shaped vertical advances and substitutions, readable logical text, and strict-profile results pass together against independent CJK producer references.

For the current capability inventory, ownership boundaries, premium conversion
contract, and remaining general engine work, read
[Docs/officeimo.pdf.current-state.md](../Docs/officeimo.pdf.current-state.md).

## Choose a deployment path

Install only the packages your application needs:

| What you need | Package or provider | What to deploy |
| --- | --- | --- |
| Create, read, fill forms and organize PDFs | `OfficeIMO.Pdf` | Managed .NET libraries; no Office installation or browser |
| Render saved HTML templates | `OfficeIMO.Html` | Managed renderer and its HTML/CSS parser packages; provide fonts for your document's languages |
| Full OpenType text shaping | Optional `OfficeIMO.Drawing.HarfBuzz` | The provider's native files for your operating system and architecture |
| Capture a live website that uses JavaScript | Optional `OfficeIMO.Html.Pdf.Browser` | HtmlTinkerX and a supported Chromium installation |
| Read scanned text | Optional `OfficeIMO.Pdf.Ocr` and an OCR provider | The selected provider and its language resources; Tesseract uses a separate process |
| Sign or check certificate-based signatures | Optional `OfficeIMO.Security` | The selected security provider and the signing or trust configuration your application requires |

NativeAOT, browser and container checks apply to the particular packages and
workflows exercised. They do not guarantee every optional provider works in every
host. See the [NativeAOT scenarios](../Build/Test-AotScenarios.ps1),
[HTML deployment guide](../OfficeIMO.Html/README.md), and
[PDF capability guide](../Docs/officeimo.pdf.current-state.md) for the tested routes
and remaining limits. Browser rendering is an explicit choice; the managed PDF
and template paths do not download or start a browser automatically.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** no third-party PDF parser, writer, renderer, or cryptographic package in the base PDF package.
- **OfficeIMO:** `OfficeIMO.Core`. PDF parsing, writing, password security, signature structure, logical recovery, manipulation, forms, diagnostics, and preservation analysis are first-party.
- **Optional security:** install `OfficeIMO.Security` for the built-in CMS/X.509/RFC 3161 adapters. It is not a transitive PDF dependency.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 0 |
| Edit | 2 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 3 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 0 | 0 | 0 | 0 | 1 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
