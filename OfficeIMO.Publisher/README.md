# OfficeIMO.Publisher

Read Microsoft Publisher publications into positioned pages, styled text
stories and embedded image assets, edit supported stories, frames, tables, pictures and
vector artwork, or create an in-memory publication with pages, masters, styled
text, embedded pictures, tables and primitive shapes. Render the same current model to SVG or PDF.
Save an unchanged native source exactly, or save the supported source edits with
an explicit fidelity policy and an operation report.
The managed reader handles the Publisher
2002-and-later compound-document generation without installing Publisher.
[The support contract](SUPPORT.md) identifies the qualified fixtures and rendering
limits.

`OfficeIMO.Publisher` references only `OfficeIMO.Core`. Native records belong to
this package; compound storage, OfficeArt properties and image envelopes, text
layout, and SVG rendering use shared OfficeIMO engines.

## Read and inspect a publication

Text runs retain native font names, size, referenced color, bold/italic,
strikethrough, baseline and supported underline patterns. Native underline weight,
long-dash and compound-pattern differences carry explicit approximation reports.
Font availability and native layout fidelity remain separate rendering concerns.

```csharp
using OfficeIMO.Publisher;

var publication = PublisherDocument.Load("newsletter.pub");

foreach (var page in publication.Pages)
    Console.WriteLine($"Page {page.Id}: {page.Width} x {page.Height} points");

foreach (var story in publication.TextStories)
    Console.WriteLine(story.Text);

foreach (var diagnostic in publication.ReadReport.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

`Pages` follows publication order and excludes utility pages and master
definitions. `MasterPages` exposes the recovered definitions; their content is
also applied to referring pages. Coordinates use points and a top-left origin.
`CreateDrawing()` returns a detached shared `OfficeDrawing` snapshot, including
the applied master. Adding artwork to that copy affects exports of the copy;
document exports use the publication model. Use the document's typed edit
operations to change its stories and text frames.

## Save a native publication

```csharp
using OfficeIMO;
using OfficeIMO.Publisher;

var publication = PublisherDocument.Load("flyer.pub");
PublisherSaveReport copyReport = publication.Save("flyer-copy.pub");

// After supported edits, accept and inspect the operation's unassessed fidelity.
var settings = new PublisherSaveOptions {
    LossPolicy = OfficeConversionLossPolicy.Allow,
    MaxOutputBytes = 64L * 1024 * 1024
};
PublisherSaveReport report = await publication.SaveAsync("flyer-edited.pub", settings);
foreach (var diagnostic in report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");

var result = publication.ToPubBytesResult(settings);
byte[] detachedBytes = result.Value;
```

`ToPubBytes`, `ToPubBytesResult`, `Save` and `SaveAsync` use the same bounded native
codec. Unchanged semantic content produces a byte-exact copy of the original
source. Supported edits include existing ordinary frame/shape/picture bounds,
existing solid scalar paint slots, qualified frame columns/full intercolumn gap,
insets and vertical alignment, and ordinary text stories with bounded
size, bold, italic, underline pattern, single strikethrough, baseline, font family
and opaque RGB color changes. Paragraph counts can change when the source paragraphs have
uniform projected formatting and matching native definitions. Nonempty native
paragraphs with different formatting can also be retained, deleted, reordered or
duplicated without changing their content or formatting. Pass the original
paragraph snapshots to identify the exact native source; copied values require
unambiguous native definitions. A changed story
can use one existing frame or a complete resolved chain of ordinary, visible,
untransformed document-page frames. Native links, ownership and frame identities
remain fixed. Known chained Quill layouts are supported; formatting and directory
pages can be allocated or retired, and font/color resources can grow while existing
entries retain their identities and bytes.

For example, removing the first paragraph while retaining the others keeps their
native formatting definitions:

```csharp
var story = publication.TextStories.First(item => item.Paragraphs.Count >= 2);
publication.ReplaceStoryParagraphs(story.Id, story.Paragraphs.Skip(1).ToArray());
```

Frame-layout saves retain links and page/master ownership. They can add missing
default fields and relocate native records; changed distances must be exact
nonnegative EMU values. This API's gap is the full distance between columns.
Existing qualified pictures can save a replacement PNG or JPEG. Replacing one
placement retains the original asset and leaves other placements sharing it
unchanged, including master pictures. Native image slots and filename resources
are allocated together; asset IDs are document-scoped and can change on reload.
Existing qualified tables can save cell text and supported run formatting while
retaining the grid, cell ownership and each cell's paragraph count. Cell text
endpoints are remapped with the complete story; preceding ordinary stories can
also change length without detaching table cells.
Retained empty paragraphs can save source typing or an explicit empty styled
run. A paragraph with no runs retains its native CR style; native reload exposes
one styled empty run. Uniform-source paragraph growth can end in an empty
paragraph. Removing every paragraph from a native story remains unsupported.
Paragraph edits include alignment, before/after spacing, margins, first-line or
hanging indents, explicit absolute/multiple line spacing, and ordinary left tabs.
For native center/right paragraphs, use tab settings with
`WithParagraphAlignment()`; the tab origin must remain `-paragraph.Margins.Left`
and the default interval 36 points. Use `lineHeightFactor: 1` for explicit single
spacing. The support contract defines native units and representability limits.
Ambiguous source definitions, unsupported directory layouts and zero-paragraph
native story states are refused. The [support contract](SUPPORT.md)
defines the precise restrictions. Other in-memory edits, including new objects,
table bounds/grid changes and page/master relationships, cannot be saved
as native PUB through this profile. Source-free native creation is unavailable.

The default loss policy is `Block`, including unassessed fidelity. Edited native
output requires explicit `Allow`; this policy still refuses unsupported edits and
unsafe source carriers. `OfficeConversionException.Report` carries a
`PublisherSaveReport` for policy or preparation refusal. Its diagnostics describe
this save operation; `SourceReadReport` separately retains the original decoding
evidence. `IsExactSourceCopy` does not reassess native validity, signatures or
repair. Saving leaves the original source baseline unchanged.

File destinations require a `.pub` extension and default to `FailIfExists`.
Select `FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace` for
atomic replacement. The complete output is prepared and its policy checked before
publication; preparation, limits and pre-publication cancellation preserve an
existing destination. An unavailable atomic replacement fails explicitly.

Stream overloads leave the caller's stream open. Seekable streams are replaced,
truncated and rewound on success; non-seekable streams receive output at their
current position. Preparation failures preserve existing bytes and position.
Cancellation or I/O failure during stream publication can leave partial output.
Asynchronous saves prepare the native bytes in process before asynchronous I/O;
synchronous stream saves observe cancellation before publication. The default
output ceiling is 64 MiB and can be set from 1 through `int.MaxValue` bytes.

## Create pages, masters and styled text

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Publisher;

var newsletter = PublisherDocument.Create(792, 612);
var firstPage = newsletter.Pages[0];
var header = newsletter.AddMasterPage("Newsletter header");
newsletter.AddTextFrame(header.Id, 36, 18, 600, 30, new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Weekly publication", 16, OfficeColor.Black,
            bold: true, fontFamily: "Arial")
    })
});
newsletter.SetPageMaster(firstPage.Id, header.Id);
var body = newsletter.AddTextFrame(firstPage.Id, 36, 84, 300, 160, new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("A complete styled story follows its linked frames.",
            12, OfficeColor.Black, fontFamily: "Arial")
    })
});
var secondPage = newsletter.AddPage(header.Id);
newsletter.AddLinkedTextFrame(body.Id, secondPage.Id, 36, 84, 720, 420);
File.WriteAllText("newsletter-page.svg", newsletter.ToSvg());
```

`Create` takes a positive finite width and height in points and creates one empty
document page without a master or native source. Its optional `PublisherReadOptions`
is copied and bounds authoring and projection as well as native reading.
`AddPage` appends at the existing publication size; omitting its master ID creates
a page without a master. `AddMasterPage` creates a centered single-page definition
with a nonempty name. `SetPageMaster(pageId, null)` clears an assignment. Direct
qualified imported masters, including separately recovered left and right members,
retain their exact registration. Master content remains owned once in `MasterPages`
and renders on each assigned physical page. Nested masters, unassessed backgrounds
and unsupported registrations reject assignment. Page resizing, deletion, reordering,
object moves and paired-master creation are not provided.

`AddTextFrame` creates a complete independent ordinary story and a frame on a
qualified document or master page. It uses one column, zero gap and padding, top
alignment, no fill or outline, and no transform or wrap references. Supplied
paragraphs retain their Core formatting, labels and tabs. Empty paragraph
collections are allowed. Carriage returns normalize to newlines; paragraph
boundaries add newlines, with no implicit final terminator for authored ordinary
stories. `AddLinkedTextFrame` continues an ordinary document-page story; master
stories cannot be linked across physical pages. All story content remains available
in `TextStories`, including overflow.

Pages, masters, objects and stories share one publication identifier allocator;
creating a frame/story pair reserves both IDs atomically. Assets retain their own
image IDs. Failed validation, limits, cancellation or projection leave the document
and allocator unchanged. Returned objects are immutable snapshots; fetch current
objects after later edits. `HasNativeSource` distinguishes a loaded document from
source-free creation. `PublisherPage.HasNativeOrigin` identifies imported page
definitions even after content or master-assignment edits; added pages and masters
have no native origin. A source-free document has zero native source/record counts.
Source-free creation supports in-memory authoring and rendering. Native saving
requires a loaded publication and the supported source-edit profile below.

Add an embedded picture to a qualified document or master page:

```csharp
var illustration = newsletter.AddPicture(firstPage.Id,
    File.ReadAllBytes("illustration.png"), 360, 84, 144, 96);
newsletter.SetPictureBounds(illustration.Id, 360, 84, 120, 80);
```

`AddPicture` accepts a static PNG or JPEG with normal encoded orientation,
validated through the shared managed image decoder. Animated PNGs, non-normal
EXIF orientations, other formats and invalid payloads are rejected. The supplied
point rectangle controls placement and size; image DPI does not choose the frame.
The new picture has no crop, fill, outline, effects, story or wrap references and
is appended in drawing order. Pictures added to a master stay owned once and
render on every assigned page. Creation does not add text-wrap exclusions.

Input bytes are copied and retained in their original encoding. Each call creates
one placement and a separate asset, even for identical input. Object and image IDs
are reserved together and neither is consumed on failure. Existing image replacement
and movement operations remain available; replacement retains its existing PNG/JPEG
contract. Per-image, current-object, image-store, retained-byte and cumulative scene
budgets apply, including repeated master use. `GetBytes()` and old page/picture/asset
references remain detached from later edits.

## Edit publication content

For a publication with an ordinary, resolved, untransformed top-level text frame:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Publisher;

var publication = PublisherDocument.Load("flyer.pub");
var frame = publication.Pages[0].TextFrames[0];

publication.ReplaceStoryParagraphs(frame.StoryId, new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Updated publication", 18, OfficeColor.Black,
            bold: true, fontFamily: "Arial")
    })
});
publication.SetTextFrameBounds(frame.Id, 72, 72, 360, 108);
File.WriteAllText("edited.svg", publication.ToSvg());
```

`ReplaceStoryParagraphs` replaces a complete non-table story. It derives `Text`
from the supplied paragraphs, normalizes carriage returns to newlines, and
retains the source's final paragraph-terminator convention, including after
clearing and repopulating the story. An empty paragraph collection clears the
story. Styled runs, paragraph layout, list labels and tab settings use the
shared drawing types. Complete text remains available when it overflows.
Body runs and labels require well-formed UTF-16 and positive finite font sizes;
labels also reject control characters. Labels count toward the story run and
character ceilings, the publication's edit budget and cumulative projection
work, while remaining separate from `PublisherTextStory.Text`.

`SetTextFrameBounds` uses page-local points. Width and height must be positive
and finite; negative offsets allow artwork to bleed beyond the clipped page.
Grouped, rotated, reflected, hidden, unresolved and table frames are rejected.
Both operations retain identifiers, ownership and frame links, then recalculate
the current page scenes, assigned text ranges and overflow. SVG and the optional
PDF package consume those scenes; the Reader adapter can project the current
in-memory document directly.

`SetTextFrameLayout` replaces a qualified frame's column count, gap, padding and
vertical alignment without changing its bounds, graphics, story or links:

```csharp
var current = publication.Pages[0].TextFrames[0];
publication.SetTextFrameLayout(current.Id, 2, 18,
    new OfficeTextPadding(8, 6, 8, 6), OfficeTextVerticalAlignment.Top);
```

`PublisherTextFrame` exposes the current `Padding`, `VerticalAlignment`,
`ColumnCount` and `ColumnSpacing`. `CanEditLayout` and `LayoutEditLimitation`
describe this operation independently of bounds editing. Visible resolved
untransformed document/master frames are eligible, including text-capable
artwork. Tables and table-story aliases are excluded. Unassessed native anchor,
text-direction, wrapping, fitting and vertical-alignment forms retain their
reading and bounds behavior but reject layout editing.

Columns are equal-width, from 1 through 256 within the configured item limit.
Gap and inset sizes use points. The gap must be nonnegative and finite, and
insets/gaps must leave a positive finite content area. An existing frame whose
current insets consume its area can still be corrected with a usable layout.
Top, Center and Bottom align text within each generated content region; native
column balancing, distributed alignment and vertical writing are not provided.
Complete stories and overflow remain available, other frames keep their stored
settings, and `AddLinkedTextFrame` copies the tail's current layout. Fresh SVG/PDF
scenes use the same Core flow. Failed validation, cancellation or resource limits
leave the current document and old snapshots intact.

`AddLinkedTextFrame` appends a continuation frame to the tail of a complete,
resolved non-table story chain. The destination is an existing document page;
the new frame shares the complete styled story and copies the preceding
frame's column count, spacing, padding and vertical alignment. It has no fill,
outline or wrap references and is appended to that page's drawing order.
The source chain must contain ordinary visible untransformed document-page
frames. Master chains, ambiguous chains and non-tail insertion are rejected.

For a two-page publication whose ordinary story starts on the first page:

```csharp
var continued = PublisherDocument.Load("two-page.pub");
var start = continued.Pages[0].TextFrames[0];
continued.SetTextFrameBounds(start.Id, 54, 108, 228, 120);
var continuation = continued.AddLinkedTextFrame(
    start.Id, continued.Pages[1].Id, 54, 72, 468, 600);
Console.WriteLine($"Frame {continuation.Id}, range {continuation.TextStart}+{continuation.TextLength}");
File.WriteAllText("continued-page.svg", continued.ToSvg(1));
```

The operation recalculates chain links, assigned text ranges and overflow.
It retains all story content even when the available frames are too small.
The returned frame is a snapshot; retrieve it again after later story or bounds
edits. Object, text-layout and projection limits apply, and failed creation
leaves the publication and its next object identity unchanged.

`PublisherPage.Pictures` exposes placed pictures separately from the embedded
assets in `Images`. Each placement identifies its page, image asset, unrotated
bounds, crop and page transform. Master pictures remain listed on the master;
their current artwork is included on every referring document page. Coordinates
are page-local points, and crop values are normalized source-edge fractions.
`CanEdit` identifies supported placements; `EditLimitation` explains why another
placement is inspection-only, with the same reason in the read report.

```csharp
var picture = publication.Pages[0].Pictures.First(item => item.CanEdit);
publication.ReplacePictureImage(picture.Id, File.ReadAllBytes("replacement.png"));
publication.SetPictureBounds(picture.Id, 72, 216, 144, 96);
File.WriteAllText("picture-edited.svg", publication.ToSvg());
```

Picture edits support visible, resolved, untransformed top-level rectangular
PNG/JPEG placements with qualified crop and simple paint. Replacement validates
the complete PNG or JPEG through the shared managed image decoder and retains
the selected placement's crop, bounds, identity, page and drawing order. It
creates a separate asset reference so another placement sharing the old image
does not change. Original native assets remain available; unreferenced assets
from superseded replacements are released. Moving a picture also recalculates
its existing rectangular text-wrap exclusions. Grouped, rotated, reflected,
hidden, effected, custom-clipped and unresolved pictures retain their reading
and rendering behavior but reject these edits. Per-image byte and pixel limits,
the image-store entry limit, and cumulative page-projection budgets apply.

`PublisherPage.Shapes` exposes vector artwork in page drawing order, with its
stable identifier, page-local bounds, page transform and edit limitation.
Applied master shapes remain listed on their master. A text-capable shape and
its text frame share one object identifier and current placement: either bounds
operation updates both views. Pictures and table artwork retain their separate
model owners.

```csharp
var shape = publication.Pages[0].Shapes.First(item => item.CanEdit);
publication.SetShapeBounds(shape.Id, 72, 300, 180, 72);
publication.SetShapePaint(shape.Id, OfficeColor.FromRgb(32, 120, 180),
    OfficeColor.Black, strokeWidth: 2);

var ellipse = OfficeShape.Ellipse(72, 48);
ellipse.FillColor = OfficeColor.FromRgb(240, 180, 40);
var added = publication.AddShape(publication.Pages[0].Id, ellipse, 300, 300);
File.WriteAllText("shapes-edited.svg", publication.ToSvg());
```

Bounds and solid-paint edits support visible, resolved, untransformed top-level
rectangles, rounded rectangles, ellipses, diamonds, triangles and lines with
qualified paint. Moving or resizing an existing wrap obstacle recalculates its
referencing text frames. `SetShapePaint` replaces the fill and outline colors,
stroke width and opacities; null colors disable that paint. Stroke widths remain
in points when geometry is resized.

`AddShape` accepts shared rectangle, rounded-rectangle, ellipse and line
descriptors on existing document or master pages. It copies the descriptor,
appends the object to its page's drawing order and returns its current snapshot.
New shapes have no story or text-wrap references. Horizontal and vertical lines
keep a zero logical extent on one axis; both endpoints must differ. Other
shapes require positive extents. Bounds use page-local points and permit clipped
bleed within the publication's coordinate range.

Grouped, transformed, hidden, custom-path, gradient, effected and unresolved
artwork remains inspection-only with its existing rendering diagnostics.
`HasPlacement = false` identifies unavailable geometry. `CreateShape()` returns
a detached descriptor for editable artwork and rejects inspection-only shapes;
their current rendered representation remains in `CreateDrawing()`. Creation
rejects unsupported descriptors rather than dropping their effects.

`PublisherPage.Tables` exposes current tracks, merged spans and styled cell
paragraphs. `CanEditCellText` and `CanEditBounds` describe separate operations;
their corresponding limitation properties explain unsupported cases.
`HasTextMapping` records successful text recovery and does not by itself grant
editing. A table's graphic belongs to the table and is excluded from `Shapes`.

```csharp
var table = publication.Pages[0].Tables.First(item => item.CanEditCellText);
publication.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Updated cell", 10, OfficeColor.Black,
            bold: true, fontFamily: "Arial")
    })
});
if (table.CanEditBounds)
    publication.SetTableBounds(table.Id, 36, 180, 240, 108);

var addedTable = publication.AddTable(publication.Pages[0].Id, 36, 324,
    new[] { 120D, 120D }, new[] { 30D, 30D });
publication.ReplaceTableCellParagraphs(addedTable.Id, 0, 0, new[] {
    new OfficeRichTextParagraph(new[] {
        new OfficeRichTextRun("Item", 10, OfficeColor.Black, fontFamily: "Arial")
    })
});
File.WriteAllText("table-edited.svg", publication.ToSvg());
```

Cell coordinates are zero-based anchors. A merged cell is edited through its
first row and column; covered coordinates are rejected. Paragraph replacement
preserves every other cell, the complete styled story and each cell's paragraph
terminator convention, including when clearing and refilling a cell. Compatible
tables sharing a story update together. Conflicting cell grids, unresolved
boundaries and stories also referenced by ordinary text frames reject cell edits.
Complete text remains inspectable when a cell overflows its rendered frame.
Clearing through an empty paragraph collection retains the final paragraph's
formatting and leaves typing unspecified. Native saving retains that cell's CR
typing; supply an empty styled run when a particular typing style is required.
Native source saving supports qualified fixed cell partitions and supported
text formatting. Paragraph counts can change when source definitions are uniform
within the edited cell; neighboring cells may have different styles. Unknown
native partition records, arbitrary count-changing rewrites of nonuniform cell
definitions and unsupported typing encodings refuse output. Publisher can expand rows when text
wraps or exceeds a cell, so retained stored bounds do not guarantee fixed native
rendered heights.
These native limits are separate from in-memory cell editing.
Independent qualification of nonuniform paragraph retention currently covers
ordinary stories; equivalent table-cell retention remains unqualified.

`SetTableBounds` uses positive finite frame dimensions and proportionally scales
column and row tracks. Imported differences between track sums and frame size
are retained in that ratio. Moving a table recalculates text frames that already
wrap around it. Hidden, grouped, rotated, reflected, unplaced and unsupported
graphic representations retain their reading/rendering behavior and reject
bounds edits. Safe cell text can remain editable independently of those graphics.
Native individual cell fills, borders and cell-specific padding remain unassessed.

`AddTable` copies positive finite column widths and row heights, appends an empty,
unmerged, borderless table to an existing qualified document or master page and
creates its complete story. Master tables remain listed once on the master;
their current artwork appears on referring pages. Populate the cells through
the same paragraph API. Table and story identities come from the shared
publication allocator. Cancellation, text, grid and cumulative projection
limits apply before the candidate replaces the current document.

Each operation validates and projects its candidate before changing the
document. Failed or cancelled edits leave the preceding state usable. Concurrent
mutation is unsupported. Fetch `Pages`, `MasterPages`, `TextStories`, `Images`, frames, tables, pictures and shapes
again after an edit: previously returned objects remain snapshots.
`ReadReport` describes the current reconstruction; `SourceReadReport` retains
the original native decoding evidence without projection findings.

These edits change the in-memory publication and rendered outputs. Native PUB
save-back, standalone publication creation, page/master relationship edits, table topology/styling changes and arbitrary frame relinking are not provided
by this API. Retaining private native source data does not establish opaque
preservation or a native save contract.

`ReadReport.CurrentObjectCount` includes retained imported objects and authored
tables, continuation frames and shapes; `ProjectedObjectCount` counts distinct current objects
represented in the page scenes. `SourceObjectCount` and `SourceReadReport` retain
the original native object inventory. Authored tables, frames and shapes do not add native records
to `InspectedRecordCount`.

## Recovered content and layout

`TextStories` retains complete decoded text, including content that cannot be
placed or does not fit a frame. Paragraphs carry recovered font names, sizes,
emphasis, alignment, spacing, and indentation. Referenced native style values
fill properties that have no direct override. Native bullet labels remain
separate from story text, with their hanging indentation and text position;
declared left tabs use the shared measured tab layout. Style-library metadata,
numbering sequences and other tab variants have explicit recovery limits.
Dynamic field markers and cached
display text remain unevaluated. `Images` retains original embedded payloads;
`GetBytes()` returns an independent copy.

Native transparent-color keys, brightness/contrast controls, grayscale and
two-color picture modes produce a processed PNG in the page scene. Original
assets remain in `Images`. The read report identifies the managed color-space
and threshold approximations; recoloring and extended color controls remain
unassessed. Picture effects require bounded raster decoding. `MaximumRasterPixels`
limits each decoded picture, and `MaximumImageProcessingPixels` limits cumulative
decode/filter work and inspected GIF, WebP or icon frame pixels across picture
references.

Native custom paths retain their declared frame canvas, including inset and
negative coordinates. Literal vertices, open or closed line and cubic paths,
subpaths, and no-fill/no-line controls use the shared drawing model. Custom
picture paths clip the processed image without changing its original asset.
Geometry guides resolve arithmetic, conditional, trigonometric and ellipse
formulas using geometry-space dimensions, signed adjustments, earlier guides,
stroke state and frame sizes in EMUs. The table is limited to 128 integer results.
Invalid references or arithmetic, unsupported device-pixel parameters,
limousine scaling, advanced commands and separately painted path groups use a
reported geometry fallback. Producer rounding, native winding and picture-mask
appearance remain unqualified against Publisher output. `Limits.MaxItems` bounds
cumulative native vertex, guide and command work and copied shape/mask commands
as well as drawing elements.

Each page's `TextFrames` exposes native frame identifiers, story links, order,
column settings and wrap-object references. Linked stories follow explicit frame
links or complete native story/frame catalog bindings across frames and pages.
Catalog counts, object references and consecutive ordinals must agree; shape
ordinals alone do not establish a chain. Columns and rectangular wrap exclusions use the
shared managed text engine; tight outlines and native break positions can differ.
`TextStart` and `TextLength` identify the assigned range in `PublisherTextStory.Text`,
including paragraph separators. These ranges describe recovered layout, rather
than stored native break positions. A null range means placement is unresolved;
`HasOverflow` identifies remaining story content after the final frame.

Frame `X`, `Y`, `Width` and `Height` describe the unrotated rectangle in points.
`PageTransform` maps frame-local points into the page, including the frame's own
rotation/reflection and enclosing group transforms, including non-uniform scale. Use it when placing an
annotation or inspecting the visible frame corners.

```csharp
foreach (var page in publication.Pages)
    foreach (var frame in page.TextFrames)
        Console.WriteLine($"Frame {frame.Id}, story {frame.StoryId}, " +
            $"range {frame.TextStart}+{frame.TextLength}, overflow {frame.HasOverflow}");

var firstFrame = publication.Pages[0].TextFrames[0];
var corner = firstFrame.PageTransform.TransformPoint(new OfficeIMO.Drawing.OfficePoint(0, 0));
Console.WriteLine($"Transformed top-left: {corner.X}, {corner.Y}");
```

## Inspect native tables

`PublisherPage.Tables` exposes native column widths, row heights, cell spans,
styled cell paragraphs and text. Grid indices are zero-based; a merged cell
appears once. Cell coordinates are local to the table and use points. Table
`PageTransform` maps those coordinates into the page, including rotation and
enclosing groups. The frame size and the grid track dimensions remain distinct.
Imported tracks use saved layout endpoints within each axis. Nominal sizes can
remain smaller when text has expanded a row. An axis with incomplete endpoints
retains nominal sizes and reports `PUB_TABLE_TRACK_POSITIONS_UNASSESSED`;
nonincreasing or excessive endpoints reject the input. These values describe
stored geometry and do not run Publisher's current text fitting or reflow.
Tables inherited from a master are listed on that master in `MasterPages`.
Hidden tables and tables with unavailable geometry retain their grid and mapped
cell text. `HasPlacement = false` identifies a table whose frame cannot be used
for page placement; its recovered cell data remains available.

```csharp
foreach (var page in publication.Pages)
    foreach (var table in page.Tables) {
        Console.WriteLine($"Table {table.Id}: {table.RowCount} x {table.ColumnCount}");
        if (!table.HasTextMapping) continue;
        foreach (var cell in table.Cells)
            Console.WriteLine($"({cell.RowIndex}, {cell.ColumnIndex}), " +
                $"span {cell.RowSpan} x {cell.ColumnSpan}: {cell.Text}");
    }
```

An unresolved native cell-text mapping retains the grid with
`HasTextMapping = false` and empty cell paragraphs/text. The read report carries
the omission; complete recovered stories remain in `TextStories`. Cell fills,
individual borders and padding remain unassessed. These source records describe
recovery; changing the separate page drawing does not update the table model.

## Export a page as SVG

```csharp
var exported = publication.ToSvgResult(pageIndex: 0);
File.WriteAllText("page-1.svg", exported.Value);

foreach (var diagnostic in exported.Report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.LossKind}: {diagnostic.Message}");
```

The result combines source recovery with image-rendering diagnostics.
`ToSvg()` returns markup directly. `PublisherSvgOptions` controls output scale,
dimension units, an optional image codec, and resource-ID prefixes for inline
composition. Default dimensions are physical points.

Reference [OfficeIMO.Publisher.Pdf](../OfficeIMO.Publisher.Pdf/README.md) for
multi-page PDF conversion.

## Resource limits and acceptance

```csharp
using OfficeIMO;

var options = new PublisherReadOptions {
    MaximumPages = 100,
    MaximumImageBytes = 8 * 1024 * 1024,
    Limits = new OfficeLegacyImportLimits {
        MaxInputBytes = 32 * 1024 * 1024,
        MaxTextCharacters = 1_000_000
    }
};
var recovered = PublisherDocument.Load(inputStream, options);
```

Path, byte-array and stream overloads use the same decoder. Streams read from
their current position, support non-seekable input, and remain open. Options are
snapshotted for each operation. Cancellation applies to input reading, native
records, and page projection. Exceeding a configured limit rejects the operation
instead of returning a truncated document. Item and text limits also bound
projected content, including repeated master use. `MaxItems` bounds image-store
entries even when their payload cannot be recovered and cumulative projected
gradient stops, including focus expansion, as well as recovered table tracks and
cells. `MaxTextCharacters` bounds copied table text. Native linear fills preserve colors,
angle, focus and transparency in the shared drawing model; inspect the conversion
report for unsupported shading, anchors and opacity ratios. `MaxInputBytes` also bounds
cumulative encoded image-payload processing, including repeated delayed
references; delayed decoding results are reused within the operation.
`MaxTextCharacters` also bounds cumulative text measured while continuing stories;
`MaxRecords` bounds wrap-region inspection. Native text frames support up to
256 columns within the configured item limit.

Recovery reports distinguish approximation, omission and unassessed content.
`RequireNoLoss()` rejects any of those categories. Current native recovery
reports unassessed publication features, so a successful load is not a claim of
complete Publisher fidelity.

Native WMF/EMF assets remain available in `Images`. Their page projection uses an
explicit application `PublisherReadOptions.ImageCodec` when supplied; otherwise
it displays a placeholder and reports the omission. Codec output is subject to
pixel and encoded-byte limits. No external decoder is installed or invoked by
the package.

Publisher 97/98 and 2000 profiles are recognized and rejected. Native `.pub`
writing, macro execution, external-resource
refresh, and embedded-object activation are outside this package's contract.
