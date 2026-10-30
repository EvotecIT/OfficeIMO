# OfficeIMO.Publisher

Read Microsoft Publisher publications into positioned pages, styled text
stories and embedded image assets, edit supported stories and frame bounds, and
render the current model. The managed reader handles the Publisher
2002-and-later compound-document generation without installing Publisher.
[The support contract](SUPPORT.md) identifies the qualified fixtures and rendering
limits.

`OfficeIMO.Publisher` references only `OfficeIMO.Core`. Native records belong to
this package; compound storage, OfficeArt properties and image envelopes, text
layout, and SVG rendering use shared OfficeIMO engines.

## Read and inspect a publication

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

## Edit stories and frame bounds

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

Each operation validates and projects its candidate before changing the
document. Failed or cancelled edits leave the preceding state usable. Concurrent
mutation is unsupported. Fetch `Pages`, `MasterPages`, `TextStories` and frames
again after an edit: previously returned objects remain snapshots.
`ReadReport` describes the current reconstruction; `SourceReadReport` retains
the original native decoding evidence without projection findings.

These edits change the in-memory publication and rendered outputs. Native PUB
save-back, object creation, table editing and frame relinking are not provided
by this API. Retaining private native source data does not establish opaque
preservation or a native save contract.

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
writing, editable Publisher reconstruction, macro execution, external-resource
refresh, and embedded-object activation are outside this package's contract.
