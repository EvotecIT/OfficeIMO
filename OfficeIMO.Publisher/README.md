# OfficeIMO.Publisher

Recover Microsoft Publisher publications into positioned pages, styled text
stories, and embedded image assets. The managed reader handles the Publisher
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
Each page's `Drawing` is the shared `OfficeDrawing` model. Callers can add artwork
to that scene for subsequent exports. Recovered page content is held in a detached
clipped group; this API does not directly edit native text frames or modify the
publication file.

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
Geometry guides, advanced commands and separately painted path groups use a
reported geometry fallback. Native winding and picture-mask appearance remain
unqualified against Publisher output. `Limits.MaxItems` bounds cumulative native
path decoding and copied shape/mask commands as well as drawing elements.

Each page's `TextFrames` exposes native frame identifiers, story links, order,
column settings and wrap-object references. Linked stories follow their native
links across frames and pages. Columns and rectangular wrap exclusions use the
shared managed text engine; tight outlines and native break positions can differ.
`TextStart` and `TextLength` identify the assigned range in `PublisherTextStory.Text`,
including paragraph separators. These ranges describe recovered layout, rather
than stored native break positions. A null range means placement is unresolved;
`HasOverflow` identifies remaining story content after the final frame.

Frame `X`, `Y`, `Width` and `Height` describe the unrotated rectangle in points.
`PageTransform` maps frame-local points into the page, including the frame's own
rotation/reflection and enclosing group transforms. Use it when placing an
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
gradient stops, including focus expansion. Native linear fills preserve colors,
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
