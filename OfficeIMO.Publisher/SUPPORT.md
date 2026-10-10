# Publisher support contract

OfficeIMO.Publisher is a native read-and-convert library. Its public model owns
publication pages, master definitions, styled source stories, native table grids, embedded assets
and a recovery report. Shared `OfficeDrawing` owns the reconstructed visual scene;
the optional PDF package consumes that scene without another native decoder.

## Native profiles

| Profile | Read behavior | Evidence boundary |
| --- | --- | --- |
| Publisher 2002-and-later Contents/Quill/OfficeArt generation (`E8 AC 2C 00`) | Bounded structured decoding and positioned reconstruction | Apache POI Simple, Sample and Sample_2010 publications; brochure and newsletter corpus fixtures |
| Earlier Contents generation (`E8 AC 22 00`) | Explicit unsupported-profile exception; the shared signature and Quill presence do not identify the exact 97/98 or 2000 release | Apache POI Sample2000 and Sample98 fixtures |
| Encrypted, damaged or other publication generations | No decryption or salvage fallback | Required streams, lengths, references and encodings are validated; unknown versions are rejected |

The generation signature does not establish qualification for every Publisher
release or object type. Independent producer files exercise real native storage;
they are not produced by an OfficeIMO writer.

## Reconstruction

| Content | Current behavior | Limit |
| --- | --- | --- |
| Document pages | Native order, physical size and page clipping | Utility/reference definitions are excluded |
| Master pages | Recovered definitions applied to referring pages | Missing references report omission; nested master inheritance is unassessed |
| Text | UTF-16 stories, styled paragraphs, referenced native style values, reciprocal frame links and ordinals, sequential columns, continuation and overflow reporting | Shared measurement approximates native frame breaks and column balancing; style names and auxiliary inheritance metadata are not exposed as an editable style library |
| Lists and tabs | Native Unicode bullet labels, qualified Symbol/Wingdings marker normalization, hanging indentation, text position and declared left tabs | Numbering sequences, additional tab alignments/leaders and drop caps are unassessed; undeclared tab stops use the shared 36-point interval |
| Tables | `PublisherPage.Tables` exposes native tracks, spans, styled paragraphs, text and placement; master tables remain on their owner | Unresolved text mappings retain the grid with `HasTextMapping = false`; individual cell borders, fills and padding are unassessed |
| Shapes | Rectangles, rounded rectangles, ellipses, diamonds, triangles and lines; solid paints and basic shadows | Other geometries and unsupported fill types report approximation; compound and non-solid strokes remain unassessed |
| Custom paths | Literal eight-byte signed and compact unsigned vertices in declared geometry space; bounded SG guide formulas and backward references; implicit open/closed line or cubic paths, explicit move/line/cubic/close/end commands, subpaths and no-fill/no-line controls | Device-pixel guide operands, limousine scaling, advanced commands and separate paint groups retain an explicit fallback; guide rounding, shared nonzero winding and native appearance remain unqualified |
| Linear fills | Native color stops, fixed-point angles, focus ramps and foreground/background opacity for fill types 4 and 7 | Custom anchors, native shading corrections and unrepresentable alpha ratios report approximation; native Publisher rendering remains unqualified |
| Line details | Native caps, joins, miter limits and dash order; triangle, stealth, diamond, oval and open-arrow ends on lines | Dash spacing and marker dimensions use stroke-relative approximations; unknown values report loss; decorations on unsupported open geometry are omitted |
| Groups | Native child coordinate spaces, nested rotation/reflection and hidden-descendant suppression | Missing anchors report unresolved transforms; comparison against native Publisher rendering remains unqualified |
| Text wrapping | Native frame exclusion references and object wrap distances; transformed rectangular exclusions | Widest available interval per horizontal band; side selection, tight/through outlines and native font metrics can differ |
| Pictures | Embedded and delayed OfficeArt payload extraction, Publisher GIF envelopes, positioned images, source crop and supported custom-path masks | WMF/EMF use a supplied application codec or reported placeholder; native custom-mask appearance remains unqualified |
| Picture controls | Source-RGB transparency keys, encoded-RGB brightness/contrast, preserve-grays tone handling, BT.709 grayscale and a 50-percent two-color threshold; original assets remain available | Color space, effect ordering and native Publisher pixels remain unqualified; recoloring and extended color controls are unassessed; invalid tone values or unresolved keys report omission |
| Fields, links and active content | Cached story characters are retained; source active content stays inert | No field evaluation, hyperlink reconstruction, macro execution, link refresh or object activation |
| Output | SVG per document page and multi-page PDF through shared engines | No native save-back or editable publication writer |

`OfficeIMO.Workflows` uses the same native codec for its `publisher-pdf` route,
file batches and document previews. Source recovery diagnostics survive output
publication and checkpoint reuse. Strict acceptance and resource-limit failures
preserve existing destinations.

`OfficeIMO.Reader.Publisher` projects complete stories once, mapped native table
rows, bullet markers, physical page inventory and optional original image bytes.
Ordinary stories use paragraph blocks; mapped table stories use one complete
table block alongside their structured rows and native page citation. Generic
column names preserve the first row as data. Merged spans flatten to one anchor
value with approximation evidence; row limits retain total counts and omission
reports. Master tables are extracted once without a physical-page citation.
Unresolved cell text remains in complete stories rather than an empty dataset.
Native losses survive Reader transport and semantic PDF projection, which
correlates table blocks and rows to avoid repeated text. Typography, page artwork
and frame flow remain explicitly omitted from the semantic projection. The same
adapter is included in `OfficeIMO.Reader.All`.

Source object counts measure recovered descriptors and distinct projected
objects. They do not measure pixel fidelity or prove that all content of an
object survived. Complete source stories remain inspectable even when their
visual placement is incomplete.

`PublisherPage.TextFrames` records recovered frame geometry and links. Assigned
UTF-16 ranges refer to normalized `PublisherTextStory.Text`; they are derived
from managed measurement, not claimed as native cached break positions. Later
empty frames have zero-length ranges. Native chains are checked for reciprocal
links, matching story ownership, consecutive ordinals, cycles and disconnected
frames. Ambiguous unlinked roots report omission instead of guessed ordering.
A missing printable frame stops subsequent assignment for that story. Frame
`X`, `Y`, `Width` and `Height` describe its unrotated rectangle; `PageTransform`
maps frame-local points into page-local points after object and ancestor-group
rotation/reflection. The same transforms reach artwork and rectangular wrap
exclusions. Hidden group descendants retain their source stories and assets but
do not paint or exclude printable text.

## Qualification and limits

[Fixture provenance](../OfficeIMO.Publisher.Tests/Fixtures/provenance.json)
records source revision and SHA-256 hashes. The corpus is distributed under the
Apache POI license and notice included beside the fixtures. Format descriptions
include [Apache POI's Publisher notes](https://poi.apache.org/components/hpbf/file-format.html)
and Microsoft's [Office Drawing binary specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/).

Managed tests check page inventories, frame coordinates, source style values,
table cell boundaries, grouped placements, raw image recovery, SVG dimensions,
PDF page/text output and carried fidelity evidence. They also cover malformed
native offsets, cyclic Quill directories, invalid UTF-16, resource ceilings,
cancellation and caller stream ownership. The native newsletter exercises ten
linked stories, including chains whose object order differs from story order,
and declared picture exclusions. Multi-column property and malformed-link
mutations protect the native codec boundary; they are synthetic evidence.
The table sample exposes six distinct cells through the native public model and
Reader rows. Cell-span, missing-text-map and master-ownership mutations protect
merged geometry, explicit unresolved text, and single-owner extraction. Reader
checks cover generic headers, row limits, canonical table views, CSV output, JSON
transport and semantic PDF text deduplication. These mutations do not qualify
individual cell styling or Publisher-rendered table appearance.
Group rotation/reflection mutations verify text-frame corners, picture placement,
hidden-child wrapping and scene copying. A reflected picture starting above the
page retains its pixels through the group transform and final page clip.
Constructed nested native records check
composition about distinct centres, quarter-turn anchor dimensions and the
projected-item ceiling. These tests do not establish Publisher-rendered fidelity.
Native line-property mutations check caps, joins, fixed-point miter limits,
dash order and spacing, all five supported marker kinds, disabled strokes,
unknown values and scene copying. SVG, raster and PDF retain open arrows as
outlines and other markers as filled geometry. Marker dimensions and native dash
rendering remain unqualified against Publisher-produced output.
Custom-path record mutations check coordinate offsets, compact vertices,
line/cubic command consumption, subpaths, paint controls, malformed arrays,
guide/command fallback reports and inset picture masks. SG guide checks cover
all 17 stored formula identifiers, unsigned constants, signed adjustment values,
geometry-space centres and dimensions, stroke use masks, physical frame EMUs,
backward references, fixed-point degrees and the 128-record ceiling.
Device-pixel operands require an output-device context and retain a reported
fallback. Malformed references, division by zero, invalid square roots and
32-bit overflow also retain fallback evidence without partial custom artwork.
Integer rounding follows the corresponding [VML formula contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.formula):
products round to nearest with ties toward positive infinity, averages truncate
toward zero, and inexact operations floor their results. These managed checks
do not qualify Publisher's rounding or discontinuous guide behaviour.
High-coordinate
shape/mask checks retain the full unsigned compact range;
eight-byte coordinates retain signed negative values. Shared array decoding
preserves following complex properties when native lengths exclude their
six-byte headers, including empty arrays. Cumulative limits account for decoded vertices/segments,
evaluated guides, expanded commands and copied shape/mask paths. The producer corpus has no custom
paths; these managed scene and SVG/raster/PDF checks do not establish native
Publisher winding or mask fidelity.
Native linear-fill mutations check physical and aspect-scaled angles, positive
and negative focus, reversed ramps, duplicate color positions, transparency,
fill-use masks and picture layering. Foreground/background opacity uses a
common shape opacity and relative stop alpha; ratios requiring RGBA rounding
report loss. Multi-color fills retain foreground opacity and report distinct
background opacity as an approximation. Native shading corrections use reported
sRGB interpolation. Custom fill rectangles and view-origin anchors retain their
primary color with a mapping diagnostic; non-rotating fills on transformed
frames report approximation. Other gradient, pattern, texture and background
fill modes retain their current approximation boundary. The five supported
producer fixtures contain no declared non-solid fills, so these record checks
and managed SVG/raster/PDF outputs do not establish native gradient fidelity.
Picture-control mutations exercise transparency keys before tone changes,
brightness endpoints, contrast, preserve-grays, grayscale and two-color
precedence. Controlled raster output in original native picture frames is
checked through SVG, raster and reopened PDF output; original embedded payloads
remain unchanged. Invalid or unqualified controls retain their diagnostics when
other valid controls apply. Per-image and cumulative pixel ceilings account for
inspected GIF frames and repeated references; cumulative encoded-byte checks
reject repeated large-payload decoding, and codec cancellation aborts the read. These are managed
record and artifact checks; the supported producer fixtures contain no enabled
visible picture controls, so native Publisher pixel equivalence remains open.
The brochure and newsletter also verify referenced style defaults, direct
formatting precedence, native bullet labels and tab-array positions. Those
checks compare decoded values with the native records, not Publisher-rendered
line breaks. Bullet labels do not change source-story character ranges.

The libmspub 0.1.5 reader independently confirms page dimensions, text-frame
coordinates, text styling and table semantics for the simple/table sample
family. It produces no output for the brochure/newsletter fixtures in the
qualification environment; those corpus tests therefore do not claim an
independent-reader rendering comparison. Publisher-produced PDF exports and
pixel comparisons against native Publisher remain an explicit qualification
gap. No external program is required by normal builds or runtime conversions.

The default resource limits are 64 MiB input, 4 million source/projected text
characters, 1 million inspected records, 250,000 source/projected items,
512 compound streams, 1,024 document/master pages, nesting depth 64,
16 MiB per extracted/projected image and 64 MiB aggregate image accounting.
Decoded picture-effect rasters and application codec output are limited to
8 million pixels per image. Cumulative picture-effect work is limited to
64 million pixels, charging the selected raster once for decoding and once for
filtering, plus each inspected GIF, WebP or icon frame across all references.
Image extraction and each actual picture decode also charge their encoded bytes
against `Limits.MaxInputBytes`. Effect decoding fails the read when the image or
selected codec cannot produce a raster within those bounds.
Multi-frame effect sources use their first frame and report omitted frames;
`Images` retains the complete payload. Limits are
caller-configurable, reject oversized input and also account for repeated
projection work. The input byte ceiling also bounds cumulative encoded image
processing, and unavailable image-store entries consume the item ceiling.
The item ceiling also bounds cumulative projected gradient-stop work, including
focus expansion and repeated master-page projection, and recovered table tracks
and cells. Copied table text consumes a cumulative text ceiling. A native color-stop table
supports at most 256 entries; malformed or larger tables retain primary paint
with an invalid-gradient diagnostic.
Delayed decoding is cached and cancellation is checked between image entries.
The text ceiling also bounds cumulative continuation measurement, and the record
ceiling bounds exclusion-region inspection. Native frames support at most
256 columns within the configured item ceiling.
Finite text/table bleed coordinates survive projection and copying until page
clipping is applied. Each source text story currently has the shared drawing
ceiling of 100,000 characters and 4,096 rich-text runs/paragraphs.

Open implementation and qualification work belongs to the
[Publisher roadmap](../Docs/ROADMAP.md#publisher-publication-recovery).
