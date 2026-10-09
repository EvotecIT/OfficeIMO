# Publisher support contract

OfficeIMO.Publisher is a native read-and-convert library. Its public model owns
publication pages, master definitions, styled source stories, embedded assets
and a recovery report. Shared `OfficeDrawing` owns the reconstructed visual scene;
the optional PDF package consumes that scene without another native decoder.

## Native profiles

| Profile | Read behavior | Evidence boundary |
| --- | --- | --- |
| Publisher 2002-and-later Contents/Quill/OfficeArt generation (`E8 AC 2C 00`) | Bounded structured decoding and positioned reconstruction | Apache POI Simple, Sample and Sample_2010 publications; brochure and newsletter corpus fixtures |
| Publisher 2000 compound generation (`E8 AC 22 00` with Quill) | Explicit unsupported-profile exception | Apache POI Sample2000 fixture |
| Publisher 97/98 generation | Explicit unsupported-profile exception | Apache POI Sample98 fixture |
| Encrypted, damaged or other publication generations | No decryption or salvage fallback | Required streams, lengths, references and encodings are validated; unknown versions are rejected |

The generation signature does not establish qualification for every Publisher
release or object type. Independent producer files exercise real native storage;
they are not produced by an OfficeIMO writer.

## Reconstruction

| Content | Current behavior | Limit |
| --- | --- | --- |
| Document pages | Native order, physical size and page clipping | Utility/reference definitions are excluded |
| Master pages | Recovered definitions applied to referring pages | Missing references report omission; nested master inheritance is unassessed |
| Text | UTF-16 stories, font names and sizes, bold, italic, underline, baseline, paragraph alignment, spacing and indents | Shared line breaking and metrics approximate native layout; measured overflow reports omission; linked frame continuation, picture wrap, named styles, custom tabs, lists and drop caps are not fully reconstructed |
| Tables | Native track sizes, cell spans and styled cell text | Individual cell borders, fills and padding are unassessed |
| Shapes | Rectangles, rounded rectangles, ellipses, diamonds, triangles and lines; solid paints and basic shadows | Other geometries and non-solid fills report approximation; compound strokes and arrowheads remain unassessed |
| Groups | Native child coordinate spaces and nested placement | Group rotation/mirroring reports approximation when present |
| Pictures | Embedded and delayed OfficeArt payload extraction, Publisher GIF envelopes, positioned images and source crop | WMF/EMF use a supplied application codec or reported placeholder; text wrapping, recoloring and other picture effects are unassessed |
| Fields, links and active content | Cached story characters are retained; source active content stays inert | No field evaluation, hyperlink reconstruction, macro execution, link refresh or object activation |
| Output | SVG per document page and multi-page PDF through shared engines | No native save-back or editable publication writer |

Source object counts measure recovered descriptors and distinct projected
objects. They do not measure pixel fidelity or prove that all content of an
object survived. Complete source stories remain inspectable even when their
visual placement is incomplete.

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
cancellation and caller stream ownership.

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
Application raster codec output is limited to 8 million pixels. Limits are
caller-configurable, reject oversized input and also account for repeated
projection work. The input byte ceiling also bounds cumulative encoded image
processing, and unavailable image-store entries consume the item ceiling.
Delayed decoding is cached and cancellation is checked between image entries.
Finite text/table bleed coordinates survive projection and copying until page
clipping is applied. Each source text story currently has the shared drawing
ceiling of 100,000 characters and 4,096 rich-text runs/paragraphs.

Open implementation and qualification work belongs to the
[Publisher roadmap](../Docs/ROADMAP.md#publisher-publication-recovery).
