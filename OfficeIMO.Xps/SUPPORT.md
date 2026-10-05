# XPS/OpenXPS support

The native package lifecycle and the rendering profile are separate contracts.
Preserving an unsupported native element does not mean that it can be rendered.

| Operation | Supported contract | Boundary |
| --- | --- | --- |
| Read | Microsoft XPS and ECMA-388 OpenXPS; OPC relationships/content types; multiple fixed documents; ordered page references; UTF-8/UTF-16 XML; bounded interleaved OPC piece assembly | Materialized loading, not progressive streaming; protected packages are not supported |
| Create | Both dialects; pages, vector paths, embedded fonts, Unicode glyph runs, PNG/JPEG/TIFF and supported JPEG XR image placement | Typed creation is a bounded fixed-page profile, not a complete schema object model |
| Edit | Detached native page XML; loaded document/page insertion, reordering, transfer and reference removal; shared backing for repeated references; encoded resource replacement; owning APIs for relationship-owned DocumentStructure and StoryFragments; atomic combined page/fragment edits | Removed parts/resources remain preserved; story references to removed pages and resulting empty stories are removed; dangling known name/story addresses reject edits; unknown semantic extensions remain opaque |
| Save | Original dialect; native page content and opaque parts retained; required-resource relationships emitted, including profiles and transitive dictionary resources; deterministic ZIP output on the same runtime | ZIP metadata/XML bytes may change; interleaved storage is normalized to atomic parts; no dialect conversion; signed packages cannot be rewritten |
| Text extraction | Page-element UnicodeString runs in markup order, excluding resources and brush visuals | No inferred reading order, paragraphs, or glyph-ID-to-Unicode reconstruction |
| Shared model and Reader | Native logical-order Unicode blocks with physical page citations; recursive structure, list markers and table spans in the shared model; authored Path/Canvas descriptions as payload-free assets; bounded rectangular tables and optional strict SVG previews; modular `.xps`/`.oxps` Reader registration | Unreferenced text follows native stories in page/markup order; missing structure and unavailable text are diagnosed; overlapping text ownership rejects projection; preview payload is bounded to 512 pages/128 MiB independently of description assets |
| Native logical structure | Relationship-owned StoryFragments; named page/Canvas/Path/Glyphs references with authored Path/Canvas accessibility descriptions; DocumentStructure story-reference order; continued paragraphs, sections, lists, figures and tables; StoryBreak boundaries; list markers and cell spans | No inferred structure on unstructured pages; unknown extensions and unresolved content produce diagnostics; missing Unicode is not reconstructed; one million work units and 16 million resolved text/description characters bound each read |
| Paths | Abbreviated geometry, fill rules, explicit figures/segments with fill/stroke suppression, dashes with separate endpoint/dash caps, triangle caps, clipped miters and the degenerate-segment limit, matrix transforms and clipping | Extended strokes use bounded adaptive vector outlines; native Windows confirmation remains for degenerate and mixed-segment cap rules where the independent engines disagree with the specification |
| Text rendering | Embedded TrueType programs/collections, obfuscation, explicit glyph IDs, cluster mappings, advances/offsets, horizontal bidi, bold/italic style simulation, sideways top-center positioning with vertical metrics or OS/2/hhea fallbacks | Outlined output; unsupported font programs are diagnosed; sideways runs require even BidiLevel |
| Brushes | Hex/scRGB and ICC ContextColor solids/gradient stops; linear/radial gradients; scoped and external package resource dictionaries; ICC-managed PNG/JPEG/TIFF/JPEG XR, native integer sRGB/gray defaults for non-ICC image descriptions, and visual brushes with absolute viewbox/viewport mapping, matrix transforms, Tile/FlipX/FlipY/FlipXY repetition, non-tiled fills/strokes, and alpha opacity masks | Unsupported image/profile channel combinations, unsupported colorimetry such as non-sRGB PNG cICP, unsupported TIFF and JPEG XR encodings are diagnosed |
| Navigation | Safe web/mail links; page/document/sequence named targets with scoped first-occurrence lookup; sequence page numbers projected into SVG filenames; PDF links, named destinations and DocumentStructure outlines | Non-page unresolved and unsafe destinations are diagnosed; known fixed-page destinations follow structural moves; links to removed pages are unresolved; PDF link hit areas are rectangles and path destination positions use conservative geometry bounds |
| Gradient transforms | Affine linear and radial gradients convert through Core, including rotation, shear and reflection; native Pad supports boundary/exterior point foci and endpoint paint outside the cone; bounded interior, boundary and exterior radial Repeat/Reflect expansion retains vector PDF shading | PDF uses function-based vector shading when finite expansion exceeds 256 stops or has no finite cycle bound; shared Drawing/raster/SVG retains explicit spread outside that bound; direct SVG preserves Pad and boundary/exterior Repeat/Reflect fields, including unbounded tangent regions; native consumer differences remain below |
| SVG | Self-contained images and glyph outlines; authored Path/Canvas descriptions retained as title/desc metadata; strict by default; explicit partial result with diagnostics | Unknown markup/attributes are diagnosed; no claim of complete XPS consumer conformance |
| Drawing/images | Existing managed Core scene and image exporters | Shared viewport, element, geometry, raster, and codec limits still apply; any reported SVG import loss rejects conversion |
| PDF | Optional thin bridge retaining vector paint, dimensions, native alpha masks and searchable Unicode clusters; native paragraph/list/table/figure tags with authored figure descriptions, continued cross-page containers, declared story order, list labels and cell spans; header/footer artifacts | Unstructured pages use markup order; unassociated fragments use page order; unknown/unresolved/overlapping semantics reject strict mapping; an explicit paint-only mode retains markup-order text; no inferred figure descriptions, PDF/UA qualification, print-ticket or signature migration; clipped/transparent source text remains searchable |
| Workflows | Native `xps-pdf` route for both extensions; single and batch conversion, untagged-source assembly and preview; bounded PDF stream serialization, reopen validation and staged publication; existing PDF encryption/compression settings | `Faithful` profile only; strict semantic mapping unless explicitly disabled; tagged PDF merging is blocked by the PDF owner; native and workflow input ceilings both apply; a successful reopen does not establish whole-document visual equivalence |
| Security | Package-local resource resolution; no external fetch; DTD prohibition; shared backing for repeated page parts; bounded ZIP/XML/page and expanded SVG node/character/resource-binding growth; cooperative cancellation; atomic path saves | Inspection does not authenticate signatures or make arbitrary native documents trusted |

ICC ContextColor uses Core's supported RGB, gray, CMYK and N-channel profiles,
converting to sRGB with media-relative colorimetric intent and no black-point
compensation. Alpha and channel values are clamped to the native range. Profile
parsing has a 4 MiB per-profile ceiling and a 64 MiB aggregate parser allowance.
PrintTicket color overrides are not interpreted. Image brushes apply associated
`ColorConvertedBitmap` profiles in preference to embedded profiles. Core decodes
RGB and gray PNG/JPEG/TIFF, including PNG/TIFF alpha, and CMYK JPEG/TIFF device
channels before ICC conversion. The first image is limited to four million pixels;
physical dimensions are retained. Malformed, oversized or incompatible profiles
and unsupported non-ICC color metadata are diagnosed. A usable associated profile
takes precedence; an unusable or channel-incompatible associated profile falls back
to a usable embedded profile. If no usable profile remains, strict conversion
reports an error. Unprofiled CMYK JPEG/TIFF/JPEG XR images require a usable ICC profile;
OfficeIMO does not substitute an approximate RGB conversion or a default SWOP
profile. Profile and decode budgets still apply before fallback.

Without ICC, supported integer gray/RGB images use the native sRGB sample rules.
PNG gamma/chromaticity and JPEG/TIFF non-ICC calibration descriptions do not
override these defaults. The adapter normalizes these resources so subsequent
SVG/PDF consumers cannot reinterpret the descriptions. TIFF uses the first IFD,
ignores the display Orientation tag, and ignores an extra sample declared as
unspecified. The managed TIFF subset accepts unsigned eight/sixteen-bit and finite
floating sixteen/twenty-four/thirty-two/sixty-four-bit gray/RGB/CMYK components and eight-bit palette indices in either byte order. Chunky/planar strips
and tiles use uncompressed, LZW, PackBits or Deflate payloads, including word-based
horizontal prediction and floating-point prediction for LZW/Deflate. Sample and associated-alpha precision is retained through
ICC conversion before eight-bit RGBA projection. Mixed component widths, reversed
bit order, sixteen-bit palette indices, signed and undefined
sample encodings are rejected. Floating samples are normalized device components;
the decoder preserves their precision for ICC conversion and unassociation, then
clips to SDR output. It does not infer linear scRGB or rescale scientific ranges,
and non-finite color or alpha samples fail content validation and decoding.
Unspecified extra channels and tile padding remain ignored.

## Qualification

Integer JPEG/TIFF default qualification includes eight independently encoded
synthetic image fixtures in both dialects: RGB, gray, alpha, calibration tags,
TIFF orientation, an embedded profile and an unspecified extra sample. Managed
pixels and independently rendered SVG/PDF interior samples agree within three channel
values. GhostXPS agrees for fourteen cases, including the correct raw TIFF viewbox;
it changes sample order for unspecified extra channels in two cases. The differing reference output is retained;
native Windows confirmation and photographic producer coverage remain open.

Floating-point qualification covers 304 independently encoded and decoded LibTIFF
fixtures: 16/24/32/64-bit samples, both byte orders, compression/prediction and
strip/tile layouts, alpha, RGB, grayscale and CMYK. All 377,264 decoded source
samples match the producer inputs, and managed normalized RGBA pixels match
exactly. The 76 float24 fixtures additionally use the unmodified imagecodecs
sample converter to verify the TN3 representation. Both XPS dialects exercise
raster, SVG and PDF-reader output, including
explicit-profile CMYK conversion. Low associated alpha retains precision through
ICC conversion. This does not establish native Windows floating-TIFF acceptance
or a general HDR/scientific tone-mapping policy.

Unsigned sixteen-bit qualification uses 47 independently encoded LibTIFF fixtures,
including 32 compression/storage combinations, gray, associated-alpha RGB/CMYK,
embedded RGB/gray/CMYK profiles and two-page input. LibTIFF verifies the original
sample words; LittleCMS supplies the ICC reference colors. Core decoding matches
unprofiled reference pixels exactly and profiled pixels within two channel values.
The 45 paintable fixtures are placed in both dialects and checked through managed
raster, standalone SVG and PDF readback. Their 22,230 pixel-center probes retain
paint within one channel value in managed output and within four in MuPDF PDF/SVG
output. Four unprofiled CMYK placements reject explicitly. Ghostscript's interpolated
PDF output differs by up to 65 on this discontinuous sample field, while disabling
interpolation preserves pixel-center samples within one channel value. GhostXPS differs
by up to 255 and omits some TIFF encodings. These consumer differences are retained
as qualification limits. Native Windows confirmation remains open.

The focused tests exercise both dialects, package reopening and native edits,
opaque-part preservation, deterministic saves, fonts/glyph positioning, image
placement, gradients, external resource dictionaries, hostile inputs, cancellation,
and preservation of the destination on a rejected signed-package save. Native-editing
checks cover repeated document/page references, reference metadata and link-target
preservation, atomic rejection at package/page limits, and resource replacement.
Generated piece fixtures cover interleaved metadata and resources, non-sequential
ZIP ordering, missing/duplicate/ambiguous pieces, and aggregate part bounds.
Microsoft `System.IO.Packaging` 10.0.11 independently rewrites 20 generated
interleaved packages across both dialects: cross-piece writes, truncation within a
piece and at a boundary, empty parts, and terminal-piece growth. Inputs include
empty pieces, 13-piece sequences, reversed ZIP ordering, and uppercase piece
suffixes. Resource bytes and both pages' SVG output survive loading; every logical
part survives OfficeIMO normalization and reopening in Microsoft's packaging reader.
The comparison library is confined to an isolated validation harness. This qualifies
rewriting of generated piece storage; independently produced interleaved documents
remain unqualified.

Independent input: [Ecma's published ECMA-388 XPS document](https://ecma-international.org/wp-content/uploads/ECMA-388.xps),
which uses the Microsoft XPS dialect and contains 494 pages (SHA-256
`579b553f499800713bdbbc3a82be6065db8611a050b585c011a29ebd8533c9ad`). The full sequence is
loaded and each page is exercised through SVG conversion. Cover, dense text/table,
and graphics pages are the representative rendering checks; successful conversion
is not a pixel-equivalence claim for every page.

Three additional [Microsoft WPF test inputs](../OfficeIMO.Xps.Tests/Fixtures/MicrosoftWpf/SOURCE.md)
qualify native page insertion/reopening with obfuscated fonts, print tickets and
thumbnail preservation. The Word/MXDC, document-structure input and printing pages
render against GhostXPS 10.08.0 at 96 DPI with mean absolute RGB differences of
3.38/255, 0.34/255 and 0.68/255 respectively. Visual checks retain text, color and
placement; glyph and edge rasterization differ. GhostXPS renders the original and
OfficeIMO-saved packages pixel-identically. These are single-page Microsoft
XPS documents with atomic storage; the structure-test input contains no authored
structure part. They do not qualify multi-document stories or OpenXPS producers.

Generated OpenXPS sequence, fixed-document, and fixed-page markup is checked against
[Ecma's OpenXPS schemas](https://ecma-international.org/wp-content/uploads/OpenXPS-WC3-Schemas.zip).
MuPDF/PyMuPDF 1.26.5 independently opens a generated Microsoft XPS document, extracts
its Unicode text, and renders its font/image placements. Its handling of this
OpenXPS package selects a generic ZIP reader, so it does not establish independent
OpenXPS rendering acceptance. These tools are isolated validation tools and are
not product dependencies or ordinary build requirements.

GhostXPS 10.08.0 independently renders generated Microsoft XPS and OpenXPS packages,
both atomic and interleaved, after document/page restructuring. The two simple
vector pages match managed raster output exactly at 96 DPI. A separate two-page
Ghostscript xpswrite document is loaded, rendered, saved, and reopened in GhostXPS;
its managed render differs by less than 0.05/255 mean channel error. This adds an
independent Microsoft-XPS producer and a representative OpenXPS consumer check;
it does not qualify an independent OpenXPS producer corpus.

A generated 16-case brush corpus is compared at 96 DPI against MuPDF's native XPS
renderer, including non-zero tile origins, all flip modes, transformed patterns,
linear-gradient transforms, and solid/gradient/tiled alpha masks. Fourteen cases agree
within small rasterization differences. The overlapping visual-brush mask and masked-group cases disagree
with MuPDF's native XPS renderer: OfficeIMO follows ECMA-388 section 18.5's isolated
composition rule, verified by opacity arithmetic and independently rendered PDF
output. These checks qualify representative cases, not every combination of native
brushes, fonts, transforms, and effects. Managed PDF readback retains the qualified
masked Canvas opacity, including overlapping children; ordinary isolated RGB
Form groups compose before invocation opacity is applied.

Sideways text follows ECMA-388 §12.1.6 for TrueType outlines: top-center origins,
vertical advances, run-relative offsets, and rotation before the page transform.
Metric-table tests cover compressed vertical metrics and both fallback sources.
A generated corpus checks horizontal and rotated/clipped runs against independent
fontTools metric extraction and MuPDF rendering of equivalent horizontal glyphs.
Direct MuPDF 1.26.5 sideways rendering agrees for the unmodified fallback fonts,
but its fixed font-ascender origin disagrees with per-glyph vertical bearings and
distinct OS/2 origins. GhostXPS 10.08.0 renders all eight native sideways cases with
matching placement, including these metrics and rotated/clipped runs. Raster
comparisons retain font rasterization differences: worst mean channel error is
1.82/255, with fewer than 1.7% of channels differing by more than 10/255.

ICC tests reuse independently generated LittleCMS reference swatches for two RGB
matrix profiles, an RGB v4 LUT profile, and a CMYK LUT profile. GhostXPS 10.08.0
renders the same 24 swatches in both dialects within one 8-bit channel value when
configured with matching relative intent and black-point compensation disabled.
TIFF placement and colors are checked in both dialects; upscaled boundaries retain
interpolation differences between the managed and independent renderers.

Eight additional image/visual stroke cases cover both dialects, non-tiled brush
coverage, transformed dashes and brush opacity. GhostXPS comparisons have a worst
mean channel difference of 0.68/255 at 96 DPI; one transformed translucent PDF is
also independently rendered. Edge antialiasing and image interpolation differ.

Style simulation checks cover the required two-percent-em default advance
increase, explicit advance overrides, horizontal/sideways italic baseline origins,
translucent overlapping glyphs, gradient fills and glyph opacity masks. Six solid-fill
single-glyph cases agree in placement with GhostXPS 10.08.0; raster differences
remain (worst mean channel error 1.87/255). Two gradient cases also verify brush
placement, but GhostXPS omits bold outline widening for non-solid glyph brushes;
that difference is not used as a fidelity target. Independent PDF rendering confirms
the gradient-filled bold outline. Bold fill is applied once through a
coverage mask in a local viewport; repeated small runs and non-tiled strokes do
not allocate a page-sized coverage layer each. These style checks do not qualify native font hinting.

The open qualification and rendering work belongs in [the roadmap](../Docs/ROADMAP.md#xpsopenxps).

ICC image qualification uses 24 independently encoded Pillow/LittleCMS fixtures,
covering associated/embedded RGB, gray, CMYK and alpha combinations. Both dialects
match the independent reference channels within two values. Of 48 GhostXPS
10.08.0 comparisons, 42 match within one channel value; its split JPEG-profile
reader and gray-alpha TIFF paths do not qualify the other six. All 48 exported
PDF swatches match the managed raster exactly when independently rendered at
96 DPI. Four nonuniform baseline/progressive JPEG checks retain EXIF orientation
with either profile association method; independent PDF/reference images confirm
placement, with raster interpolation differences at the color boundary. These
checks are not a broad photographic image corpus.

Fallback checks cover malformed and channel-incompatible associated profiles with
embedded RGB PNG/JPEG/TIFF and CMYK JPEG/TIFF profiles. They preserve the qualified
reference colors and keep strict errors when no usable profile remains. Reusing a
discarded image profile as ContextColor still reports its unsupported color. Both
dialects reject unprofiled CMYK JPEG/TIFF resources explicitly.

Searchable PDF projection is checked with native Unicode clusters, whitespace,
ligatures, surrogate pairs, right-to-left advances, sideways glyphs, affine
transforms, explicit offsets and blank pages. Poppler independently extracts the
expected text from both dialects; Ghostscript renders the generated vector PDFs
with a mean channel difference of 0.71/255 from the managed raster on the text
fixture. Interactive selection/highlighting in independent viewers remains
unqualified. Text follows source markup order and retains clipped/transparent
source content; this is not a redaction or accessibility reconstruction contract.

Native DocumentStructure edits have generated coverage in both dialects for
insertion, reordering, transfer, removal, repeated pages and atomic rejection of
malformed known metadata. The independent Ecma document reopens with 495 pages
after insertion and retains all 695 outline entries. It has no story references.
Story-page remapping follows ECMA-388 section 16.1.1.6's payload-global prose;
the adjacent attribute table describes document-local ordering. Independent
multi-document story fixtures are needed to resolve that interoperability ambiguity.

StoryFragments have generated lifecycle coverage in both dialects for reading order,
continuation, list markers, table-cell spans, repeated page occurrences, shared-part
validation, cancellation and atomic edits. ECMA-388 example 16-5 reconstructs its
three-row table from two two-row fragments. An independently produced
[Microsoft WPF sample](https://github.com/microsoft/WPF-Samples/blob/811d01e95c8c929e68539d698d0a0609e94fd185/Documents/Fixed%20Documents/DocumentStructure/content/spec_wiithstructure.xps)
resolves all six fragments on its two pages: body sections/tables, headers and
footers. Adding explicit body story addresses, saving and reopening retains a
complete logical reconstruction. The sample predates the final XPS specification;
it does not qualify independently produced OpenXPS or multi-document story addresses.
Microsoft resource-key namespaces are supported alongside the existing XAML key
spelling; legacy Microsoft image-brush `Stretch="Fill"` uses the native fill mapping.

Authored Path and Canvas accessibility descriptions have generated coverage in
both dialects for native save/reopen, SVG title/desc metadata, Reader JSON assets
and PDF Figure alternative text. A 24-case name-only, HelpText-only and combined
description matrix covers structured and unstructured pages. PyMuPDF independently
checks the PDF structure dictionaries and confirms that descriptions add no search
text. Managed output, independently rendered SVG/PDF and GhostXPS agree at all
1,560 sampled interior/background pixels within one channel level. This qualifies
the generated metadata projection, not independent producer accessibility or
PDF/UA conformance.

PDF navigation has generated coverage in both dialects for forward and same-page
links, repeated page references, page moves, percent-encoded targets, nested
transforms, clipping and outline hierarchy. An independent pypdf inspection
confirms destination pages and coordinates, Unicode outline titles and child URI
actions in both dialects. Ghostscript renders the navigation fixture successfully.
This does not qualify interactive navigation in every PDF viewer or reconstruct
StoryFragments reading order and PDF accessibility tags.

Radial-gradient qualification covers 48 generated cases in both dialects: rotation,
shear, reflection, translucent stops, offset interior foci and strokes with Pad,
Repeat and Reflect spread. GhostXPS and Ghostscript independently render the native
packages and PDFs. Worst mean channel errors at 96 DPI are 4.76/255 for native XPS
and 1.02/255 for PDF; repeat boundaries and stroke edges have the largest raster
differences. Managed tests check analytic color samples, SVG round trips, opaque
PDF reimport, opacity/clone retention and rejection at the expansion limit. These
cases do not qualify all focal positions or every producer's gradient conventions.

Native Pad boundary/exterior qualification adds generated XPS/OpenXPS cases for
elliptical fields, rotation, shear, reflection, strokes and translucent endpoints.
Core reverses the circle sequence to select the smallest containing ellipse;
PDF retains vector shading and explicitly paints the outside endpoint color and
alpha. Independent PDF consumers exercise both cone paint and outside paint.
GhostXPS agrees on qualified interior samples but leaves some outside regions
unpainted; that difference is retained in the evidence. Direct `ToSvg`, SVG image
exports and shared Drawing SVG exports preserve native Pad fields through SVG 2 shrinking-circle
patterns with separately composed color and alpha. These exports retain vector
paint and require an SVG 2 consumer; the shared SVG importer accepts their shrinking-circle
Pad representation. Direct `ToSvg` preserves its native glyph, mask,
and VisualBrush projection paths. A 40-case direct-SVG browser comparison across
both dialects covers fills, strokes, affine visual/brush transforms, opacity masks
and VisualBrush content; full-rectangle pixels differ by at most 1/255 per channel
and the maximum whole-page mean difference is 0.145/255. Sixteen additional normal
and simulated-bold glyph cases retain the field with a maximum whole-page mean
difference of 0.801/255, including text-edge rasterization differences. A 20-case Drawing browser comparison covers boundary/exterior
foci, opaque/translucent stops, inset paths, reflected/sheared placement, strokes and
markers. Full-rectangle pixels differ from managed output by at most 1/255 per
channel; maximum whole-page mean difference is 0.285/255, with geometry-edge
rasterization differences retained. Native radial colors and alpha interpolate
consistently across rectangle, path and stroke rendering. Managed opaque PDF
readback retains shrinking elliptical fields with a point end. A 108-case SVG import/browser
comparison covers 32 standalone shrinking fields, the 56 direct native SVG
exports above and 20 Drawing SVG exports. Another 22 native tiled-brush cases
compare native PNG output with emitted SVG in the browser. Plain fills differ from browser rendering by at most 2/255 per
channel; the maximum whole-page mean difference is 1.024/255 including glyph,
stroke, transformed-pattern and cone-edge rasterization. Pattern tile origins,
stroke coverage offsets and clipped overflow are retained through import. Arbitrary
photographic or producer coverage and PDF/UA are not qualified by these cases.

Native exterior Repeat/Reflect uses the same smallest-containing-ellipse rule.
Core bounds the required cycles over the painted region and retains at most 256
expanded stops; Reflect uses offset zero outside the cone, while Repeat uses
offset one. Focus positions exactly on the end ellipse boundary are supported when the
conservative painted bounds lie strictly inside the focus tangent half-plane
and fit the same stop budget. Bounds touching or crossing the tangent can require
an unbounded cycle count. PDF retains these fields as vector function-based
shadings with separate alpha masks. Drawing, raster and SVG retain an
explicit spread mode and sample the original field without stop expansion. Near-boundary
exterior fields use an additional finite half-plane bound to avoid excessive
expansion from coordinate rounding.
A 32-case PDF comparison covers both dialects, both spread modes, alpha,
linear-light RGB and affine transforms across the tangent. Managed reopening at
96 dpi has a maximum whole-image mean channel difference of 0.054/255 against
native XPS sampling. Ghostscript independently renders the vector files; its
maximum mean difference is 2.218/255, including dense cycle sampling differences.
MuPDF 1.28.2 also renders the 32 periodic RGB cases, with a maximum whole-image
mean channel difference of 3.700/255 against managed PDF readback. Near-tangent
cycle sampling still differs between renderers. These are generated fixtures,
not independent-producer or native Windows proof.
Explicit print-condition conversion reuses the PDF engine's ICC gradient sampler,
retaining vector CMYK component functions and scalar alpha. Caller-supplied
profiles, rendering intent and black-preservation settings follow the existing
PDF print contract. The converted sample count is bounded at 4,096.
Managed readback retains the PDF reader's output-intent/transparency diagnostic;
this conversion does not qualify print soft-proof equivalence or PDF/X conformance.
The 36 ordinary generated cases cover both dialects, fill/stroke, affine
transforms and stop alpha. Managed sampling and Ghostscript PDF rendering track
the analytic field in these cases, while MuPDF and GhostXPS retain visible
consumer differences. Separate near-boundary numerical regressions cover large
radii, shrinking-circle roots, PDF focal coordinates and nonempty endpoint/alpha
Form bounds. Managed rendering and opaque PDF readback retain these extreme
fields; both MuPDF and Ghostscript have numerical or raster differences in the
large-radius stress cases. This evidence qualifies the bounded field and emitted
vector geometry, not every consumer's raster output or independent native Windows
behavior. Direct `ToSvg`, Drawing and SVG image exports use vector pattern
composition after bounded spread expansion. The direct route shares Core's cycle
bounds, preserves authored color interpolation, and caps expanded stops at 256.
When finite expansion is unavailable, direct `ToSvg` uses SVG 2 repeating
shrinking-circle fields without expanding stops. Separate color and alpha fields
retain the native outside-cone endpoint for both Repeat and Reflect. A 32-case
browser comparison covers both dialects, both spread modes, stop alpha,
linearRGB and affine transforms across the tangent boundary. The maximum mean
channel difference against the analytical native field is 0.139/255; maximum
channel difference is 6/255 at high-frequency transformed cycle edges. This
qualifies direct SVG output; native Windows qualification remains open.
Shared SVG import, Drawing and raster preserve these fields and ordinary
shrinking-circle Repeat/Reflect gradients. A 96-pair comparison covers native
PNG, imported SVG raster output and Drawing SVG for the same 32 cases; maximum
mean channel error is 0.161/255. Exact Repeat seams can differ by 255/255 when
numerical rounding chooses opposite sides of a discontinuity. PDF still requires
finite expansion.

For ordinary SVG point foci exactly on the end circle, repeating gradients use
the offset-weighted average color and alpha outside the tangent half-plane, as
defined by [SVG 2](https://www.w3.org/TR/SVG2/pservers.html#RadialGradientNotes).
This compatibility rule is separate from native XPS endpoint paint. The checked
browser paints that ordinary outside region transparent instead; browser agreement
is not claimed for the original SVG compatibility case. The composed re-export
matches managed output within 1/255 per channel. The full 110-pair check contains
108 comparisons with maximum mean error 0.235/255 and these two explicitly
separated original-consumer differences, including four ordinary shrinking fields
and two public Drawing interior-alpha fields.

An 84-case direct-SVG browser comparison covers 36 boundary fields and 48 exterior
fields across both dialects, Repeat/Reflect, stop alpha, transforms, strokes and
non-endpoint stops and mixed filled/unfilled stroke figures. The 76 sRGB cases have a maximum whole-page mean difference
of 1.366/255 against managed PNG output, including geometry and cycle edges.
All 84 SVG files reimport without unsupported-feature diagnostics. Native
`ScRgbLinearInterpolation` is retained through shared Drawing, raster, SVG and
PDF output. A 184-pair browser comparison includes direct and Drawing SVG for
all 84 cases, plus managed PDF readback and Ghostscript for eight linearRGB
cases. Those eight cases have maximum mean channel differences of 0.071/255
for SVG/PNG, 0.099/255 for managed PDF readback and 0.549/255 for Ghostscript,
including cone and cycle-edge rasterization. Direct and Drawing SVG differ
from managed PNG by at most one channel value in these linearRGB cases; managed
PDF readback differs by at most two. PDF uses calibrated RGB shading; color-stop
alpha remains independent of color interpolation. This qualifies these generated
fields, not arbitrary native-producer coverage.


Boundary Repeat/Reflect qualification adds 36 generated cases across both dialects,
fill, stroke, translucent stops, shear and reflection. Managed samples differ from
the analytic field by at most 1.09/255; opaque PDF readback differs by at most
0.51/255. Ghostscript and GhostXPS render the native vector fields, with maximum
sample differences of 9.16/255 and 14.03/255 respectively at 96 DPI. Raster edge
and interpolation differences remain; native Windows acceptance is not established.
The shared renderer preserves radial fill, gradient stroke and marker coordinates
through shape transforms, including paths inset within a declared canvas. PDF readback accepts up to 256 stitched child functions,
with its separate recursion, breakpoint and evaluation limits still enforced.

Stroke qualification includes 48 generated XPS/OpenXPS cases and 16 focused rendering
regressions covering clipped miters,
asymmetric and triangle caps, separate dash caps, overlapping translucent dashes,
closed seams, curve/affine placement, fill/stroke suppression, and gradient, image
and visual-brush paint. Core constructs the stroke as one nonzero union before the
native transform, so overlapping pieces do not compound the brush opacity. Curve
approximation accounts for native and visual-brush transforms; subdivision, expanded
points and SVG output remain bounded.
Open figures retain their caps when their endpoints coincide. Painted dashes retain
authored degenerate vertices, including closed seams. Small line/arc coordinates
remain distinct before affine magnification; equivalent transformed and untransformed
managed fixtures have the same coverage.

GhostXPS/Ghostscript and MuPDF render the same generated fixtures at 96 DPI. Their
rendered evidence is supplemented by analytic managed checks. Both XPS engines omit
fully degenerate strokes and apply authored line caps at internal breaks created by
unstroked segments; the managed checks follow ECMA-388 18.6.5, 18.6.8 and 18.6.10 for
those cases. Exact dash-boundary caps, dashed degenerate joins, leading degenerate
closed seams and magnified short geometry also differ between consumers. These
discrepancies remain visible qualification gaps, requiring
a native Windows consumer or another producer/consumer fixture; they are not treated
as proof that every native stroke case is independently qualified.

PDF semantic mapping is exercised with generated native structures in both dialects
and Microsoft's two-page WPF structured sample. pypdf independently resolves every
marked-content reference and ParentTree entry in these PDFs. The generated stories
retain the declared page-2-before-page-1 order, labels, empty cells and spans.
PyMuPDF 1.28.2 renders tagged and untagged exports pixel-identically at 96 DPI on
all six pages. This proves the representative structure and paint-preservation
contracts; it does not establish a broader OpenXPS producer corpus or PDF/UA
conformance. The PDF bridge uses the same native structure limits and rejects
multiple semantic owners for the same glyph or graphic paint.

## PNG declarations without ICC

Native integer PNG samples without a usable ICC profile use the sRGB defaults
required by ECMA-388 15.3.7, Table 15-3. `gAMA` and `cHRM` do not override those
native defaults. Core decodes the channel samples, and the native image path
re-encodes them before SVG/PDF projection so another consumer cannot apply PNG
calibration to the resulting paint. A usable associated or embedded ICC profile
retains the native profile precedence described above.

PNG resources with APNG chunks use the static `IDAT` image, including when it
differs from animation frame zero. Native SVG/PDF/raster conversion normalizes
that static image for ordinary, gamma-declared and ICC-managed resources. It
does not play animation. This normalization uses the same four-million-pixel
limit as native color conversion; malformed containers and exceeded limits are
reported instead of selecting an animation frame.

Nine qualified generated PNG inputs cover RGB, grayscale, indexed color and alpha, with gamma-only
and gamma/chromaticity declarations. Canonical sRGB `cICP` is accepted; non-sRGB
`cICP`, other unsupported color declarations and noncanonical TIFF colorimetry
remain diagnosed. The native four-million-pixel decode limit and cancellation
policy still apply. This does not qualify HDR, general PNG color management,
a broad photographic corpus or all non-ICC image colorimetry. Raw Core raster
decoding continues to return channel samples without color calibration.

## JPEG XR image resources

Image brushes accept `image/jxr` and the legacy `image/vnd.ms-photo` content type. Core decodes single-image tagged containers with unsigned eight-bit gray/RGB/BGR/BGRA, unsigned sixteen-bit gray/RGB/RGBA, and finite sixteen/thirty-two-bit fixed-point or floating-point gray/RGB/RGBA samples, 4:4:4/4:2:2/4:2:0 chroma with defined sampling-grid centering, spatial/frequency packets, lossy/lossless quantization, overlap filtering, tiled images, reduced subbands, and straight or premultiplied alpha. Unsigned eight/sixteen-bit CMYK and CMYKDirect images require an associated or embedded four-component ICC profile. Three-to-eight-channel unsigned eight/sixteen-bit images require a matching associated or embedded ICC profile. Associated or embedded RGB/gray/CMYK/multichannel ICC profiles use the existing image color pipeline at source precision before rounding to the rendering buffer. SVG carries normalized PNG pixels, and raster/PDF conversion uses the same decoded image.

The checked-in reference corpus covers independently encoded eight/sixteen-bit images, sample shifts, and premultiplied-alpha variants. Color and alpha retain source precision through unassociation and ICC conversion before rounding to the eight-bit rendering buffer. Without a usable ICC profile, fixed-point and floating-point samples use linear scRGB and are converted to sRGB, clipping colors outside the SDR output gamut. NaN and infinity samples are rejected without partial pixels. Focused XPS tests exercise both package dialects and raster, SVG, and PDF-reader pixels. Interleaved alpha supports the same or fewer frequency bands than the primary plane. All six reduced-alpha band combinations have unsigned eight/sixteen-bit regression coverage in both packet orders: spatial pixels match the independent ITU decoder, while frequency pixels match equivalent spatial encodings. Direct native frequency-order qualification remains open because the independent decoders fail or disagree on these mixed-band streams. Multiple image directories remain unsupported. This is a bounded rendering contract, not complete JPEG XR or OpenXPS consumer conformance.

CMYK JPEG XR qualification includes 96 independent encodings with 768,768 exact source samples across CMYK/CMYKDirect, unsigned eight/sixteen-bit, spatial/frequency packets, quantization, hard tiles and separate/interleaved alpha. Both XPS dialects preserve ICC-normalized pixels and PDF-reader alpha. The Microsoft comparison decoder accepts the 48 ordinary CMYK cases, agrees on all color samples, and differs by one alpha level in four lossy eight-bit interleaved-alpha cases; it rejects CMYKDirect. The ITU reference fixtures correct container identifiers and separate-alpha lengths and pad the input TIFF alpha strip where required by that encoder. These corrections do not alter encoded pixel packets.

N-channel qualification covers three through eight unsigned 8/16-bit color channels, spatial/frequency packets, and interleaved or separate alpha. The 60 fixtures match 422,730 ITU-decoded source samples exactly; 48 are direct reference encodings and 12 combine independently encoded primary and grayscale alpha streams. The latter avoid a reference-encoder assertion when producing separate N-channel alpha; the assembled containers are independently decoded again. The shared ICC engine matches 153 LittleCMS swatches for `3CLR`–`8CLR` LUT8/LUT16 and variable-grid `mAB` profiles within two 8-bit levels. Both XPS dialects preserve normalized SVG pixels and PDF-reader alpha. Four representative rendered image pairs match reference PNG placements exactly. The Microsoft comparison decoder agrees exactly on 23 of the 48 direct encodings; it differs on interleaved alpha and one eight-channel frequency case. Broader native N-channel interoperability remains unqualified.
