# XPS/OpenXPS support

The native package lifecycle and the rendering profile are separate contracts.
Preserving an unsupported native element does not mean that it can be rendered.

| Operation | Supported contract | Boundary |
| --- | --- | --- |
| Read | Microsoft XPS and ECMA-388 OpenXPS; OPC relationships/content types; multiple fixed documents; ordered page references; UTF-8/UTF-16 XML; bounded interleaved OPC piece assembly | Materialized loading, not progressive streaming; protected packages are not supported |
| Create | Both dialects; pages, vector paths, embedded fonts, Unicode glyph runs, PNG/JPEG/TIFF image placement | Typed creation is a bounded fixed-page profile, not a complete schema object model |
| Edit | Detached native page XML; loaded document/page insertion, reordering, transfer and reference removal; shared backing for repeated references; encoded resource replacement; owning APIs for relationship-owned DocumentStructure and StoryFragments; atomic combined page/fragment edits | Removed parts/resources remain preserved; story references to removed pages and resulting empty stories are removed; dangling known name/story addresses reject edits; unknown semantic extensions remain opaque |
| Save | Original dialect; native page content and opaque parts retained; required-resource relationships emitted, including profiles and transitive dictionary resources; deterministic ZIP output on the same runtime | ZIP metadata/XML bytes may change; interleaved storage is normalized to atomic parts; no dialect conversion; signed packages cannot be rewritten |
| Text extraction | UnicodeString runs in markup order | No inferred reading order, paragraphs, or glyph-ID-to-Unicode reconstruction |
| Native logical structure | Relationship-owned StoryFragments; named page/Canvas/Path/Glyphs references; DocumentStructure story-reference order; continued paragraphs, sections, lists, figures and tables; StoryBreak boundaries; list markers and cell spans | No inferred structure on unstructured pages; unknown extensions and unresolved content produce diagnostics; missing Unicode is not reconstructed; one million work units and 16 million resolved text characters bound each read |
| Paths | Abbreviated geometry, fill rules, explicit figures/segments with fill/stroke suppression, dashes with separate endpoint/dash caps, triangle caps, clipped miters and the degenerate-segment limit, matrix transforms and clipping | Extended strokes use bounded adaptive vector outlines; native Windows confirmation remains for degenerate and mixed-segment cap rules where the independent engines disagree with the specification |
| Text rendering | Embedded TrueType programs/collections, obfuscation, explicit glyph IDs, cluster mappings, advances/offsets, horizontal bidi, bold/italic style simulation, sideways top-center positioning with vertical metrics or OS/2/hhea fallbacks | Outlined output; unsupported font programs are diagnosed; sideways runs require even BidiLevel |
| Brushes | Hex/scRGB and ICC ContextColor solids/gradient stops; linear/radial gradients; scoped and external package resource dictionaries; ICC-managed PNG/JPEG/TIFF and visual brushes with absolute viewbox/viewport mapping, matrix transforms, Tile/FlipX/FlipY/FlipXY repetition, non-tiled fills/strokes, and alpha opacity masks | Unsupported image/profile channel combinations, non-ICC colorimetry, unsupported TIFF encodings, and JPEG-XR rendering are diagnosed |
| Navigation | Safe web/mail links; page/document/sequence named targets with scoped first-occurrence lookup; sequence page numbers projected into SVG filenames; PDF links, named destinations and DocumentStructure outlines | Non-page unresolved and unsafe destinations are diagnosed; known fixed-page destinations follow structural moves; links to removed pages are unresolved; PDF link hit areas are rectangles and path destination positions use conservative geometry bounds |
| Gradient transforms | Affine linear and radial gradients convert through Core, including rotation, shear and reflection; bounded radial Repeat/Reflect expansion retains vector PDF shading | Radial spread requires a point focus strictly inside the end ellipse and at most 256 expanded stops; boundary/exterior focal behavior is not qualified |
| SVG | Self-contained images and glyph outlines; strict by default; explicit partial result with diagnostics | Unknown markup/attributes are diagnosed; no claim of complete XPS consumer conformance |
| Drawing/images | Existing managed Core scene and image exporters | Shared viewport, element, geometry, raster, and codec limits still apply; any reported SVG import loss rejects conversion |
| PDF | Optional thin bridge to the existing PDF engine, retaining page dimensions, bounded vector tile expansion, and native alpha-mask Forms | Searchable native Unicode clusters alongside vector outlines; no print-ticket/accessibility-structure/signature migration; source markup order, not reconstructed logical reading order; clipped/transparent source text remains searchable |
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
and unsupported non-ICC color metadata are diagnosed. Default CMYK/SWOP and
incompatible-profile fallback behavior are not qualified.

## Qualification

The focused tests exercise both dialects, package reopening and native edits,
opaque-part preservation, deterministic saves, fonts/glyph positioning, image
placement, gradients, external resource dictionaries, hostile inputs, cancellation,
and preservation of the destination on a rejected signed-package save. Native-editing
checks cover repeated document/page references, reference metadata and link-target
preservation, atomic rejection at package/page limits, and resource replacement.
Generated piece fixtures cover interleaved metadata and resources, non-sequential
ZIP ordering, missing/duplicate/ambiguous pieces, and aggregate part bounds.
Independent-producer qualification of interleaved storage remains open.

Independent input: [Ecma's published ECMA-388 XPS document](https://ecma-international.org/wp-content/uploads/ECMA-388.xps),
which uses the Microsoft XPS dialect and contains 494 pages (SHA-256
`579b553f499800713bdbbc3a82be6065db8611a050b585c011a29ebd8533c9ad`). The full sequence is
loaded and each page is exercised through SVG conversion. Cover, dense text/table,
and graphics pages are the representative rendering checks; successful conversion
is not a pixel-equivalence claim for every page.

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
brushes, fonts, transforms, and effects.

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
