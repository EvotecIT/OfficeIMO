# XPS/OpenXPS support

The native package lifecycle and the rendering profile are separate contracts.
Preserving an unsupported native element does not mean that it can be rendered.

| Operation | Supported contract | Boundary |
| --- | --- | --- |
| Read | Microsoft XPS and ECMA-388 OpenXPS; OPC relationships/content types; multiple fixed documents; ordered page references; UTF-8/UTF-16 XML; bounded interleaved OPC piece assembly | Materialized loading, not progressive streaming; protected packages are not supported |
| Create | Both dialects; pages, vector paths, embedded fonts, Unicode glyph runs, PNG/JPEG/TIFF image placement | Typed creation is a bounded fixed-page profile, not a complete schema object model |
| Edit | Detached native page XML; loaded document/page insertion, reordering, transfer and reference removal; shared backing for repeated references; encoded resource replacement | Removed parts/resources remain preserved; opaque semantic metadata is not rewritten |
| Save | Original dialect; native page content and opaque parts retained; required-resource relationships emitted, including profiles and transitive dictionary resources; deterministic ZIP output on the same runtime | ZIP metadata/XML bytes may change; interleaved storage is normalized to atomic parts; no dialect conversion; signed packages cannot be rewritten |
| Text extraction | UnicodeString runs in markup order | No inferred reading order, paragraphs, or glyph-ID-to-Unicode reconstruction |
| Paths | Abbreviated geometry, fill rules, explicit path figures/segments, fills, strokes, dashes, matrix transforms, clipping | Per-segment fill/stroke suppression, asymmetric/triangle or separate dash caps, and over-limit clipped miters (including the native degenerate-segment rule) are diagnosed |
| Text rendering | Embedded TrueType programs/collections, obfuscation, explicit glyph IDs, cluster mappings, advances/offsets, horizontal bidi, bold/italic style simulation, sideways top-center positioning with vertical metrics or OS/2/hhea fallbacks | Outlined output; unsupported font programs are diagnosed; sideways runs require even BidiLevel |
| Brushes | Hex/scRGB and ICC ContextColor solids/gradient stops; linear/radial gradients; scoped and external package resource dictionaries; PNG/JPEG/TIFF and visual brushes with absolute viewbox/viewport mapping, matrix transforms, Tile/FlipX/FlipY/FlipXY repetition, non-tiled fills/strokes, and alpha opacity masks | Color-converted images, embedded TIFF color management, unsupported TIFF encodings, and JPEG-XR rendering are diagnosed |
| Navigation | Safe web/mail links; page/document/sequence named targets with scoped first-occurrence lookup; sequence page numbers projected into SVG filenames | Non-page unresolved and unsafe destinations are diagnosed; known fixed-page destinations follow structural moves; links to removed pages are unresolved; document navigation is not a PDF preservation contract |
| Gradient transforms | Affine transforms retained in SVG; affine linear gradients and axis-aligned scaled/translated radial gradients convert through Core | Rotated/sheared radial gradients and non-Pad radial spread reject drawing/image/PDF conversion |
| SVG | Self-contained images and glyph outlines; strict by default; explicit partial result with diagnostics | Unknown markup/attributes are diagnosed; no claim of complete XPS consumer conformance |
| Drawing/images | Existing managed Core scene and image exporters | Shared viewport, element, geometry, raster, and codec limits still apply; any reported SVG import loss rejects conversion |
| PDF | Optional thin bridge to the existing PDF engine, retaining page dimensions, bounded vector tile expansion, and native alpha-mask Forms | Vector outlines rather than searchable text; no print-ticket/structure/signature migration |
| Security | Package-local resource resolution; no external fetch; DTD prohibition; shared backing for repeated page parts; bounded ZIP/XML/page and expanded SVG node/character/resource-binding growth; cooperative cancellation; atomic path saves | Inspection does not authenticate signatures or make arbitrary native documents trusted |

ICC ContextColor uses Core's supported RGB, gray, CMYK and N-channel profiles,
converting to sRGB with media-relative colorimetric intent and no black-point
compensation. Alpha and channel values are clamped to the native range. Profile
parsing has a 4 MiB per-profile ceiling and a 64 MiB aggregate parser allowance.
PrintTicket color overrides are not interpreted. TIFF uses Core's managed decoder
for the first image, up to four million pixels, retaining its physical dimensions;
embedded color-management metadata is diagnosed rather than discarded.

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
not allocate a page-sized coverage layer each. These checks do not imply searchable PDF text or native font hinting.

The open qualification and rendering work belongs in [the roadmap](../Docs/ROADMAP.md#xpsopenxps).
