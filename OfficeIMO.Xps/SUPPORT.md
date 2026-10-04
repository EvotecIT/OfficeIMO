# XPS/OpenXPS support

The native package lifecycle and the rendering profile are separate contracts.
Preserving an unsupported native element does not mean that it can be rendered.

| Operation | Supported contract | Boundary |
| --- | --- | --- |
| Read | Microsoft XPS and ECMA-388 OpenXPS; OPC relationships/content types; multiple fixed documents; ordered page references; UTF-8/UTF-16 XML | Conventional ZIP parts; interleaved OPC piece streams and protected packages are not supported |
| Create | Both dialects; pages, vector paths, embedded fonts, Unicode glyph runs, PNG/JPEG image placement | Typed creation is a bounded fixed-page profile, not a complete schema object model |
| Edit | Detached native page XML with explicit replacement; new resource parts | Loaded document-sequence restructuring and resource replacement are not exposed |
| Save | Original dialect; native page content and opaque parts retained; direct required-resource relationships emitted; deterministic ZIP output on the same runtime | ZIP metadata/XML bytes may change; no dialect conversion; signed packages cannot be rewritten |
| Text extraction | UnicodeString runs in markup order | No inferred reading order, paragraphs, or glyph-ID-to-Unicode reconstruction |
| Paths | Abbreviated geometry, fill rules, explicit path figures/segments, fills, strokes, dashes, matrix transforms, clipping | Per-segment fill/stroke suppression, asymmetric/triangle or separate dash caps, and over-limit clipped miters (including the native degenerate-segment rule) are diagnosed |
| Text rendering | Embedded TrueType programs/collections, obfuscation, explicit glyph IDs, cluster mappings, advances/offsets, horizontal bidi direction | Outlined output; sideways glyphs, style simulations, and unsupported font programs are diagnosed |
| Brushes | Hex/scRGB solid colors; linear/radial gradients; scoped and external package resource dictionaries; PNG/JPEG and visual brushes with absolute viewbox/viewport mapping, matrix transforms, Tile/FlipX/FlipY/FlipXY repetition, and alpha opacity masks | Non-tiled brush strokes, ICC ContextColor, color-converted images, and TIFF/JPEG-XR rendering are diagnosed |
| Navigation | Safe web/mail links; page/document/sequence named targets with scoped first-occurrence lookup; sequence page numbers projected into SVG filenames | Non-page unresolved and unsafe destinations are diagnosed; document navigation is not a PDF preservation contract |
| Gradient transforms | Affine transforms retained in SVG; affine linear gradients and axis-aligned scaled/translated radial gradients convert through Core | Rotated/sheared radial gradients and non-Pad radial spread reject drawing/image/PDF conversion |
| SVG | Self-contained images and glyph outlines; strict by default; explicit partial result with diagnostics | Unknown markup/attributes are diagnosed; no claim of complete XPS consumer conformance |
| Drawing/images | Existing managed Core scene and image exporters | Shared viewport, element, geometry, raster, and codec limits still apply; any reported SVG import loss rejects conversion |
| PDF | Optional thin bridge to the existing PDF engine, retaining page dimensions, bounded vector tile expansion, and native alpha-mask Forms | Vector outlines rather than searchable text; no print-ticket/structure/signature migration |
| Security | Package-local resource resolution; no external fetch; DTD prohibition; shared backing for repeated page parts; bounded ZIP/XML/page and expanded SVG node/character/resource-binding growth; cooperative cancellation; atomic path saves | Inspection does not authenticate signatures or make arbitrary native documents trusted |

## Qualification

The focused tests exercise both dialects, package reopening and native edits,
opaque-part preservation, deterministic saves, fonts/glyph positioning, image
placement, gradients, external resource dictionaries, hostile inputs, cancellation,
and preservation of the destination on a rejected signed-package save.

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

A generated 16-case brush corpus is compared at 96 DPI against MuPDF's native XPS
renderer, including non-zero tile origins, all flip modes, transformed patterns,
linear-gradient transforms, and solid/gradient/tiled alpha masks. Fourteen cases agree
within small rasterization differences. The overlapping visual-brush mask and masked-group cases disagree
with MuPDF's native XPS renderer: OfficeIMO follows ECMA-388 section 18.5's isolated
composition rule, verified by opacity arithmetic and independently rendered PDF
output. These checks qualify representative cases, not every combination of native
brushes, fonts, transforms, and effects.

The open qualification and rendering work belongs in [the roadmap](../Docs/ROADMAP.md#xpsopenxps).
