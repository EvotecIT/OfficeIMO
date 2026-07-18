# OfficeIMO To-Image Conversion Assessment

Date: 2026-07-18

## Outcome

OfficeIMO now has one dependency-free image-export contract across Drawing, Excel, Word, PowerPoint, HTML, OneNote, Visio, and PDF. `OfficeIMO.Drawing` owns encoded format and dimension identity, options, fluent mechanics, allocation safety, raster/SVG encoding, bounded source decoding, caller-codec fallback, filenames, and shared diagnostic codes. Format packages keep only their source semantics, selection, layout, and projection.

This work fixed a mislabeled-format correctness bug, removed duplicated option and raster-limit state, made raster limits pre-allocation across every main converter, added first-party PDF PNG/JPEG/TIFF/SVG/WebP export, added bounded TIFF and OfficeIMO-WebP source decoding, made HTML fluent async rendering real, and introduced one paged-image bridge for all source-to-PDF adapters. It adds no runtime dependency.

## Conversion Matrix

| Source | Single image | Batch images | Output formats | Contract and limits |
| --- | --- | --- | --- | --- |
| `OfficeDrawing` | Yes | Adapter-owned | PNG, JPEG, TIFF, SVG, WebP | Shared encoder, validated result/options/builders, raster planner, source decoders |
| Excel range | Yes | No | PNG, JPEG, TIFF, SVG, WebP | Mature range coverage and visual baselines |
| Excel worksheet/workbook | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Used range, explicit range, print area, manual-page-break slices |
| PowerPoint slide/presentation | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Fixed-layout projection; representative authored fixture coverage |
| Word document page/range | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | OfficeIMO-estimated pagination, not Microsoft Word pagination |
| HTML continuous/paged render | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Resource-aware direct and fluent async paths |
| OneNote page/section/notebook | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Ink, math, pictures, selection, and shared raster safety |
| Visio page/document | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Native Visio projection with shared results, encoders, safety, and batch filenames |
| Loaded PDF page/document | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | First-party page-to-Drawing projection, DPI/thumbnails, ordered page selection, capability diagnostics |
| Markdown, RTF, AsciiDoc, LaTeX, and other source-to-PDF adapters | Through PDF result | Yes | PNG, JPEG, TIFF, SVG, WebP | `PdfDocumentConversionResult.ToImages()` is the single paged adapter and preserves source warnings |
| OpenDocument, EPUB, email | No direct visual adapter | No | None directly | A direct adapter needs a deliberate visual contract; a format-label wrapper would be misleading |

## Issues Found And Addressed

1. **Declared HTML formats did not match their bytes.** JPEG, TIFF, and WebP requests returned PNG payloads with another label. HTML now encodes the requested format, and `OfficeImageExportResult` rejects any future declared-format or encoded-dimension mismatch.
2. **Raster safety was fragmented and often late.** OneNote and Visio had private limit policies, HTML used dimension caps, and Excel, Word, and PowerPoint could allocate before encoder rejection. `OfficeRasterExportPlanner` now combines caller, renderer, and encoder limits before allocation. It either emits `IMAGE_RASTER_SCALE_REDUCED` or throws typed `OfficeImageExportLimitException`.
3. **Shared options were incompletely cloned or redeclared.** Each package now inherits and snapshots the Drawing-owned scale, background, pixel budget, overflow policy, caller codec, and raster encoding. Duplicate OneNote and Visio option properties were removed.
4. **PDF was documented as missing even though a substantial first-party page painter already existed.** The existing `PdfReadPage.ToDrawing()` projection is now exposed through the canonical five-format result/builder contract, with ordered batch selection, DPI, thumbnails, cancellation, source-image fallback, and shared safety.
5. **Source decode breadth lagged output breadth.** Drawing now reads bounded baseline chunky RGB/RGBA TIFF strips using no compression or PackBits and the literal-lossless VP8L subset emitted by OfficeIMO. It does not pretend to be a general TIFF or WebP decoder; unsupported variants use `ImageCodec` or a visible fallback diagnostic.
6. **Generic async builders always rendered synchronously.** The shared builder can now accept a true asynchronous render delegate. HTML uses it for resource-aware scene construction; CPU-only in-memory projection remains synchronous and only file/stream commit is asynchronous.
7. **Source adapters risked multiplying APIs and option models.** Markdown, AsciiDoc, LaTeX, RTF, OneNote, Word, Excel, PowerPoint, and HTML already converge on `PdfDocumentConversionResult`. Its `ExportImages()` / `ToImages()` bridge now retains source conversion warnings on each page image.
8. **Visio and Drawing both contain SVG readers.** They are not interchangeable duplicate brains. Visio's embedded-SVG path supports linked images, clipping, nested opacity, and CSS needed by stencil artwork; deleting it would regress fidelity. Visio still delegates final encoding, allocation policy, results, and filenames to Drawing.
9. **Batch filenames varied by host operating system.** Shared batch export now emits bounded, unique, Windows-portable names on every host and protects reserved device names.

## Intentionally Retained Low-Level APIs

`PdfPageImageRenderer.RenderPage(...)` remains valuable as the PDF-to-`OfficeDrawing` projection used by inspection and visual comparison. The older `PdfPageRenderResult` batch also carries elapsed time, continue-on-error state, and typed PDF capability diagnostics used by OCR, destructive-crop verification, and redaction verification. It is therefore retained as a low-level operational contract, while `PdfReadPage.ToImage()` and `PdfReadDocument.ToImages()` are the canonical general export APIs.

## Remaining Product Limits

- Word output remains an OfficeIMO pagination estimate. Exact Microsoft Word pagination requires Word's layout engine and is not claimed.
- Arbitrary PDFs can contain operators, fonts, transparency, forms, annotations, and producer-specific constructs outside the current first-party page projection. The generated capability manifest and per-page diagnostics are the coverage contract.
- TIFF support is baseline chunky RGB/RGBA strips, not BigTIFF, tiled, planar, palette, CMYK, or arbitrary compression. WebP source decoding is limited to OfficeIMO's literal-lossless encoder output.
- GIF decoding uses the first frame; animation timing and multi-frame TIFF are not part of the static image result contract.
- Color-profile, ICC, EXIF-preservation, and explicit cross-format DPI metadata are separate product decisions. The current result contract guarantees encoded format, pixels/CSS dimensions, MIME type, extension, source metadata, and diagnostics.
- Representative visual evidence remains strongest for Excel and Visio. PowerPoint, Word, HTML, OneNote, and PDF need continued authored-fixture growth as new fidelity slices are implemented.

## Premium Converter Gates

The implemented contract now requires:

- encoded bytes match the declared format and dimensions;
- PNG/JPEG/TIFF/SVG/WebP share one result/options/builder model;
- raster bounds are enforced before allocation;
- names are deterministic and portable;
- unsupported source images are decoded by an explicit caller codec or rendered visibly with a stable diagnostic;
- async claims correspond to real resource or destination I/O;
- adapter diagnostics survive the image bridge;
- document-specific semantics remain in their owning package and reusable mechanics remain in Drawing;
- no new runtime dependencies are introduced.

Future fidelity work should add evidence to these contracts, not create parallel renderers or package-local copies of export mechanics.
