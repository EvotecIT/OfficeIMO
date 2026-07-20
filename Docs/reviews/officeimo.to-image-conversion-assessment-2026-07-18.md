# OfficeIMO To-Image Conversion Assessment

Date: 2026-07-18

## Outcome

OfficeIMO now has one dependency-free image-export contract across Drawing, Excel, Word, PowerPoint, HTML, OneNote, Visio, PDF, OpenDocument, EPUB, and email. `OfficeIMO.Drawing` owns format identity, raster/SVG encoding, density metadata, deterministic fonts, source-image decoding, diagnostic policy, batch budgets, cancellation, progress, save conflicts, paths, and shared renderer primitives. Each document package owns only its source semantics, selection, layout, and projection.

The work did not add an output format or a runtime dependency. It concentrated on correctness and production use: encoded bytes must match their declared format and dimensions, saves return the path actually committed, large batches can stream without retaining every payload, aggregate limits stop runaway work, and fidelity loss can be accepted or rejected through structured policy.

The raster text path also has one complex-script contract. Callers can attach an `IOfficeTextShapingProvider` and optional language hint to the shared image options used by Word, Excel, PowerPoint, HTML, OneNote, Visio, and loaded PDF pages. Without a provider, Drawing applies a bounded core-Arabic and bidirectional fallback. That fallback is useful, but it is not presented as full OpenType shaping: `IMAGE_TEXT_SHAPING_FALLBACK` records an approximation and strict policies can reject it.

## Conversion Matrix

| Source | Single image | Batch images | Visual owner | Important contract |
| --- | --- | --- | --- | --- |
| `OfficeDrawing` | Yes | Adapter-owned | Drawing | Shared encoding, validation, DPI metadata, fonts, policies, and raster limits |
| Excel range | Yes | No | Excel → Drawing | Range selection, worksheet semantics, decoded-pixel and approved visual baselines |
| Excel worksheet/workbook | Yes | Yes | Excel → Drawing | Used/explicit range, print areas, manual-page-break slices, workbook-wide budgets |
| PowerPoint slide/presentation, including binary `.ppt` import | Yes | Yes | PowerPoint → Drawing | Fixed-layout projection, authored fixtures, and LibreOffice raster references |
| Word document page/range | Yes | Yes | Word → Drawing | OfficeIMO-estimated pagination; it does not claim Microsoft Word pagination |
| HTML continuous/paged render | Yes | Yes | HTML → Drawing | Resource-aware sync/async rendering with HTML safety limits |
| Email body/pages | Yes | Yes | Email → HTML → Drawing | HTML, RTF, or text body selection; inline MIME/CID resources; message chrome |
| EPUB chapters/pages | No single-chapter shortcut | Yes | EPUB → HTML → Drawing | Retained chapter HTML/resources, selection, diagnosed text fallback |
| OneNote page/section/notebook | Yes | Yes | OneNote → Drawing | Ink, math, pictures, hierarchy selection, and shared raster safety |
| Visio page/document | Yes | Yes | Visio geometry → Drawing encoding | Native Visio projection, deterministic fonts, shared result/save/batch contracts |
| Loaded PDF page/document | Yes | Yes | PDF → Drawing | Page selection, thumbnails, target DPI, capability diagnostics |
| ODT | Yes | Yes | ODT → Word → Drawing | ODF conversion diagnostics are attached to every image |
| ODS | Through selected worksheets | Yes | ODS → Excel → Drawing | ODF conversion diagnostics and workbook-wide export budgets |
| ODP | Through selected slides | Yes | ODP → PowerPoint → Drawing | ODF conversion diagnostics and ordered presentation export |
| Markdown, RTF, AsciiDoc, LaTeX, and existing PDF adapters | Through PDF result | Yes | Source → PDF → Drawing | `PdfDocumentConversionResult.ToImages()` preserves source warnings |

All rows produce the existing PNG, JPEG, TIFF, SVG, or WebP formats where their visual owner supports image export. The OpenDocument, EPUB, and email packages are thin adapters; they do not contain another layout or encoding engine.

## Product Problems Addressed

1. **Format labels could disagree with payload bytes.** Results now identify and validate their encoded content and dimensions at construction.
2. **Density was ambiguous.** `AtDpi(...)` / `TargetDpi` map document units to output pixels. PNG, JPEG, and TIFF carry encoded density, and results expose DPI plus physical dimensions.
3. **Save results did not prove where data landed.** Direct and fluent saves normalize extensions, use explicit conflict policy, and return `SavedPath`. Batch `SaveFiles()` returns payload-free file metadata for large jobs.
4. **Existing files could be replaced accidentally.** The default is `FailIfExists`; callers must select `Replace` or `CreateUnique`.
5. **Batch APIs encouraged full materialization.** Every main paged/document family has streaming consumer paths, async streaming where real async work exists, deterministic order, cancellation, and optional bounded concurrency.
6. **Per-image pixel limits did not bound a whole job.** Shared defaults now cap output count, aggregate raster pixels, and aggregate encoded bytes. Violations throw `OfficeImageExportBatchLimitException`.
7. **Diagnostics were strings without a consistent loss model.** Diagnostics classify approximation, omission, failure, or no loss. `OfficeImageExportPolicy` can reject all loss, omissions, failures, or selected stable codes before data is returned or saved.
8. **Font behavior varied by package.** Caller-supplied TrueType faces are resolved first across the shared pipeline. A missing requested family emits `IMAGE_FONT_SUBSTITUTED` instead of silently changing typography.
9. **Excel and Visio retained typography-specific branches.** Excel uses the shared font diagnostic and caller-font collection. Visio raster and SVG text now honor the requested family; SVG can embed supplied faces.
10. **Supported SVG pictures could disappear from raster output.** Drawing now parses and rasterizes the bounded SVG subset it owns. Unsupported SVG features continue through the caller codec or a visible fallback with `IMAGE_SOURCE_DECODE_FALLBACK`.
11. **Raster safety was fragmented and sometimes late.** `OfficeRasterExportPlanner` combines caller, renderer, and encoder limits before allocation and either reduces scale or throws a typed limit exception.
12. **HTML format requests once returned mislabeled PNG bytes.** HTML now uses the shared encoder for the selected format.
13. **Generic async builders could run synchronous render delegates.** Resource-aware HTML, email, and EPUB paths have real asynchronous resolution; in-memory projection remains synchronous.
14. **OpenDocument, EPUB, and email had no direct visual bridge.** ODT/ODS/ODP reuse their existing Office conversion owners, while EPUB and email reuse HTML. Conversion and fallback diagnostics remain attached to the image result.
15. **Batch filenames varied by operating system.** Shared batch export emits bounded, unique, Windows-portable names and protects reserved device names on every host.
16. **Visual tolerances could hide a severe small-area regression.** Shared comparisons now report and gate mean absolute, root-mean-square, and luminance error in addition to changed pixels. Binary PowerPoint fixtures compare against LibreOffice output.
17. **Complex-script behavior stopped at the PDF writer.** The shared text-shaping provider, language hint, and cooperative cancellation now reach direct raster export in Word, Excel, PowerPoint, HTML, OneNote, Visio, and loaded PDF pages, including nested Drawing effects/patterns, package-backed Visio SVG previews, and OneNote visual-PDF rasterization.
18. **Common TIFF source pictures required an application codec.** Drawing now reads classic 8-bit chunky grayscale, palette, RGB/RGBA, and DeviceCMYK strips with no compression, PackBits, or Deflate, including horizontal prediction. The same decoder is used by every document renderer.
19. **PDF annotations without appearance streams disappeared from page images.** Supported free-text, text-markup, shape, line, ink, path, stamp, and caret annotations now reuse the existing bounded appearance synthesizer. The result carries `render.annotation.appearance-synthesized` as an approximation; unsupported annotation kinds remain diagnosed omissions.

## Intentionally Retained Low-Level APIs

`PdfReadPage.ToDrawing()` is the public PDF-to-`OfficeDrawing` projection; the internal page renderer continues to serve inspection and visual comparison. `PdfPageRenderResult` also carries elapsed time, continue-on-error state, and typed PDF capability diagnostics needed by OCR, destructive-crop verification, and redaction verification. `PdfReadPage.ToImage()` and `PdfReadDocument.ToImages()` remain the general export surface.

Visio also retains its embedded-SVG interpreter. Stencil artwork needs linked images, clipping, nested opacity, and CSS that are outside Drawing's deliberately bounded general SVG reader. Visio still delegates output encoding, result validation, fonts, save policy, and batch mechanics to Drawing.

## Remaining Product Limits

- Word output is an OfficeIMO pagination estimate. Exact Microsoft Word pagination requires Word's layout engine.
- Arbitrary PDFs can contain operators, fonts, transparency, forms, annotations, and producer-specific constructs outside the first-party projection. Per-page diagnostics and the capability manifest remain the coverage contract.
- TIFF input support is bounded to classic 8-bit chunky grayscale, palette, RGB/RGBA, and DeviceCMYK strips with no compression, PackBits, or Deflate. LZW/JPEG compression, tiles, planar data, floating-point samples, BigTIFF pixels, and multi-page TIFF remain caller-codec cases. WebP input support is limited to the literal-lossless form OfficeIMO emits.
- GIF input uses the first frame. Animated output and multi-frame TIFF are outside the static-image contract.
- ICC conversion, color-management parity, EXIF preservation, and CMYK workflows are not implemented.
- SVG raster input accepts the bounded Drawing subset. Complex filters, scripts, external documents, and unsupported SVG features require a caller codec or produce a diagnosed visible fallback.
- EPUB fidelity depends on loading with raw chapter HTML and resource bytes retained. Encrypted or incomplete packages remain diagnostic-driven.
- ODT/ODS/ODP visuals inherit the fidelity of both the OpenDocument conversion and the target Word/Excel/PowerPoint renderer. The combined diagnostics make that explicit.
- Font parity requires callers to provide the intended TrueType faces when platform fonts are not deterministic.
- The built-in complex-text fallback covers core Arabic contextual forms and bounded bidirectional reordering. Full OpenType substitutions/positioning, the complete Unicode bidi algorithm, and scripts such as Indic or Southeast Asian families require a caller `IOfficeTextShapingProvider`.

## Release Gate

An image converter change is complete only when:

- declared format, dimensions, MIME type, extension, and encoded bytes agree;
- raster density metadata and result physical dimensions agree;
- all allocations are bounded before the pixel buffer is created;
- direct and batch paths apply the same diagnostic policy and aggregate budgets;
- save APIs use explicit conflict behavior and report normalized committed paths;
- callers can stream large batches, cancel work, and observe progress;
- unsupported source content remains visible when possible and always produces stable diagnostics;
- document-specific semantics stay in their owner and reusable behavior stays in Drawing;
- approved visual or decoded-pixel evidence protects the changed fidelity slice;
- no new runtime dependency is introduced.
