# OfficeIMO To-Image Conversion Assessment

Date: 2026-07-18

## Outcome

OfficeIMO has a credible first-party image-export foundation. The strongest architectural choice is already in place: document packages resolve their own semantics and route reusable drawing, text, raster, SVG, and encoding work through `OfficeIMO.Drawing`.

The audit found one correctness bug, two material coverage gaps, stale capability documentation, and several cross-package hardening opportunities. This branch fixes the correctness bug, brings HTML and Visio to the shared five-format contract, adds Visio document batch export, and makes batch filenames portable across operating systems.

## Current Conversion Matrix

| Source | Single image | Batch images | Output formats | Evidence and limits |
| --- | --- | --- | --- | --- |
| `OfficeDrawing` | Yes | Adapter-owned | PNG, JPEG, TIFF, SVG, WebP | Shared encoder, raster canvas, SVG exporter, result/options/builders |
| Excel range | Yes | No | PNG, JPEG, TIFF, SVG, WebP | Mature range coverage and visual baselines |
| Excel worksheet/workbook | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Used range, explicit range, print area, manual-page-break slices |
| PowerPoint slide/presentation | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Fixed-layout projection; representative fixture coverage should grow |
| Word document page/range | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Estimated OfficeIMO pagination, not Microsoft Word pagination |
| HTML continuous/paged render | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Resource-aware async APIs; this branch fixes non-PNG raster encoding |
| OneNote page/section/notebook | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | Ink, math, images, selection, and bounded raster output |
| Visio page/document | Yes | Yes | PNG, JPEG, TIFF, SVG, WebP | This branch adds shared formats, fluent selection, and document batch export |
| PDF page | No | No | None | Logical/content parsing exists; a first-party page painter is still missing |
| Markdown, RTF, AsciiDoc, LaTeX | No direct adapter | No direct adapter | None directly | Existing native/HTML/PDF projections may support thin future adapters |
| OpenDocument, EPUB, email | No direct adapter | No direct adapter | None directly | Requires a deliberate visual projection rather than a format-label wrapper |

## Correctness And Coverage Findings

### Fixed in this branch

1. **HTML returned PNG bytes for JPEG, TIFF, and WebP requests.** `ExportImage(format)` treated every non-SVG format as PNG, then labelled the result with the requested format. The renderer now creates one raster scene and passes it to `OfficeRasterImageEncoder`.
2. **HTML cloning discarded raster encoder settings.** JPEG quality/progressive settings and TIFF compression were lost during `HtmlRenderOptions.Clone()`. Clones now own an independent copy of `RasterEncoding`.
3. **Visio stopped at PNG and SVG despite a reusable raster scene.** The native Visio raster renderer now exposes its resolved `OfficeRasterImage` internally, so JPEG, TIFF, and WebP use the shared encoders without a second geometry implementation.
4. **Visio lacked document batch and shared fluent APIs.** Pages and documents now support `ExportImage`, `ExportImages`, `ToImage()`, `ToImages()`, page selection, batch folder output, and all five shared formats.
5. **Batch filenames varied by host operating system.** Names containing characters such as `:`, `?`, or `*` were accepted on Unix and rejected on Windows. Shared batch export now produces bounded, unique, Windows-portable names on every host and protects reserved device names such as `CON`, `COM1`, and `COM¹`.
6. **The image capability documents described an older implementation.** They still called Word first-page-only, treated JPEG source decoding as missing, and documented removed builder aliases.
7. **Extreme HTML scales and configured fluent batches had weak contracts.** Surface validation now rejects non-finite scaled dimensions before numeric conversion, and `ToImage(options)` / `ToImages(options)` preserve a cloned caller configuration for fluent export and folder saves.

### Remaining high-priority gaps

1. **PDF page-to-image is the largest missing first-party converter.** The implementation should paint the PDF content model directly and reuse Drawing primitives. Wrapping Poppler, PDFium, a browser screenshot, or Office automation would violate the repository's dependency rule.
2. **Raster allocation policy is inconsistent.** OneNote and the new Visio surface reduce oversized raster requests with diagnostics. HTML has dimension caps, while Excel, PowerPoint, and Word can still reach large allocations before an encoder rejects the output. A shared pixel-budget contract belongs in Drawing options and render planning.
3. **Source raster decoding does not match output encoding breadth.** Drawing decodes PNG, baseline/progressive JPEG, uncompressed BMP, and the first GIF frame. TIFF and WebP can be written but not painted as embedded source images without a caller codec.
4. **Async builder methods do not cancel document rendering.** The shared builders render synchronously and asynchronously commit the resulting bytes. HTML's direct async APIs can cancel resource resolution and scene projection, but the generic builder delegate has no async rendering contract.
5. **Visual evidence is uneven.** Excel and Visio have strong visual gates. PowerPoint has selected authored fixtures; Word, HTML, and OneNote need smaller representative galleries that cover real documents, not only synthetic feature tests.
6. **Diagnostics are not uniformly typed.** Excel exposes stable diagnostic constants, while several other converters still publish string literals. A shared package-specific diagnostic-code convention would make policy filtering safer.
7. **Image metadata and color management need a product decision.** JPEG exposes useful encoding controls, but DPI metadata, ICC profiles, EXIF preservation, multipage TIFF, and explicit alpha-flattening policy are not a consistent cross-format contract.

## Recommended Delivery Order

1. Add a repository-wide conformance test that verifies declared format, MIME type, magic bytes, dimensions, deterministic output, and path/stream behavior for every public converter.
2. Add a shared raster budget to `OfficeImageExportOptions`, propagate it through every clone, and emit a standard scale-limited diagnostic before allocation.
3. Add TIFF and WebP source decoders in `OfficeIMO.Drawing`, with malformed-input, pixel-budget, alpha, and round-trip tests.
4. Build a first-party PDF content painter, starting with paths, text, images, clipping, transforms, and transparency before exposing `PdfPage.ToImage()`.
5. Add thin Markdown and RTF image adapters only after choosing the fidelity-preserving owner: native projection when available, otherwise the shared HTML render model with conversion diagnostics retained.
6. Expand representative visual galleries for PowerPoint, Word, HTML, and OneNote and keep heavy comparison artifacts outside the repository.

## Acceptance Gates For A Premium Converter

Every converter should satisfy the same observable contracts:

- the declared format matches the encoded bytes;
- single, batch, path, stream, and cancellation behavior is documented and tested;
- result names and saved filenames are deterministic and portable;
- raster requests have an explicit allocation budget and never fail through uncontrolled memory growth;
- unsupported source features produce stable diagnostics or visible placeholders instead of disappearing;
- embedded source images cover the shared decoder set and use an explicit caller codec outside it;
- at least one representative real-world artifact is visually reviewed for each document family;
- adapters remain thin and reuse the owning document model plus `OfficeIMO.Drawing`.
