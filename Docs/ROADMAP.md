# OfficeIMO roadmap

This is the repository's single product backlog. It contains open work only. Implemented behavior is documented in package READMEs, support matrices, generated inventories, and current-state guides linked from [the documentation index](README.md).

An item belongs here when it has a clear product outcome and an owning package. Implementation checkpoints, completed task lists, release-wait notes, architectural rules, and competitor parity tables do not belong here.

Deliberately bounded compatibility contracts are not backlog by themselves. A preserved or rejected profile with a documented diagnostic becomes roadmap work only when the repository adopts a concrete supported shape and evidence plan for it.

## Release-wide quality

- [ ] Extend the generated Office compatibility catalog beyond the current Word, Excel, and PowerPoint legacy-format families into a package-neutral operation model for create, read, edit, preserve, inspect, convert, export, reject, and unsupported behavior.
- [ ] Generate compatible package README sections, website capability pages, MCP discovery, and support matrices from that model wherever one source can truthfully own the claim.
- [ ] Expand cross-producer fixture corpora with producer/version provenance and stable package or semantic diff policies.
- [ ] Add reproducible correctness, file-size, elapsed-time, peak-memory, allocation, cancellation, and deterministic-output evidence for representative workloads on every supported operating system.
- [ ] Add shared conversion reports and strict no-loss policies wherever an adapter can simplify, omit, rasterize, or preserve unsupported content.
- [ ] Promote conversion routes from Targeted to Established, Advanced, or ReferenceVerified only when the owning package adds the required reopen, realistic-fixture, visual or structural regression, and independent-producer evidence. Prioritize reverse PDF adapters, OpenDocument bridges, and OneNote, AsciiDoc, LaTeX, MHTML, and Visio routes that remain Targeted in the generated conversion catalog.
- [ ] Link every conversion support assessment to either an open roadmap outcome or an explicit intentional boundary, and expose that linkage in the generated route matrix. Fail catalog verification when any route gains a non-intentional limitation without an owner.

## Image engine

- [ ] Add bounded ICC-to-sRGB conversion only after representative matrix, LUT, RGB, CMYK, and malformed profiles have an explicit color-correctness corpus. Keep the current byte-preservation and metadata-loss report separate from color conversion, and reject profiles that exceed the declared parser, pixel, or allocation limits.
- [ ] Evaluate lossy or animated WebP, JPEG-in-TIFF, and BigTIFF as separate products only when real consumer files establish the required decode, encode, fallback, animation/page-loss, and memory contracts.
- [ ] Add bounded EMF/WMF rasterization when older Office-document fixtures demonstrate a material fidelity gain. Keep ICO, PCX, BMP/GIF encoding, and uncommon or expensive codecs behind the caller-codec boundary until a first-party consumer justifies ownership.

## PDF engine and conversion fidelity

- [ ] Complete managed rendering for difficult Type 3 glyph programs, higher-channel ICC profiles, certifiable ICC/output-profile shading composition, post-composition output-intent conversion for explicit PDF transparency, higher-dimensional alternate-color transforms beyond the bounded Type 0/2/3/4 function pipeline, advanced pattern resources, transparency groups, masks, and optional-content interactions. Require producer fixtures, per-page diagnostics, and raster plus semantic proof before removing a limitation.
- [ ] Apply resource-dictionary `/DefaultGray`, `/DefaultRGB`, and `/DefaultCMYK` substitutions consistently to content paint and image rendering, with diagnostics and producer fixtures for malformed or unsupported substitutes.
- [ ] Expand native Word, Excel, and PowerPoint reference corpora around DrawingML, SmartArt, chart variants, floating and grouped objects, headers and footers, print layout, and theme inheritance. Keep each improvement in the shared Drawing or PDF owner when more than one adapter needs it.
- [ ] Add exact-artifact external validation lanes for every supported PDF/A-2, PDF/A-3, PDF/A-4, PDF/UA-1, PDF/UA-2, Factur-X, and ZUGFeRD profile across representative generated artifacts and at least one independent producer corpus.
- [ ] Deepen PDF-to-PowerPoint and PDF-to-ODP semantic reconstruction beyond the current native text boxes, detected tables, safe basic shapes, and supported images. Add bounded support for clipping-aware placement, mixed-rotation text lines, multi-column reading order, richer vector paths, transparency and z-order, links and navigation, annotations, and embedded assets while preserving explicit reconstructed, visual-only, approximated, and omitted evidence.
- [ ] Deepen PDF-to-DOCX, PDF-to-ODT, and PDF-to-RTF reconstruction beyond the shared crop-, rotation-, spanning-band-, and column-aware reading-order foundation. Add richer page regions, images and supported drawings, links, annotations, headers and footers, and explicit page-break hints while preserving semantic editability and diagnostics rather than promising exact source pagination.
- [ ] Deepen PDF-to-XLSX and PDF-to-ODS table recovery beyond the shared continuation, repeated-header, typed-value, confidence, and OCR-aligned-column evidence now consumed by both routes. Add merged and spanning cells, richer row and column geometry, and border and fill evidence while keeping free-form page layout and invented formulas outside the spreadsheet contract.
- [ ] Deepen PDF-to-HTML positioned-review output for clipping, rotations, z-order, transparency, supported vector paths, optional-content state, annotations, form appearance, and deterministic resource placement. Keep semantic and positioned-review profiles separate, with visual-distance fixtures for the latter.
- [ ] Define format-appropriate editable, visual, hybrid, and narrow-data profiles across every reverse PDF adapter. Keep Word semantic, Excel table-shaped, PowerPoint/ODP slide-shaped, and HTML review-shaped, but use consistent selection, limits, page geometry, strict-loss behavior, warning codes, and report vocabulary so no route implies editability it did not produce.
- [ ] Deepen the 56-case DOCX, HTML, XLSX, PPTX, ODT, ODS, ODP, and PNG reverse-conversion gate with raster visual-distance baselines, stricter page-geometry thresholds, and broader native-editability and deterministic-loss assertions. Extend the current executable scanned, mixed-content, encrypted, and malformed cases with independent producer artifacts.
- [ ] Publish optional OCR provider packages and searchable-PDF text-layer output over the stable bounded core provider, confidence, provenance, and warning-merge contracts; keep OCR runtimes outside `OfficeIMO.Pdf` and verify each provider artifact independently.
- [ ] Define an explicit lossy scan-compression product with image-selection, downsampling, quality, color, metadata, accessibility, signature-invalidation, and measurable visual-difference policies. Keep it separate from the deterministic lossless optimizer.
- [ ] Evaluate an explicit XFA inspection or conversion product only with licensed specification coverage, hostile-input limits, external fixtures, and a fail-closed migration path to AcroForm or static visual output. Do not execute XFA in the core reader.

## HTML, RTF, and lightweight markup

- [ ] Deepen RTF semantic parsing and writing beyond the current Broad and Preserved contracts for complex and nested tables, fields and form-field data, embedded pictures and objects, advanced destination groups, Unicode and code-page interactions, lists and overrides, and producer-specific controls. Extend the independent Word, WordPad, and Outlook corpus with editable semantic round trips and deterministic preservation diagnostics rather than treating syntax preservation or adapter reopening as full semantic coverage.

## Security and protected content

- [ ] Add OCR-backed concealed-text assessment for raster images using bounded OCR regions plus pixel, geometry, and contrast evidence, with explicit safe-redaction policy. Never classify image metadata alone as visible or concealed text.
- [ ] Add native SVG text-visibility inspection and exact cleanup across presentation attributes, computed CSS, clipping, opacity, geometry, paint order, and background resolution, with browser-rendered adversarial fixtures.
- [ ] Add package-preserving concealed-HTML inspection and cleanup for MHTML and EPUB, including bounded MIME/resource handling, stylesheet resolution, signed-package mutation policy, and reopen validation.
- [ ] Add legacy encrypted-DOC import, encrypted-XLS authoring, and additional legacy ODF encryption profiles only with external producer corpora, explicit password/key and resource policies, dependency-free format ownership, and fail-safe preservation evidence.
- [ ] Materialize encrypted OneNote revisions only after the producer corpus, key-acquisition contract, dependency-chain semantics, and safe partial-result policy are defined. Until then, any encrypted current revision or dependency must fail closed without older-plaintext fallback.
- [ ] Extend ODF and EPUB signature interoperability beyond the bounded OfficeIMO XML package-manifest profile only with independent producer corpora, explicit trust policy, mutation/invalidation rules, and deterministic validation evidence.
- [ ] Link every non-intentional `NotSupported` protected-content catalog operation to an owning roadmap outcome, while keeping deliberate format and provider boundaries explicit rather than treating them as missing implementations.

## Document-format depth

- [ ] Improve DOCX-to-PDF rendering for floating and wrapped drawings, section-specific headers and footers, fields, footnotes and endnotes, SmartArt fallbacks, and pagination controls using Word-produced fixtures and page-level visual comparisons.
- [ ] Improve XLSX-to-PDF rendering for charts, pivot-table snapshots, conditional formatting, print titles and areas, manual page breaks, repeated rows and columns, external-link cached values, and advanced page setup using Excel-produced fixtures.
- [ ] Improve PPTX-to-PDF rendering for SmartArt fallbacks, master and theme inheritance, grouped and uncommon DrawingML, media poster frames, transitions' stable visual state, shadows, transparency, and advanced effects using PowerPoint-produced fixtures. Do not imply animation or media playback in static PDF output.
- [ ] Extend Excel/ODS conversion beyond the current typed formula, value, annotation, number-format, and validation subsets to date/time/custom-formula validations, conditional formatting, charts, pivot tables, and producer-backed advanced style fidelity.
- [ ] Add recursive typed ODT and ODP inline syntax with inherited-style resolution so nested spans and hyperlinks can convert without the current explicit flattening approximation.
- [ ] Expand Word/ODT and PowerPoint/ODP conversion for fields, notes, section/master-specific layout, advanced drawing geometry, media, and animation timing with producer fixtures and strict loss evidence.
- [ ] Improve ODT/ODS/ODP-to-PDF rendering for inherited styles, page and master layout, advanced drawings, charts, media poster frames, nested inline content, and office-suite-produced pagination evidence.
- [ ] Deepen OneNote section import, editing, and export for ink, embedded files, rich positioning, page metadata, internal links, attachments, and independently produced `.one` fixtures. Keep unsupported binary records preserved or diagnosed instead of silently flattening them.
- [ ] Deepen Visio-to-PDF rendering for masters and instances, grouped shapes, layers and visibility, themes, data graphics, connector routing, embedded objects, and page backgrounds using Visio-produced fixtures and page-level visual comparisons.

## Completion rule

Remove an item when its public API, compatibility boundary, tests, generated evidence, and user documentation agree. GitHub Releases records delivered history, while `MIGRATION.md` retains only upgrade actions; this file does not retain completed milestones.
