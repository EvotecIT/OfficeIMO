# OfficeIMO.Pdf Roadmap

Date: 2026-05-24

This roadmap tracks the path for `OfficeIMO.Pdf` to become a serious MIT-licensed, dependency-free PDF library that can eventually replace PSWritePDF/iText workflows and compete with QuestPDF-style report generation. The goal is not to clone either library feature-for-feature in one jump. The goal is to build a correct core, a pleasant document model, and a visual quality gate that lets us move forward safely.

## Direction

`OfficeIMO.Pdf` should become the PDF engine for the OfficeIMO family:

- MIT licensed.
- No runtime dependencies in the core package.
- Cross-platform and COM-free.
- Able to create, read, inspect, split, merge, and transform PDF files.
- Built around a logical document model rather than ad hoc string writing.
- Strong enough to become the export engine for `OfficeIMO.Word`, `OfficeIMO.Excel`, and `OfficeIMO.PowerPoint`.
- Friendly enough to expose later through PSWriteOffice PowerShell cmdlets.

QuestPDF can remain a temporary engine for `OfficeIMO.Word.Pdf`, but the strategic target is to replace that dependency with `OfficeIMO.Pdf` once the first-party layout engine is good enough.

The public API should stay Word-like and document-model oriented. Polished invoices, statements, reports, letters, and similar samples are quality fixtures and wrapper examples, not first-class concepts in the engine. If a fixture needs something, the reusable primitive should normally be a section, paragraph, table, style, drawing, image, header/footer, page setup, or layout-flow feature rather than an `Invoice`-style API.

## Architecture Target

PDF has no friendly AST in the same way Markdown, HTML, or Open XML do. Still, OfficeIMO can define layered models that serve the same purpose.

### 1. Logical Document Model

This is the user-facing model. It should describe intent, not PDF operators.

Examples:

- `PdfDocument`
- `PdfSection`
- `PdfPage`
- `PdfBlock`
- `PdfParagraph`
- `PdfTextRun`
- `PdfHeading`
- `PdfList`
- `PdfTable`
- `PdfTableRow`
- `PdfTableCell`
- `PdfImage`
- `PdfShape`
- `PdfCanvas`
- `PdfAnnotation`
- `PdfForm`
- `PdfStyle`
- `PdfTheme`. Initial theme bundles now apply reusable default text, heading, list, panel, horizontal rule, image, drawing, paragraph, and table styles at options, document, or page scope; `PdfTheme.WordLike()` gives callers a generic opt-in document rhythm instead of a template-specific API.

This is the model that report authors, Office exporters, and PowerShell wrappers should target.

### 2. Layout Model

This is the intermediate model produced after measuring and flowing content.

Examples:

- Page boxes.
- Block boxes.
- Inline boxes.
- Line boxes.
- Glyph runs.
- Table grid boxes.
- Image boxes.
- Vector boxes.
- Header/footer boxes.
- Floating or absolute-positioned boxes.

This is where pagination, wrapping, table sizing, column flow, page breaks, keep-together rules, and overflow diagnostics belong.

### 3. PDF Syntax Model

This is the low-level object model for reading and writing actual PDF files.

Examples:

- Indirect objects.
- Dictionaries.
- Arrays.
- Names. Generated and rewrite-style PDF name escaping now share one syntax escaper.
- Strings. Generated metadata, outlines, link URI strings, text operators, and rewrite-style metadata/string objects now share one literal-string escaper.
- Indirect references. Generated page, catalog, outline, annotation, image soft-mask, resource, trailer, and rewrite-style references now share one validated reference formatter.
- Streams.
- Cross-reference tables and streams. Generated PDFs and rewrite-style manipulation outputs now share one file assembler for the classic xref table, trailer, and startxref section.
- Trailers. Generated PDFs and manipulation rewrites share trailer assembly instead of carrying separate final-file writers.
- Catalogs. Generated catalog dictionaries and rewrite-style catalog prefix/name/reference entries now share one internal catalog dictionary builder, while manipulation-specific preservation of version, language, name trees, open actions, viewer preferences, XMP, output intents, attachments, associated files, and optional content remains in the rewrite layer.
- Info dictionaries. Generated PDFs, metadata editing, and merge outputs share one Info dictionary builder for title, author, subject, keywords, and producer metadata.
- Page trees. Generated page objects now reference a reserved `/Pages` object directly instead of using a temporary `/Parent 0 0 R` sentinel and patching serialized page objects later; generated PDFs, page extraction, and merge outputs share one `/Pages` dictionary builder.
- Page dictionaries. Generated page dictionaries now share one builder for parent references, media boxes, resource dictionaries, content references, and annotation arrays.
- Indirect objects. Generated indirect-object creation, explicit object reservation/replacement, and rewrite-style object wrapping now share one object-byte helper instead of rebuilding object envelopes at each call site.
- Resource dictionaries. Page resource reference dictionaries for Font, XObject, ExtGState, and Shading entries now use a shared formatter with PDF name escaping instead of separate inline string joins; generated ExtGState alpha and axial shading object bodies now use one visual resource dictionary builder with opacity and finite-coordinate validation instead of inline page-writer assembly.
- Content streams. Generated page content stream objects now use a focused stream-object helper instead of inline object/header/length/endstream assembly; image XObjects, PNG alpha soft-mask streams, and rewrite-style `PdfStream` bodies now use the same object-byte helper with explicit stream dictionaries.
- Font dictionaries. Generated standard Type1 font objects and stamp-injected standard Type1 font resources now share one internal builder for base-font validation, PDF name escaping, and WinAnsi encoding declaration.
- Image XObjects. Generated and stamp-injected JPEG/PNG image dictionaries now share one image XObject dictionary builder for `/Image` metadata, filters, PNG predictor decode parameters, soft-mask references, and stream-object dictionary output; generated image dictionaries are separated from stream-object wrapping so JPEG, PNG, soft masks, and future form/image objects can share the same stream writer path.
- Form XObjects.
- Annotation dictionaries. Generated URI link annotations now share one annotation dictionary builder for `/Annot` `/Link` objects, URI literal-string escaping, finite coordinate validation, and positive rectangle checks.
- Outline dictionaries. Generated outline root and item dictionaries now share one outline dictionary builder for title escaping, parent/previous/next/child links, descendant counts, and `/Dest` destination arrays.

This layer should be reusable for split, merge, repair, page import, metadata editing, and form operations.

### 4. Content Operator Model

This is a typed representation of page drawing commands.

Examples:

- Text begin/end operators. The first internal content-stream builder now covers focused simple text, rich paragraph, table-cell, and header/footer text emission.
- Text positioning operators. The first internal content-stream builder now covers focused simple text, rich paragraph, table-cell, and header/footer text positioning.
- Font selection operators. The first internal content-stream builder now covers focused simple text, rich paragraph, table-cell, and header/footer font selection.
- Text-show operators. The first internal content-stream builder now covers focused simple text, rich paragraph, table-cell, and header/footer hex text-show emission.
- Transformation and XObject placement operators. The shared internal content-stream builder now covers generated document image placement, text/image stamp and watermark streams, and shared-shape local transform matrices.
- Path operators. Initial shared path commands now render ordinary and transform-local lines, rounded rectangles, ellipses, polygons, and freeform paths through PDF move, line, cubic Bezier, and close operators.
- Stroke/fill operators. The first internal content-stream builder now covers reused fill, stroke, stroke width, stroke cap, stroke join, dash arrays, rectangle, fill-stroke, line, and path-paint primitives for ordinary and transformed vector shapes, table, panel, rule, text-decoration, and separator rendering.
- Graphics state operators. The first internal content-stream builder now covers reused save/restore graphics state emission plus ExtGState resource application for opacity and shadow rendering, including image clipping, vector clipping, gradient-fill, transformed-shape, shadow, and opacity wrappers.
- Shading operators. Gradient fill draw calls now use the shared internal content-stream builder for normal and transform-local vector shapes.
- Image placement operators. Generated document image placement now uses the shared internal content-stream builder.
- Clipping operators. The shared internal content-stream builder now covers rectangle, rounded rectangle, ellipse, polygon, and freeform path clipping for shape, drawing-scene, gradient-fill, transform-local, and image XObject placement paths.
- Save/restore graphics state.

This is needed for both writing new pages and understanding existing pages.

### 5. Compiler Pipeline

The long-term pipeline should be:

1. User API or Office exporter builds the logical document model.
2. Layout engine turns logical content into positioned layout boxes.
3. Renderer turns layout boxes into content operators and resources.
4. PDF writer serializes syntax objects into a valid PDF.
5. Visual regression checks rasterized output against approved baselines.

For reading and manipulation:

1. PDF parser builds the syntax model.
2. Resource resolver and content parser build page/operator models.
3. Text/image/vector extractors expose useful high-level information.
4. Manipulation APIs copy, rewrite, or transform object graphs safely.

## Current State

`OfficeIMO.Pdf` already has a useful seed:

- Basic PDF generation.
- Metadata.
- Standard PDF fonts, with shared option, composition, stamping, writer style-selection, metric, and base-font-name enum validation that rejects invalid values instead of falling back silently; writer font-family normalization preserves Helvetica, Times, and Courier families across regular, oblique, bold, and bold-oblique variants, Helvetica and Times family text measurement plus standard-font text span readback use built-in glyph-width tables, including common WinAnsi punctuation and accented Latin letters, instead of average character widths, and generated/stamped text now reports unsupported WinAnsi characters instead of replacing them with `?`.
- Headings and paragraphs. Shared simple text wrapping now preserves explicit hard line breaks for headings, table cells, list items, captions, and other non-rich text surfaces instead of collapsing them into ordinary spaces.
- Rich text runs. Initial rich paragraph wrapping, alignment, justification, underline/strike/link rectangles, superscript/subscript baseline shifts, scaled run measurement, Word-compatible default half-inch tab stops with paragraph-style overrides, explicit paragraph tab runs with dotted, hyphen, and underscore leaders plus left, center, right, and decimal-aligned values, and line breaking use proportional Helvetica/Times-family standard-font glyph widths instead of average character counts.
- Bullets and numbered lists.
- Horizontal rules with reusable `PdfHorizontalRuleStyle` thickness, color, spacing, keep-with-next behavior, and document/page/theme defaults.
- Panels.
- Tables.
- Links. Paragraph, heading, image, shape, drawing-scene, vector convenience, and table-cell URI annotations are now generic document primitives, including wrapped heading lines, aligned row/column headings, top-level/row-column images, fixed visual flow objects, top-level and compose/row-column vector helper calls, top-level tables, compose/row-column table flows, and linked column/row-spanned table cells whose annotations cover the merged text frame, with escaped `/Contents` metadata sourced from link text, heading text, image/shape/drawing/vector metadata, or cell text and inspector readback for generated heading, image, shape, drawing, and convenience-vector links.
- JPEG images and simple non-interlaced 8-bit grayscale/grayscale-alpha/RGB/RGBA PNG images, including PNG alpha soft masks, with reusable `PdfImageStyle` alignment, fit, clip, spacing, keep-with-next behavior, and document/page/theme defaults plus validation backed by `OfficeIMO.Drawing.OfficeImageReader`.
- PDF RGB color values validate components before writer/operator use.
- Rows/columns with configurable and reusable style-driven gutters, row-level spacing rhythm, keep-together and keep-with-next page flow, column-local item groups, bullet/numbered lists, panel paragraphs, and compact tables.
- Headers and footers/page numbers, with simple header text/page-token rendering, Word-like left/center/right text zones for running, first-page, and even-page variants, zone fit/overlap validation, visible `{page}` / `{pages}` tokens that continue across flows by default, configurable visible page-number starts, decimal/roman/alphabetic page-number styles, and section-local first/even/odd variant selection for Word-like section numbering, header font/alignment/placement validation, footer segment construction validation, assigned/readback footer segment list snapshots, direct segment-template rendering without requiring the page-number flag, and shared footer placement validation for page-number and segment-based footers.
- Invisible `Spacer(...)` flow blocks for document, page-content, column, nested element, and row/column composition, plus direct page-content, column, item, and nested element `PageBreak()` flow transitions, so business-shaped fixtures can add generic rhythm and pagination without inserting fake blank text or template-specific engine concepts into extracted content.
- Save to file/bytes, with shared sync/async path validation and async path cancellation before creating directories, rendering, or writing files.
- Basic PDF syntax parsing, including object-boundary scanning that ignores `stream` and `endobj` tokens inside literal strings so ordinary text values cannot truncate parsed objects.
- Core path read helpers reject null, empty, or whitespace paths before attempting file reads.
- Encrypted PDFs are detected and rejected with a clear unsupported diagnostic before parser-supported read/manipulation helpers attempt to process page content.
- Signed PDFs, form PDFs, complex outline/bookmark PDFs, complex page-label number-tree PDFs, unsupported catalog name-tree PDFs, unsupported named-destination name-tree PDFs, complex open-action dictionary PDFs, complex viewer-preference PDFs, complex XMP metadata PDFs, complex catalog URI PDFs, tagged PDFs, complex output-intent PDFs, complex embedded-file/associated-file PDFs, complex optional-content/layer PDFs, and active-content PDFs are detected and rejected with clear unsupported diagnostics before rewrite-style manipulation helpers copy, merge, edit, metadata-rewrite, stamp, or watermark page content; simple direct catalog `/PageMode`, `/PageLayout`, `/Version`, `/Lang`, simple direct `/PageLabels` number trees, simple outline trees including simple GoTo action outline entries whose destinations point only at copied pages, direct `/Dests` dictionaries, simple `/Names` `/Dests` name trees including leaf `/Kids`, destination-array and simple GoTo dictionary `/OpenAction` entries, simple `/ViewerPreferences` dictionaries, simple catalog `/Metadata` XMP XML streams, simple catalog `/URI` base dictionaries, simple `/OutputIntents` metadata graphs, simple `/Names` `/EmbeddedFiles` attachment trees, simple catalog `/AF` associated-file arrays, and simple `/OCProperties` optional-content metadata are preserved during rewrite-style manipulation, with copied-page page labels reindexed.
- Manipulation path input helpers reject null, empty, or whitespace input paths before attempting file reads.
- Text extraction.
- Text spans.
- Simple column-aware text extraction, including `PdfTextExtractor` facade overloads that accept `PdfTextLayoutOptions` for bytes, paths, streams, byte/path/stream whole-document UTF-8 output to paths or caller-owned streams, page-file output, stream-to-page-file output, and byte-array-to-page-file output. `PdfTextExtractor.ExtractTextByPageRanges(...)` accepts reusable `PdfPageRange` lists for byte/path/stream readback and deterministic selected source-page-numbered text file output with or without layout options, preserving caller order while deduplicating overlapping selections for wrapper-friendly range grammar. `PdfTextExtractor.ExtractStructuredByPage(...)` now returns per-page structured lines, lists, dot/hyphen/underscore leader rows that preserve decimal/currency value punctuation, and simple detected tables for bytes, paths, and streams; `ExtractStructuredByPageRanges(...)` applies the shared range grammar to structured readback. `PdfTextExtractor.ExtractTablesByPage(...)` returns page-numbered detected table groups directly for wrappers that only need table data, `ExtractTablesByPageRanges(...)` keeps source page numbers for selected range-list table readback, and the byte-array/path/stream output-directory overloads write deterministic escaped CSV files for all pages or selected source-page ranges. Column-aware readback splits wide same-baseline runs before gutter detection so generated row/column documents can be extracted in left-column then right-column order, while structured table readback keeps clear one-line table split candidates so generated simple tables can round-trip into detected table rows without exposing template-specific APIs.
- Simple structured extraction for lines, lists, leader rows, and tables.
- Text and image extraction directory-output helpers validate/create output directories before reading path inputs, reject file targets with clear argument errors, accept byte-array and stream inputs with caller-provided base names, and write deterministic page-numbered files for wrapper-friendly PSWritePDF parity; image extraction also accepts `PdfPageRange` range lists for byte/path/stream/document inputs and deterministic selected source-page image files, preserving caller order while deduplicating overlaps. Compatible grayscale/RGB Flate image XObjects with grayscale `/SMask` alpha extract as gray-alpha/RGBA PNG files so OfficeIMO-authored PNG alpha can round-trip through extraction.
- Header-version, encryption-marker, digital-signature-marker, form-field-marker, annotation-marker, outline/bookmark-marker, catalog-view-setting-marker, page-label-marker, catalog-name-tree-marker, named-destination-marker, open-action-marker, viewer-preference-marker, tagged-structure-marker, XMP-metadata-marker, catalog-URI-marker, output-intent-marker, embedded-file-marker, optional-content-marker, and active-content-marker probing through `PdfInspector.Probe` without full parsing.
- Read/rewrite preflight reports through `PdfInspector.Preflight`, including `CanRead`, `CanRewrite`, parsed document info when available, structured `ReadBlockers` / `RewriteBlockers`, `HasReadBlocker(...)` / `HasRewriteBlocker(...)` helpers, and diagnostics suitable for PSWriteOffice wrappers; encrypted PDFs, missing-header inputs, empty-page-tree PDFs, parser-unsupported PDFs, and PDFs whose page content streams use unsupported filters expose read blockers, form PDFs remain readable but are rewrite-blocked until form preservation/fill/flatten support exists, complex open-action dictionaries, complex non-GoTo outline actions, complex page labels, unsupported catalog name trees, malformed or unsupported named-destination name trees, complex viewer preferences, complex XMP metadata, complex catalog URI dictionaries, complex output intents, complex embedded/associated files, complex optional content, or active-content PDFs remain readable but are rewrite-blocked until catalog/document metadata preservation exists, tagged PDFs remain readable but are rewrite-blocked until accessibility structure preservation exists, and simple direct catalog view settings, simple outlines including simple GoTo action outline entries, simple direct page labels, direct named destinations, simple destination name trees including leaf `/Kids`, destination-array open actions, simple GoTo open-action dictionaries, simple viewer preferences, simple catalog XMP metadata streams, simple catalog URI base dictionaries, simple output intents, simple embedded-file attachment trees, simple associated-file arrays, plus simple optional-content metadata are detected without blocking rewrite.
- Page count, page size, orientation, rotation, catalog page mode/layout/version/language values, simple page-label rules, simple document open-action targets, simple viewer preference entries, simple AcroForm field names/types/simple values, simple page URI link annotation summary counts, distinct document-level link URI targets, document-level page-aware link lists, named destination names/targets, and per-page annotations with contents metadata, header version, digital signature presence, form-field presence, annotation presence, outline/bookmark presence, catalog-view-setting presence, page-label presence, catalog-name-tree presence, named-destination presence, open-action presence, viewer-preference presence, tagged-structure presence, XMP metadata presence, catalog URI presence, output-intent presence, embedded-file presence, optional-content presence, and active-content presence inspection through `PdfInspector`.
- Initial byte/path/stream page extraction, single-range and multi-range extraction with output stream writes for byte, stream, and path inputs, byte-returning path helpers, repeated selected-page/range cloning, single-page splitting, inclusive `PdfPageRange` chunk splitting through `PdfPageExtractor`, and enumerable file-list merge output to paths or streams through `PdfMerger`, with `PdfPageRange` overloads for wrapper-friendly range selection, preserving simple direct catalog `/PageMode`, `/PageLayout`, `/Version`, `/Lang`, simple direct `/PageLabels` number trees, simple outline trees including simple GoTo action outline entries whose destinations point only at copied pages, direct `/Dests` dictionaries, simple `/Names` `/Dests` name trees, destination-array and simple GoTo dictionary `/OpenAction` entries, simple `/ViewerPreferences` dictionaries, simple catalog `/Metadata` XMP XML streams, simple catalog `/URI` base dictionaries, simple `/OutputIntents` metadata graphs, simple `/Names` `/EmbeddedFiles` attachment trees, simple catalog `/AF` associated-file arrays, and simple `/OCProperties` optional-content metadata while reindexing copied-page labels, pruning stale destinations/open actions, and dropping stale outline trees/name-tree destinations whose target pages are not copied.
- Rewrite-style stream serialization now separates `PdfStream` dictionary emission from stream body wrapping so extraction, merge, edit, stamp, and watermark outputs share a cleaner object serialization path, copied-object references are normalized to generation 0 because rewrite outputs currently emit all cloned indirect objects as generation 0, rewrite graph collection/serialization rejects wrong-generation source references instead of silently remapping them to the active object, and `PdfInspector.Preflight` reports invalid rewrite object references before wrapper operations start.
- Page extraction path output helpers validate output paths before reading inputs, create parent directories, and reject empty paths or existing directory targets with clear argument errors.
- Split-to-directory helpers validate/create output directories before reading inputs, reject file targets with clear argument errors, and write deterministic page-numbered or page-range files for wrapper-friendly PSWritePDF parity.
- Initial byte/path/stream PDF merge with output stream writes and byte-returning file merge helpers through `PdfMerger`, preserving simple direct catalog `/PageMode`, `/PageLayout`, `/Version`, `/Lang`, simple direct `/PageLabels` number trees, simple outline trees including simple GoTo action outline entries, direct `/Dests` dictionaries, simple `/Names` `/Dests` name trees, destination-array and simple GoTo dictionary `/OpenAction` entries, simple `/ViewerPreferences` dictionaries, simple catalog `/Metadata` XMP XML streams, simple catalog `/URI` base dictionaries, simple `/OutputIntents` metadata graphs, simple `/Names` `/EmbeddedFiles` attachment trees, simple catalog `/AF` associated-file arrays, and simple `/OCProperties` optional-content metadata from the first source.
- Merge file output helpers validate output paths before reading inputs, create parent directories, and reject empty paths or existing directory targets with clear argument errors.
- Initial byte/path/stream selected-page import through `PdfPageImporter.AppendPages`, `PrependPages`, `InsertPages`, `InsertPageRange`, `AppendPageRanges`, `PrependPageRanges`, and `InsertPageRanges`, importing selected one-based source pages including repeated selections, inclusive source ranges from `firstPage` / `lastPage` pairs or `PdfPageRange`, parsed range lists with repeated/overlapping ranges as cloned pages in caller order, or all source pages when no selection is supplied, before, after, or inside a target PDF while reusing extraction plus merge object-copy behavior. Helpers can return bytes, write to paths, or write byte, stream, or path inputs to caller-owned output streams for wrapper pipelines. Insert operations keep the target document as the primary catalog/metadata source even when inserted pages become the first visible pages.
- Page import path helpers validate source-page selection before file reads, validate output paths before reading inputs, create parent directories, and reject empty paths or existing directory targets with clear argument errors.
- Initial byte/path/stream page duplication, movement, deletion, inclusive-range deletion, reordering, and rotation with output stream writes for byte, stream, and path inputs through `PdfPageEditor`; range edits accept `firstPage` / `lastPage` pairs, reusable `PdfPageRange` values, or parsed range lists for duplicate/move/delete/reorder/rotate flows so wrappers can use one generic page-range model, page duplication keeps original document order and inserts cloned copies after selected source pages, including repeated selections or repeated/overlapping parsed ranges as repeated clones, page movement moves selected pages or parsed range lists as a group in original relative order before another source page or to the end, and range-list move/rotate treats overlapping ranges as one selected page set.
- Page editing path output helpers validate output paths before reading inputs, create parent directories, and reject empty paths or existing directory targets with clear argument errors.
- Initial byte/path/stream metadata editing with byte-returning path helpers plus output stream writes for byte, stream, and path inputs through `PdfMetadataEditor`.
- Metadata editing path output helpers validate output paths before reading inputs, create parent directories, and reject empty paths or existing directory targets with clear argument errors.
- Initial byte/path/stream text/image stamping and watermarking with byte-returning path helpers plus output stream writes for byte, stream, and path PDF inputs through `PdfStamper`; text/image stamp streams now use the shared internal content-stream helper, default text watermark placement uses the same standard-font glyph-width measurement as generated layout instead of average character widths, and text/image stamp option models snapshot page-number arrays, provide `UsePageRange(...)` helpers for inclusive one-based page ranges from `firstPage` / `lastPage` pairs or reusable `PdfPageRange` values plus `UsePageRanges(...)` for parsed range lists without wrappers materializing page arrays, treat overlapping range-list selections as one page selection set, and reject invalid intrinsic coordinates, sizes, rotation, fonts, and duplicate/non-positive page selections before stamping.
- Stamper path output helpers validate output paths before reading PDF inputs or image payloads, create parent directories, and reject empty paths or existing directory targets with clear argument errors for wrapper-friendly PSWritePDF parity.
- Content-stream quality checks for stamper/watermark placement, rotation, dimensions, PNG alpha soft masks, and above/below-content layering, including custom image watermark sizing that still preserves watermark layering.
- Generated page resource dictionaries avoid unused header/footer-only font resources when headers, footers, or page numbers are disabled.
- Reusable theme bundles now apply Word-like default text, heading, list, panel, horizontal rule, image, drawing, paragraph, and table styles to following content through `PdfTheme`, `PdfOptions.ApplyTheme(...)`, `PdfDoc.Theme(...)`, and `PdfPageCompose.Theme(...)`; `PdfTheme.WordLike()` provides a built-in generic document theme with neutral typography, heading hierarchy, readable paragraph/list/table rhythm, and flow-object spacing. Rich paragraph layout now treats tab characters as default half-inch tab stops instead of collapsed single spaces, which gives future Word export a more faithful generic primitive without introducing template-specific APIs.
- Reusable default text style now applies Word-like default font, font size, and color to following page-flow content through `PdfTextStyle`, `PdfDoc.DefaultTextStyle(...)`, and `PdfPageCompose.DefaultTextStyle(...)`; no-options documents now start with Helvetica body/header/footer fonts so the plain engine default is proportional and document-like instead of monospace.
- Default heading styles now apply reusable Word-like H1/H2/H3 font size, line height, color, spacing before/after, and keep-with-next behavior to top-level and row/column headings when a heading does not provide an explicit style; heading spacing-before is preserved between visible blocks but suppressed at fresh page/column starts to avoid artificial top gaps; headings can keep with following visible paragraph/list/panel/table/rule/image/shape/drawing/row-section neighbors instead of only following paragraphs; callers can set defaults up front with `PdfOptions.DefaultHeadingStyles`, incrementally with `PdfOptions.SetDefaultHeadingStyle(...)`, fluently with `PdfDoc.DefaultHeadingStyle(...)`, page-by-page with `PdfPageCompose.DefaultHeadingStyle(...)`, or directly per heading through `H1/H2/H3(..., style: ...)`; compose item/element and row-column heading helpers now also expose explicit `align` and `color` overloads so local visual control stays generic and Word-like.
- Default list styles now apply reusable Word-like bullet and numbered list font size, line height, left indent, marker gap, color, spacing before/after, inter-item rhythm, keep-together, and keep-with-next page flow to top-level and row/column lists when a list does not provide an explicit style; callers can set them up front with `PdfOptions.DefaultListStyle`, fluently with `PdfDoc.DefaultListStyle(...)`, page-by-page with `PdfPageCompose.DefaultListStyle(...)`, or directly per list through `Bullets/Numbered(..., style: ...)`.
- Default panel styles now apply reusable Word-like boxed paragraph background, border, padding, max width, alignment, spacing, keep-together, and keep-with-next behavior to top-level and row/column panel paragraphs when a panel does not provide an explicit style; callers can set them up front with `PdfOptions.DefaultPanelStyle`, fluently with `PdfDoc.DefaultPanelStyle(...)`, page-by-page with `PdfPageCompose.DefaultPanelStyle(...)`, or directly per panel through `PanelParagraph(..., style: ...)`.
- Default horizontal rule styles now apply reusable Word-like separator thickness, color, spacing before/after, and keep-with-next behavior to top-level and row/column rules when a rule does not provide an explicit style; callers can set them up front with `PdfOptions.DefaultHorizontalRuleStyle`, fluently with `PdfDoc.DefaultHorizontalRuleStyle(...)`, page-by-page with `PdfPageCompose.DefaultHorizontalRuleStyle(...)`, or directly per rule through `HR(..., style: ...)`.
- Default image styles now apply reusable Word-like image alignment, fit, clipping, spacing before/after, and keep-with-next behavior to top-level and row/column images when an image does not provide an explicit style; callers can set them up front with `PdfOptions.DefaultImageStyle`, fluently with `PdfDoc.DefaultImageStyle(...)`, page-by-page with `PdfPageCompose.DefaultImageStyle(...)`, or directly per image through `Image(..., style: ...)`.
- Default drawing styles now apply reusable Word-like shape and drawing-scene alignment, spacing before/after, and keep-with-next behavior to top-level and row/column vector objects when an object does not provide an explicit style; callers can set them up front with `PdfOptions.DefaultDrawingStyle`, fluently with `PdfDoc.DefaultDrawingStyle(...)`, page-by-page with `PdfPageCompose.DefaultDrawingStyle(...)`, or directly per shape/drawing through `Shape(..., style: ...)` and `Drawing(..., style: ...)`.
- Default row styles now apply reusable Word-like column gutters, spacing before/after, keep-together, and keep-with-next page flow to row/column primitives when a row does not provide an explicit style; multi-column rows also have a built-in Word-like gutter when neither the row nor a default row style specifies one, while `Gap(0)` or `PdfRowStyle { Gap = 0 }` remains the explicit opt-out. Callers can set defaults up front with `PdfOptions.DefaultRowStyle`, fluently with `PdfDoc.DefaultRowStyle(...)`, page-by-page with `PdfPageCompose.DefaultRowStyle(...)`, through `PdfTheme.RowStyle`, or directly per row through `Style(...)`.
- Default paragraph styles now apply reusable Word-like paragraph geometry and page-flow settings to top-level and row/column paragraphs when a paragraph does not provide its own explicit style; paragraph spacing-before is preserved between visible blocks but suppressed at fresh page/column starts to avoid artificial top gaps; callers can set defaults up front with `PdfOptions.DefaultParagraphStyle` or fluently with `PdfDoc.DefaultParagraphStyle(...)`.
- Default table styles now apply reusable Word-like table appearance, rhythm, generic header/body/footer typography, cell line height, and keep-with-next behavior to top-level and row/column tables when a table does not provide its own explicit style; table spacing-before is preserved between visible blocks but suppressed at fresh page/column starts to avoid artificial top gaps; callers can set defaults up front with `PdfOptions.DefaultTableStyle`, fluently with `PdfDoc.DefaultTableStyle(...)`, page-by-page with `PdfPageCompose.DefaultTableStyle(...)`, or from the currently supported Word table style names, including `TableGridLight` plus Word's display-name alias `Grid Table Light`.
- Generic flow-object spacing-before is now treated as separation between visible neighbors and is suppressed at fresh page/column starts across lists, panels, horizontal rules, images, shapes, drawing scenes, rows, paragraphs, headings, and tables so style rhythm does not create artificial top gaps.
- Compose pages can set page-scoped default heading, list, panel, horizontal rule, image, drawing, paragraph, and table styles that snapshot input styles and do not leak to later pages.
- Paragraph visual-quality checks now cover justified wrapped lines expanding inter-word spacing while final lines and explicit line-break lines keep natural spacing and remain extractable.
- Paragraph page-flow checks now cover Word-like keep-together, keep-with-next, and widow/orphan styles in top-level and row/column flows, including keep-with-next across following visible paragraph/list/panel/table/rule/image/shape/drawing/row-section neighbors and clear diagnostics when a kept paragraph is taller than the page content frame.
- Paragraph indentation checks now cover Word-like first-line and hanging indents in top-level and row/column flows, with matching wrap measurement and rendered positions.
- Heading page-flow and style checks now cover Word-like orphan prevention, style snapshotting, theme propagation, page-scoped defaults, rendered font size/color, compose item/element and row-column alignment/color overrides, aligned heading-link rectangles, spacing-before/after rhythm with fresh page/column top suppression, and proportional standard-font wrapping for wide/narrow glyph runs in top-level and row/column flows so headings stay with following paragraphs and no longer rely only on hardcoded renderer constants.
- List style checks now cover snapshotting, theme propagation, page-scoped defaults, rendered font size/color, marker indentation, marker gap, spacing-after rhythm, fresh page/column spacing-before suppression, keep-together page flow, keep-with-next page flow, and proportional standard-font wrapping for wide/narrow glyph runs in top-level and row/column flows so bullets and numbering can improve without invoice-specific templates.
- Panel style checks now cover snapshotting, theme propagation, page-scoped defaults, rendered background color, max-width alignment, padding, spacing rhythm including fresh page/column spacing-before suppression, and keep-with-next page flow in top-level and row/column flows so callouts can improve as reusable boxed paragraphs instead of report-specific widgets.
- Row/column visual-quality checks now cover ordinary Word-like column primitives, asserting that extracted text lines remain inside their column frames, preserve explicit/default gutter clearance, keep readable baseline rhythm plus row-level breathing room, suppress flow-object spacing-before at column starts, and move kept rows together instead of splitting awkwardly at the bottom of a page.
- Table visual-quality checks cover proportional standard-font cell wrapping for wide and narrow glyph runs in top-level and row/column flows plus long unspaced token wrapping so prefixed identifiers and generated IDs stay inside the page content frame.
- Table visual-quality checks cover right alignment for common report numbers in top-level and row/column table flows, including currency symbols, percentages, and parenthesized negative/accounting values.
- Generic line-item visual gates now verify Word-like table primitives with weighted/min-width columns, wrapped product text, separated numeric columns, footer/summary row separation, margin containment, and follow-on rhythm without introducing invoice-specific engine APIs.
- Table visual-quality checks keep header rows distinct from body striping when explicit header fills are disabled.
- Shared `OfficeIMO.Drawing` shape descriptors now include two-stop linear gradient fill intent, and `OfficeIMO.Pdf` renders that intent as PDF axial shading resources in both normal and transformed vector-shape flows.
- Shared `OfficeIMO.Drawing` shape descriptors now include simple offset shadow intent, and `OfficeIMO.Pdf` renders that intent behind vector shape geometry using PDF graphics-state alpha.
- The professional report visual gate now includes a shared `OfficeIMO.Drawing` gradient ribbon with a simple shadow plus a translucent PNG status badge, with content-stream shading/graphics-state/soft-mask signals, so polished report fixtures exercise reusable drawing and image-alpha paths instead of only flat PDF-only shapes.
- Alignment guardrails reject unsupported heading, list, image, shape, drawing-scene, panel-box, table, caption, header, and footer alignment values before layout; table alignment is guarded at model construction across top-level, compose, and link-enabled APIs; mutable header, footer, panel-box, table-caption, and table column alignment properties reject unsupported values on assignment; table alignment list assignments snapshot caller collections; compose page blocks expose read-only content block collections; paragraph and panel paragraph blocks snapshot rich text runs into read-only model collections; list blocks snapshot caller items into read-only model collections; panel paragraph blocks snapshot caller styles at add time; paragraph and panel text alignment preserves supported justification and rejects invalid enum state.

`OfficeIMO.Word.Pdf` currently uses QuestPDF/SkiaSharp and can export selected Word content. This is useful as a bridge, but it is not the desired final engine.

`OfficeIMO.Drawing` is allowed and encouraged as the shared first-party engine for colors, image metadata, image fitting, font metadata, text measurement, reusable drawing primitives, and eventually office-wide drawing scene concepts. The PDF engine should reuse it where the concept is office-wide instead of growing PDF-only copies of the same primitives; PDF-specific serialization, page objects, and layout decisions should stay in `OfficeIMO.Pdf`. As the PDF visual layer grows, prefer lifting reusable drawing behavior into `OfficeIMO.Drawing` when Word, Excel, PowerPoint, Visio, or future OfficeIMO packages can consume it too.

The current visual output is not good enough to claim report-grade quality. In particular:

- Text spacing can look uneven.
- Table cell text that would escape its cell box now has PDF-level clipping to the cell rectangle with a small antialiasing tolerance, but richer Word-to-PDF table fidelity is still not good enough.
- Word-to-PDF tables are too primitive.
- Shape export is weak.
- Existing tests mostly prove that bytes/text exist, not that the PDF looks good.
- The `OfficeIMO.Pdf` README understates current features and does not explain the bigger ambition.

## Quality Bar

Every visible feature should have three gates before it is considered done:

1. PDF validity: generated files open in strict readers and can be parsed by OfficeIMO.
2. Text/structure validity: extractable text appears in the expected order where applicable.
3. Visual validity: rasterized page snapshots are compared against approved images.

Do not mark visual features complete using text extraction alone.

### Visual Test Harness

Add a repo-local visual test harness that is used for development and CI.

- Generate deterministic sample PDFs.
- Rasterize pages to PNG.
- Compare against approved baselines.
- Store small, stable baselines for core scenarios.
- Emit expected, actual, and diff PNG artifacts for raster failures.
- Keep rasterizer tooling in test/dev infrastructure, not as runtime dependencies of `OfficeIMO.Pdf`.
- Initial Poppler-backed PNG approval exists for the professional report, a two-page line-item statement fixture, a Word-like table style gallery with compact Accent1-6 swatches, a landscape showcase dashboard, plus compact hello-world, core-layout, style-cheatsheet, styled-runs, drawing-gallery, row-columns, links-rules, lists-tables, default-styles, three-page flow-dsl, and two-page headers-footers scenarios in `PdfDocRasterVisualBaselineTests`. These fixtures deliberately exercise generic layout primitives rather than introducing template-specific engine APIs. It is intentionally a test/dev lane: set `OFFICEIMO_REQUIRE_PDF_RASTERIZER=1` to make missing `pdftoppm` fail, `OFFICEIMO_UPDATE_PDF_RASTER_BASELINE=1` to approve refreshed PNGs, `OFFICEIMO_PDF_RASTER_PIXEL_TOLERANCE` for per-channel tolerance, and `OFFICEIMO_PDF_RASTER_ALLOWED_DIFF_PIXELS` for limited changed-pixel allowance.

Initial visual baseline scenarios:

- Hello world.
- Paragraph wrapping.
- Rich text runs.
- Headings.
- Links. Keep link support model-level and flow-agnostic: paragraphs, headings, images, and table cells should share annotation emission, contents metadata, and validation instead of template-specific helpers.
- Bullets and numbered lists.
- Simple table.
- Initial Word-like table style vocabulary and name resolver.
- Word-like table style gallery.
- Paginated list-table archetype.
- Table with long text.
- Table with column widths.
- Image placement.
- Header/footer/page numbers.
- Two-column layout.
- Page break.
- Panel/callout.
- Word-export paragraph sample.
- Word-export table sample.

## Roadmap

### 0. Stabilize The Mission

Make the intent visible and prevent accidental dependency creep.

- Update `OfficeIMO.Pdf` README to describe current features accurately.
- Document the split between `OfficeIMO.Pdf` and `OfficeIMO.Word.Pdf`.
- Add package guardrails that fail if `OfficeIMO.Pdf` gains runtime package dependencies.
- Add a public support matrix for create/read/manipulate/export capabilities.
- Add examples that produce professional-looking output, not toy PDFs.

Exit criteria:

- Users can understand what `OfficeIMO.Pdf` is today.
- Contributors can understand where it is going.
- The dependency-free promise is tested.

### 1. Visual Foundation

Fix the visible quality of the current builder before expanding the surface.

- Replace rough word spacing with predictable left/center/right alignment behavior. Initial simple text block and header/footer alignment now use standard-font glyph width estimates instead of character-count width guesses.
- Implement real justification only when spacing rules are intentionally designed.
- Improve proportional font metrics for standard 14 fonts.
- Add line height, paragraph spacing, padding, and margins as style concepts. Initial document and page default text styling exists through reusable `PdfTextStyle` objects and `DefaultTextStyle(...)` fluent configuration; initial `PdfTheme` bundles group default text, heading, list, panel, horizontal rule, image, drawing, row, paragraph, and table styles for options, document, or page scope, with `PdfTheme.WordLike()` as a reusable built-in theme for better generic document rhythm; initial invisible `Spacer(...)` flow gaps exist at document, compose item/element, and row/column scope for caller-managed rhythm without fake blank text; initial heading font size, line height, color, spacing before/after, and keep-with-next settings exist through `PdfHeadingStyle` / `PdfHeadingStyles`, options/document/page defaults, and per-heading overrides; initial bullet and numbered list font size, line height, left indent, marker gap, color, spacing before/after, inter-item rhythm, keep-together, and keep-with-next page flow exist through `PdfListStyle`, options/document/page defaults, and per-list overrides; initial panel background, border, padding, max width, alignment, spacing, keep-together, and keep-with-next defaults exist through `PanelStyle`, options/document/page defaults, and per-panel overrides; initial horizontal rule thickness, color, spacing before/after, and keep-with-next defaults exist through `PdfHorizontalRuleStyle`, options/document/page defaults, and per-rule overrides; initial image alignment, fit, clip, spacing before/after, and keep-with-next defaults exist through `PdfImageStyle`, options/document/page defaults, and per-image overrides; initial shape and drawing-scene alignment, spacing before/after, and keep-with-next defaults exist through `PdfDrawingStyle`, options/document/page defaults, and per-object overrides; initial table keep-with-next defaults exist through `PdfTableStyle`, options/document/page defaults, and per-table overrides; initial row gutter, spacing before/after, keep-together, and keep-with-next page-flow defaults exist through `PdfRowStyle`, options/document/page defaults, theme defaults, and per-row overrides, with a built-in multi-column gutter unless callers explicitly opt out; initial rich paragraph line height, left/right horizontal indents, first-line and hanging indents, spacing before/after, and Word-like keep-together / keep-with-next / widow-control page-flow options exist through `PdfParagraphStyle`; `PdfOptions.DefaultParagraphStyle` and `PdfDoc.DefaultParagraphStyle(...)` apply those settings to top-level and row/column paragraphs that do not provide an explicit style, while `PdfPageCompose.DefaultParagraphStyle(...)` applies them page-by-page in the compose DSL; scalar paragraph style setters reject invalid intrinsic values on assignment while combined text-frame width remains a layout-time diagnostic, paragraph style snapshots preserve caller intent after mutation, first-line and hanging indents affect both wrap measurement and rendered positions in top-level and row/column flows, kept paragraphs and lists move to the next page instead of splitting in top-level and row/column flows, keep-with-next paragraphs, headings, lists, panels, tables, rules, images, shapes, drawing scenes, and row sections avoid being stranded away from following visible flow neighbors including lists, panels, tables, rules, images, shapes, drawing scenes, and row sections, widow/orphan control avoids single-line paragraph starts at the bottom of a page where the page frame allows it, and kept paragraphs or lists that exceed the available page content height report a clear diagnostic.
- Panel scalar style setters now reject invalid border width, padding, max width, and outer spacing on assignment while layout-dependent padding conflicts remain render-time diagnostics; panel paragraphs participate in flow rhythm through `PanelStyle.SpacingBefore` and `PanelStyle.SpacingAfter`.
- Fix table overflow with clipping, wrapping, or diagnostics. Initial table cell and caption wrapping now uses proportional standard-font glyph metrics instead of average character widths, so wide glyphs wrap before crowding cell edges and narrow glyphs do not over-wrap.
- Add table column width rules: fixed, relative, auto, min/max. Initial fixed widths exist through `PdfTableStyle.ColumnWidthPoints`, min/max constraints through `PdfTableStyle.ColumnMinWidthPoints` and `ColumnMaxWidthPoints`, relative weights through `PdfTableStyle.ColumnWidthWeights`, and content-aware auto-fit sizing through `PdfTableStyle.AutoFitColumns` backed by `OfficeIMO.Drawing.OfficeTextMeasurer` with standard-font token minimums for wrapping stability; column sizing list setters now reject non-positive/non-finite intrinsic values and snapshot assigned collections; non-null fixed/min/max entries and relative weights outside the actual table grid fail during table layout/preflight; rendered visual gates cover fixed, relative, min/max, and content-aware column sizing in top-level and row/column table flows.
- Add table cell padding, borders, background, vertical alignment, text alignment, captions, typography, vertical spacing, and page-flow policies. Initial symmetric and side-specific cell padding exists through `PdfTableStyle.CellPaddingX`, `CellPaddingY`, `CellPaddingLeft`, `CellPaddingRight`, `CellPaddingTop`, and `CellPaddingBottom`; initial side-specific per-cell border overrides exist through `PdfTableStyle.CellBorders` / `PdfCellBorder`; initial body column background fills exist through `PdfTableStyle.BodyColumnFills`; initial absolute per-cell fills exist through `PdfTableStyle.CellFills`; explicit cell fills and borders now use combined row height for simple row-spanned cells and rectangular merged cells in top-level and row/column flows, explicit fill/border coordinates outside the table grid fail with clear diagnostics, and explicit fill/border coordinates targeting row-span or column-span continuation slots are skipped because those grid positions are occupied by the spanning cell; row striping and header/footer row fills skip continuation columns occupied by row-spanned cells from previous rows; body column fills skip continuation columns occupied by row-spanned or column-spanned cells; non-null body-column fills plus horizontal and vertical alignments outside the actual table grid fail during table layout/preflight; initial table captions exist through `PdfTableStyle.Caption`; initial table left indentation exists through `PdfTableStyle.LeftIndent`; initial table width caps exist through `PdfTableStyle.MaxWidth`; both honor left/center/right placement in top-level and row/column flows; initial `PdfTableCell` column spans render across combined column widths in top-level and row/column flows; initial `PdfTableCell.RowSpan` support lets simple vertically merged and rectangular merged cells occupy following-row grid columns, use combined row height, honor horizontal/vertical alignment inside the combined box, reject row spans that extend beyond the available table rows, reject row spans that cross resolved header/body/footer boundaries, and keep row/header/footer separators plus default table border grids from crossing the merged cell interior in top-level and row/column flows, including rectangular merged-cell vertical-grid gaps on row-span continuation rows; `PdfTableCell` can own URI link metadata so linked column/row-spanned and rectangular merged cells emit annotations over the merged text frame; initial table spacing exists through `PdfTableStyle.SpacingBefore` and `PdfTableStyle.SpacingAfter`; initial table keep-together and keep-with-next page-flow exists through `PdfTableStyle.KeepTogether` and `PdfTableStyle.KeepWithNext` for top-level and row/column tables; table keep-with-next first-row placement estimates use the same header/footer row-count, explicit cell-style coordinate bounds, column-scoped style bounds, configured column widths, and row-span boundary validation as rendering; initial oversized-row page splitting policy exists through `PdfTableStyle.AllowRowBreakAcrossPages`; initial table typography controls exist through generic header/body/footer font-size settings, cell line height, plus header/footer bold toggles; initial row/header/footer separator controls exist through `RowSeparatorColor` / `RowSeparatorWidth`, `HeaderSeparatorColor` / `HeaderSeparatorWidth`, and `FooterSeparatorColor` / `FooterSeparatorWidth`; Word-like table presets and `PdfTheme.WordLike()` include footer separator defaults for summary/footer rows, supported Word table style names include `TableNormal` plus Accent1-6 variants with Word default theme border, separator, and soft band colors for the existing light grid/list presets, and `TableStyles.CanonicalWordStyleNames` plus canonical-name helpers separate clean display/storage names from accepted alias spellings; table cell text that would escape its cell box is clipped to the cell rectangle at the PDF content-stream level with a small antialiasing tolerance; oversized caption-plus-first-row combinations fail with a clear layout diagnostic in top-level and row/column table flows; document default table styles exist through `PdfOptions.DefaultTableStyle` and `PdfDoc.DefaultTableStyle(...)`, page-scoped defaults through `PdfPageCompose.DefaultTableStyle(...)`, including supported Word table style names; table model/style validation now rejects invalid border width, padding, max width, left indent, spacing, captions, unsupported table/caption justification, alignment enum values, negative or oversized header/footer row counts, row height, row baseline offsets, font sizes, line height, negative or out-of-grid cell-fill/border coordinates, out-of-grid column-scoped style entries, invalid column spans, invalid, overlong, or boundary-crossing row spans, invalid cell link URIs, impossible column sizing, kept tables taller than the available page content height, and oversized rows when row splitting is disabled; table blocks snapshot rows, explicit cells including span metadata, styles, and link dictionaries into read-only model state; scalar setters reject invalid intrinsic values on assignment; and alignment, column-sizing, fill, border, and default-style collection setters snapshot assigned collections.
- Add deterministic visual baselines for all current features. Initial text geometry snapshots now cover representative and professional report fixtures, mixed Word-like flow rhythm across headings, paragraphs, panels, lists, tables, images, shapes, and row columns with no-cramped-baseline, same-baseline text-collision, and ambiguous-run-gap guards, row/column text-frame rhythm and gutter clearance, generic line-item table rhythm, the two-page line-item statement fixture, plus content-stream signals for images, PNG alpha soft masks, clipping, and vector drawing; initial Poppler raster PNG approval now covers the professional report reference, a two-page line-item statement fixture, a Word-like table style gallery with compact Accent1-6 swatches, a landscape showcase dashboard, plus compact hello-world, core-layout, style-cheatsheet, styled-runs, drawing-gallery, row-columns, links-rules, lists-tables, default-styles, three-page flow-dsl, and two-page headers-footers fixtures. Business-shaped fixtures should remain proof documents for the generic Word-like primitives rather than new invoice-specific surface area.

Exit criteria:

- Current examples look respectable.
- Long text does not collide in tables.
- Many-row tables and oversized wrapped rows continue across pages without drawing below the bottom margin.
- A report made from headings, paragraphs, lists, tables, and links looks polished.

### 2. Core PDF Syntax Engine

Turn the reader/writer internals into a dependable PDF object engine.

- Parse classic xref tables.
- Parse xref streams. Initial parser support now reads active xref stream entries, follows active xref-stream `/Prev` chains, follows stream-to-classic `/Prev` chains for object entries and inherited trailer `/Root` metadata, applies classic-trailer `/XRefStm` hybrid-reference supplements including trailer-like catalog metadata, treats active-chain `/Type /XRef` stream dictionaries as trailer-like metadata when no classic trailer is active, requires direct classic and xref-stream entries to match both object number and generation before replacing an object, resolves indirect object references through generation-aware lookup across reader and rewrite helpers, and ignores stale xref-stream dictionaries plus stale object streams outside the active classic or stream `startxref` chain for active revision selection; broader repair diagnostics remain roadmap work.
- Parse trailers and incremental updates.
- Preserve object identity.
- Decode common stream filters. Initial stream decoding covers uncompressed, Flate, ASCIIHex, ASCII85, RunLength, and LZW content streams, including LZW `/EarlyChange` and predictor `DecodeParms` for parser-supported text extraction paths.
- Support object streams. Initial `/ObjStm` expansion populates compressed objects and uses source-order precedence so later object streams can replace stale earlier compressed or explicit objects while later explicit objects still win.
- Support encrypted-file detection with clear unsupported diagnostics. Initial trailer/xref-stream detection rejects encrypted files before parser-supported read/manipulation helpers process page content.
- Add safe object copying between documents.
- Add resource renaming to avoid collisions when importing pages.
- Add repair diagnostics for malformed PDFs where possible.

Exit criteria:

- OfficeIMO can open, inspect, and rewrite common generated PDFs without losing basic structure.
- Page copy/import has the object graph support needed for split and merge.

### 3. Text And Font Engine

Text quality is the heart of PDF generation.

- Build a first-party font abstraction.
- Support standard PDF fonts well. Initial generated and stamp-injected standard Type1 font dictionaries now share one internal builder.
- Support embedded TrueType/OpenType fonts without third-party runtime dependencies.
- Add Unicode text support with ToUnicode maps.
- Add glyph measurement and shaping boundaries.
- Add fallback font selection.
- Add text extraction support for simple encoded fonts and ToUnicode maps. Initial readback already includes layout-option overloads on the `PdfTextExtractor` facade, structured-by-page and table-by-page facade readback, common text-showing operators, `TD` text-positioning line advances, a generated two-column regression gate for column-aware reading order, and a generated simple-table regression gate for structured table rows.
- Add diagnostics when text cannot be represented in the selected font. Initial generated/stamped standard-font text now rejects unsupported WinAnsi characters with the exact Unicode code point and explains that embedded Unicode fonts are required; raw control characters are rejected with layout-oriented diagnostics instead of being written as invisible PDF text bytes.

Exit criteria:

- English and Polish business reports render correctly.
- Text extraction returns readable text for generated PDFs.
- Font behavior is deterministic across Windows/Linux/macOS when fonts are supplied.

### 4. Layout Engine

Build the engine that can eventually replace QuestPDF for OfficeIMO scenarios.

- Blocks: paragraph, heading, bulleted list, numbered list, table, image, canvas, panel.
- Inline content: runs, links, line breaks, spans, superscript/subscript. Initial explicit rich paragraph line breaks exist through `PdfParagraphBuilder.LineBreak()` and newline normalization in text runs; tabs inside rich paragraph runs are treated as Word-like word spacing so raw tab control bytes do not reach PDF text-show operators; superscript/subscript run placement exists through `TextRun.Baseline`, `TextRun.Superscript(...)`, `TextRun.Subscript(...)`, and matching `PdfParagraphBuilder` helpers, with scaled measurement and PDF text-rise output; paragraph and panel text alignment reject invalid enum state while preserving justification, and heading, bullet list, numbered list, image, shape, and drawing-scene blocks reject unsupported alignment values before layout.
- Page flow: automatic pagination, page breaks, keep-with-next, keep-together, widow/orphan control. Initial long rich paragraph continuation across pages exists for top-level flow paragraphs, paragraph keep-together, keep-with-next, and widow/orphan control exist for top-level and row/column flows, headings avoid being orphaned from following visible flow neighbors in top-level and row/column flows, bullet/numbered lists, panel paragraphs, and horizontal rules can keep with the following visible block when they fit the page frame, lists and panels can keep together, and oversized bullet/numbered list items now continue across pages without crossing the bottom margin.
- Sections: page size, orientation, margins. Initial options-level, document-default, page-scoped, and section-scoped flow size, portrait/landscape orientation helpers, scalar margins, and reusable Word-compatible margin presets exist through `PdfOptions.PageSize`, `PdfOptions.Margins`, `PdfDoc.Size(...)`, `PdfDoc.Orientation(...)`, `PdfDoc.Portrait()`, `PdfDoc.Landscape()`, `PdfDoc.Margin(...)`, `PdfDoc.Margin(PageMargins)`, top-level `PdfDoc.Page(...)` / `PdfDoc.Section(...)`, `PdfDoc.Compose(...Page...)` / `Compose(...Section...)`, matching `PdfPageCompose` methods, and `PageMargins`; richer section inheritance and mid-page section breaks remain roadmap work.
- Headers and footers. Initial simple page header/footer text exists through `PdfOptions`, document-level `PdfDoc.Header(...)` / `PdfDoc.Footer(...)`, and page-scoped `PdfPageCompose.Header(...)` / `PdfPageCompose.Footer(...)`, with literal text formats, `{page}` / `{pages}` tokens, composed text/token segment builders for headers and footers, alignment, font, size, text color, margin-relative offsets with placement validation, document first-page header/footer overrides, and odd/even page overrides for Word-like cover-page/report flows.
- Multi-column flow.
- Absolute positioning escape hatch.
- Overflow diagnostics. Initial page setup diagnostics reject invalid intrinsic page sizes and margins at fluent assignment time, while render-time page option diagnostics reject invalid default/header/footer font selections and font sizes, header/footer alignment/placement, and impossible content frames; `PdfDoc.Create(options)` snapshots caller-provided options before rendering. Initial fixed-size flow block diagnostics exist for images, horizontal rules, vector shapes, and drawing scenes that exceed the available page content width or height; image blocks snapshot caller-provided bytes and validate intrinsic model state, while image, drawing, and horizontal rule styles validate intrinsic spacing at block construction; row/column composition rejects empty rows plus invalid gutters, non-finite, non-positive, over-100%, and over-allocated column widths before rendering, rejects gutters that exceed the available content width during render, and exposes read-only row/column model collections after composition; kept-together paragraphs and panels also report when their measured height exceeds the available page content height, and panel styles validate border width, padding, outer spacing, max width, panel-box alignment, and text-frame viability.
- Layout debug overlays.

Exit criteria:

- A multi-page report can be generated without manual page positioning.
- Layout failures explain what overflowed and where.

### 5. Tables

Tables need to be a flagship feature because PowerShell reporting depends on them.

- Header rows. Initial configurable leading header row count exists through `PdfTableStyle.HeaderRowCount`.
- Repeating headers across pages. Initial support exists for generated simple tables, including configured multi-row headers when they fit on the page.
- Footer rows. Initial trailing footer row count exists through `PdfTableStyle.FooterRowCount`, with footer fill/text styling and optional footer separators above the first footer row.
- Row striping. Initial body row striping is computed relative to the first body row and does not let configured header rows shift the stripe pattern.
- Table captions. Initial caption text, alignment, color, font size, and spacing exist through `PdfTableStyle.Caption`, `CaptionAlign`, `CaptionColor`, `CaptionFontSize`, and `CaptionSpacingAfter`, with rendered visual gates for top-level and row/column table flows.
- Cell padding. Initial symmetric and side-specific cell padding exists through `PdfTableStyle.CellPaddingX`, `CellPaddingY`, `CellPaddingLeft`, `CellPaddingRight`, `CellPaddingTop`, and `CellPaddingBottom`.
- Table spacing before and after. Initial spacing exists through `PdfTableStyle.SpacingBefore` and `PdfTableStyle.SpacingAfter`.
- Cell border model. Initial side-specific per-cell border overrides exist through `PdfTableStyle.CellBorders` / `PdfCellBorder`; assigned border dictionaries and border values are snapshotted, negative border coordinates fail on assignment, out-of-grid border coordinates fail during table layout/preflight, `PdfCellBorder.Width` rejects invalid intrinsic widths on assignment, row-spanned explicit cell borders use combined row height, explicit border coordinates skip row-span and column-span continuation slots, default table border grids skip row-spanned and rectangular merged-cell interiors in top-level and row/column table flows, and rendered content-stream gates cover top-level and row/column table flows. Richer border conflict resolution and collapse/merge behavior remains roadmap work.
- Cell background. Initial body column fills exist through `PdfTableStyle.BodyColumnFills`; initial absolute per-cell fills exist through `PdfTableStyle.CellFills`; assigned fill collections are snapshotted, negative cell-fill coordinates fail on assignment, out-of-grid fill coordinates fail during table layout/preflight, row-spanned explicit cell fills use combined row height, explicit fill coordinates skip row-span and column-span continuation slots, body column fills skip continuation columns occupied by row-spanned cells from previous rows or column-spanned cells in the current row, and rendered content-stream gates cover top-level and row/column table flows.
- Row striping. Initial body row striping exists through `PdfTableStyle.RowStripeFill`; stripe, header, and footer row fills skip continuation columns occupied by row-spanned cells from previous rows; rendered content-stream gates verify stripes are calculated relative to the first body row, do not apply to configured header rows, and stay out of row-spanned cell interiors in top-level and row/column table flows.
- Row separators. Initial body row separators exist through `PdfTableStyle.RowSeparatorColor` / `RowSeparatorWidth`; initial header separators exist through `PdfTableStyle.HeaderSeparatorColor` / `HeaderSeparatorWidth`; initial footer separators exist through `PdfTableStyle.FooterSeparatorColor` / `FooterSeparatorWidth`; row/header/footer separators skip row-spanned cell interiors in top-level and row/column table flows; Word-like table presets and `PdfTheme.WordLike()` provide neutral footer separator defaults; rendered content-stream gates cover top-level and row/column table flows.
- Column width strategies. Initial table left indentation exists through `PdfTableStyle.LeftIndent`; initial table max-width caps exist through `PdfTableStyle.MaxWidth`; fixed column widths exist through `PdfTableStyle.ColumnWidthPoints`, min/max constraints through `PdfTableStyle.ColumnMinWidthPoints` and `ColumnMaxWidthPoints`, relative column width weights through `PdfTableStyle.ColumnWidthWeights`, and content-aware auto-fit sizing through `PdfTableStyle.AutoFitColumns`; rendered visual gates cover top-level and row/column table flows.
- Row height strategies. Initial minimum row height exists through `PdfTableStyle.MinRowHeight`.
- Text wrapping in cells.
- Cell alignment. Initial horizontal alignment exists through `PdfTableStyle.Alignments`; initial vertical alignment exists through `PdfTableStyle.VerticalAlignments`; both are honored in top-level and row/column table flows, reject out-of-grid column entries during table layout/preflight, and include rectangular merged cells that align text inside the combined box.
- Row/page break policies. Initial row-by-row pagination, configurable oversized-row line splitting, and keep-together page-flow exist for generated simple tables, including row/column flows when the kept table or split row segment fits the page frame.
- Colspan and rowspan. Initial `PdfTableCell.ColumnSpan` support exists for simple column spans in top-level and row/column table flows, including linked spanned cells; initial `PdfTableCell.RowSpan` support exists for simple vertically merged cells and rectangular merged cells in top-level and row/column table flows, including linked row-spanned cells; explicit fills and borders on row-spanned and rectangular merged cells now paint over the combined box, explicit fill/border coordinates skip row-span and column-span continuation slots, linked merged-cell annotations cover the combined text frame, text alignment resolves against the combined box, overlong row spans fail with a clear model-level diagnostic, row spans that cross resolved header/body/footer boundaries fail with clear diagnostics, and row/header/footer separators plus default table border grids skip the merged cell interior, including column-span interior boundaries on row-span continuation rows; richer merged-cell conflict behavior remains roadmap work.
- Nested tables only after the simpler model is stable.

Exit criteria:

- DomainDetective/TestimoX-style report tables look clean.
- Long values and many rows paginate predictably.

### 6. Drawing, Images, And Charts

Build enough graphics support for reports before aiming for full creative layout.

- Reuse and expand `OfficeIMO.Drawing` for shared color, font, image, text measurement, and drawing primitives where possible.
- Treat `OfficeIMO.Drawing` as the home for reusable drawing engine concepts that can serve PDF, Word, Excel, PowerPoint, Visio, and future OfficeIMO packages. `OfficeIMO.Pdf` should consume those shared descriptors and keep only PDF-specific layout/serialization locally. Initial shared vector shape descriptors exist through `OfficeShape` / `OfficeShapeKind`, and simple grouped drawing scenes exist through `OfficeDrawing`.
- Initial color interop exists through `PdfColor` conversions to and from `OfficeIMO.Drawing.OfficeColor`.
- JPEG and simple PNG validation now use `OfficeIMO.Drawing.OfficeImageReader`; unsupported recognized formats fail explicitly until rendering support exists.
- Lines, rectangles, rounded rectangles, ellipses, polygons, paths. Initial validated horizontal rules plus flow lines, rectangles, rounded rectangles, ellipses, polygons, paths, and grouped scenes exist; flow vector shapes are backed by shared `OfficeIMO.Drawing.OfficeShape` and `OfficeDrawing` descriptors, use reusable `PdfDrawingStyle` for PDF flow placement/rhythm, and reject unsupported alignment values before layout.
- Fill/stroke colors. Initial two-stop linear gradient fill intent exists through `OfficeLinearGradient`; PDF maps vector shape gradients to native axial shading resources clipped to the shape geometry.
- Simple shape effects. Initial offset shadow intent exists through `OfficeShadow`; PDF maps vector shape shadows to alpha-backed offset geometry. Soft blur, glow, and richer effects remain future work.
- Fill and stroke opacity. Initial shared opacity descriptors exist through `OfficeShape.FillOpacity` and `OfficeShape.StrokeOpacity`; PDF maps them to `/ExtGState` alpha resources for vector shapes.
- Stroke width, dash styles, line caps, and line joins. Initial shared stroke descriptors exist through `OfficeStrokeDashStyle`, `OfficeStrokeLineCap`, and `OfficeStrokeLineJoin`; PDF flow lines, rectangles, rounded rectangles, ellipses, polygons, and paths render solid, dashed, dotted, and dash-dot strokes with configurable cap/join operators through the shared content-stream builder.
- Basic transformations. Initial shared transform descriptors exist through `OfficeTransform`; PDF maps them to graphics state matrices for flow and grouped vector shapes.
- Clipping paths. Initial shared clip path descriptors exist through `OfficeClipPath`; PDF maps rectangle, rounded rectangle, and freeform path clips to shared content-stream clipping operators for flow and grouped vector shapes.
- JPEG support.
- Broader PNG support, including palettes, interlace, richer color handling, and broader alpha/transparency cases beyond initial grayscale-alpha/RGBA soft masks.
- Image scaling and aspect ratio modes. Initial flow image stretch/contain/cover placement exists through shared `OfficeImageFit`; cover mode clips overflow to the target box, and unsupported alignment values are rejected before layout.
- Image clipping. Initial flow image clipping exists through shared `OfficeClipPath`; PDF maps rectangle, rounded rectangle, and freeform path clips to shared content-stream clipping operators around image XObjects.
- SVG import as a later optional parser if it can be done without runtime dependencies.
- Chart images from OfficeIMO chart engines as an export bridge.

Exit criteria:

- Logos, status badges, simple diagrams, and chart snapshots can be placed cleanly.
- Drawing primitives remain useful to Word, Excel, PowerPoint, Visio, and other OfficeIMO packages rather than becoming PDF-only types.

### 7. Annotations, Navigation, And Metadata

Make generated PDFs feel complete.

- Links and destinations. Initial URI link annotations exist for paragraphs, headings, images, shapes, drawing scenes, vector convenience helpers, and table cells, including escaped annotation contents metadata and linked merged-cell rectangles for column/row-spanned table cells; generic `Bookmark(...)` flow anchors emit simple `/Names` `/Dests` named destinations from top-level and row/column flows; paragraph `LinkToBookmark(...)` runs emit internal GoTo annotations targeting those named destinations with missing-target validation; `PdfReadPage` and `PdfInspector` can read simple page URI and named-destination link annotations with contents metadata; `PdfDocumentInfo` reports document-level readable link counts, distinct URI targets, distinct internal destination targets, plus a page-aware flattened link list; simple direct/names-tree named destinations are readable as named document targets; and simple destination-array plus GoTo dictionary open actions are readable as document navigation targets. Next slices should broaden destination/action support, richer keyboard/readback metadata, and preservation rules without becoming template-specific.
- Bookmarks/outlines. Initial `PdfOptions.CreateOutlineFromHeadings` support writes nested PDF outlines from H1/H2/H3 blocks through a shared outline dictionary builder, and `PdfInspector` can read simple outline trees, indirect destinations, direct/name-tree named-destination targets, and simple GoTo action destinations from the trailer-root catalog. Rewrite-style manipulation preserves simple outline trees, including simple GoTo action outline entries, whose destinations point only at copied pages, drops outline trees when a selected-page operation would leave stale outline destinations, and still blocks complex non-GoTo or additional-action outline trees.
- Catalog identity. Simple catalog `/Version` and `/Lang` values are readable through `PdfReadDocument` / `PdfInspector` and preserved during rewrite-style manipulation so split/merge/edit helpers do not strip document-level version or language metadata.
- Viewer preferences. Simple catalog `/ViewerPreferences` dictionaries are readable through `PdfReadDocument` / `PdfInspector` as generic key/value entries with boolean helpers, and preserved during rewrite-style manipulation; complex viewer preference graphs remain blocked until richer typed models exist.
- Page labels. Simple direct catalog `/PageLabels` number trees are readable through `PdfReadDocument` / `PdfInspector` as start-page/style/prefix/start-number rules and preserved/reindexed during rewrite-style manipulation using the trailer-root page tree; complex page-label trees remain blocked until richer number-tree support exists.
- Catalog URI. Simple catalog `/URI` base dictionaries are preserved during rewrite-style manipulation, complex catalog URI dictionaries are blocked with structured preflight diagnostics, and link annotation `/URI` actions are not misreported as catalog-level URI metadata.
- Catalog name trees. Supported `/Names` buckets are limited to simple `/Dests`, simple `/EmbeddedFiles`, and active-content detection; unsupported catalog name-tree buckets block rewrite so split/merge/edit helpers do not silently drop document-level structures.
- Associated files. Simple catalog `/AF` arrays are preserved during rewrite-style manipulation; complex associated-file graphs remain blocked when they reference page content or otherwise exceed the safe catalog metadata graph.
- Named destinations. Direct `/Dests` dictionaries and simple `/Names` `/Dests` name trees, including leaf `/Kids`, are preserved during rewrite-style manipulation by flattening supported name trees, with stale destinations and stale named-destination link annotations pruned when their target pages are not copied; malformed or unsupported destination name trees remain blocked until full name-tree normalization exists.
- Document metadata.
- Page labels.
- Attachments.
- Basic accessibility metadata where feasible.
- Document info inspection and editing.

Exit criteria:

- Generated reports have usable navigation and metadata.

### 8. PDF Manipulation Toolkit

This is the PSWritePDF/iText replacement track.

- Merge PDFs. Initial `PdfMerger` support exists for parser-supported PDFs, including byte-array inputs, readable stream inputs, path inputs, byte-returning path helpers, output stream writes, and enumerable file-list output to paths or streams for wrapper pipelines.
- Split by page ranges. Initial `PdfPageExtractor.ExtractPageRange` and `SplitPageRanges(..., PdfPageRange...)` support exists for parser-supported PDFs, including byte/path/stream helpers, `PdfPageRange` overloads, wrapper-friendly `PdfPageRange.ParseMany("1-3,5")` parsing, and deterministic range-file output for wrappers.
- Extract pages. Initial `PdfPageExtractor.ExtractPages`, `ExtractPageRange`, and `ExtractPageRanges` support copies selected page object graphs into a fresh PDF, including repeated selections and overlapping ranges as cloned page objects.
- Delete pages. Initial `PdfPageEditor.DeletePages`, `DeletePageRange`, and `DeletePageRanges` support exists for parser-supported PDFs, including `PdfPageRange` overloads, overlapping range-list deletion, byte-returning path helpers, and output stream writes for byte, stream, or path inputs for wrappers.
- Duplicate pages. Initial `PdfPageEditor.DuplicatePages`, `DuplicatePageRange`, and `DuplicatePageRanges` support exists for parser-supported PDFs, inserting cloned copies immediately after selected source pages, inclusive page ranges, or parsed range lists, including `PdfPageRange` overloads, repeated selections/ranges as repeated clones, byte-returning path helpers, and output stream writes for byte, stream, or path inputs for wrappers.
- Move pages. Initial `PdfPageEditor.MovePages`, `MovePageRange`, and `MovePageRanges` support exists for parser-supported PDFs, moving selected pages, inclusive page ranges, or parsed range lists as a group before another source page or to the end, including `PdfPageRange` overloads, overlap deduplication for range-list moves, byte-returning path helpers, and output stream writes for byte, stream, or path inputs for wrappers.
- Reorder pages. Initial `PdfPageEditor.ReorderPages` and `ReorderPageRanges(..., PdfPageRange...)` support exists for parser-supported PDFs, including byte-returning path helpers, output stream writes for byte, stream, or path inputs, and shared `PdfPageRange.ParseMany("3,1-2")` grammar for wrappers.
- Rotate pages. Initial `PdfPageEditor.RotatePages`, `RotatePageRange`, and `RotatePageRanges` support sets page `/Rotate` for selected pages, all pages, inclusive page ranges, or parsed range lists, including `PdfPageRange` overloads, overlap deduplication for range-list selections, byte-returning path helpers, and output stream writes for byte, stream, or path inputs for wrappers.
- Import pages. Initial `PdfPageImporter.AppendPages`, `PrependPages`, `InsertPages`, `InsertPageRange`, `AppendPageRanges`, `PrependPageRanges`, and `InsertPageRanges` support imports selected, repeated, ranged, range-list, or all source pages before, after, or inside a target PDF for parser-supported PDFs, including `PdfPageRange` overloads plus byte/path/stream helpers and output stream writes for byte, stream, or path inputs for wrappers; repeated or overlapping import ranges create cloned pages in caller order, and insert operations preserve the target document catalog/metadata anchor.
- Stamp text/images onto pages. Initial byte/path/stream `PdfStamper.StampText` support and byte/path/stream PDF plus byte/stream image `StampImage` support exists for parser-supported PDFs, including byte-returning path helpers plus output stream writes for byte, stream, and path PDF inputs, inclusive option page ranges/range lists, and emitted stamp streams through the shared internal content-stream helper.
- Watermark pages. Initial byte/path/stream `PdfStamper.WatermarkText` support and byte/path/stream PDF plus byte/stream image `WatermarkImage` support exists for parser-supported PDFs, including byte-returning path helpers plus output stream writes for byte, stream, and path PDF inputs, inclusive option page ranges/range lists, and emitted watermark streams through the shared internal content-stream helper.
- Update metadata. Initial `PdfMetadataEditor` support exists for parser-supported PDFs.
- Extract text. Initial `PdfTextExtractor.ExtractAllText`, `ExtractTextByPage`, and `ExtractTextByPageRanges` support reads all pages or selected inclusive `PdfPageRange` lists from bytes, paths, or streams, with byte/path/stream whole-document UTF-8 output to paths or caller-owned streams plus deterministic source-page-numbered text file output for wrapper pipelines, including layout-option-aware selected range output.
- Extract images by page. Initial `PdfImageExtractor` support returns page image XObjects for parser-supported PDFs and can write deterministic extracted image files from path or stream inputs; `ExtractImagesByPageRanges(..., PdfPageRange...)` adds reusable range-list selection for byte/path/stream/document inputs and selected source-page file output.
- Inspect page sizes and page count.

Exit criteria:

- PSWriteOffice can expose credible replacements for common PSWritePDF cmdlets.
- Manipulation works by copying object graphs, not by brittle byte concatenation. Extraction, split, import, delete, reorder, rotate, metadata rewrite, merge, and stamp flows preserve simple reachable URI and named-destination link annotation objects, including contents metadata, when their pages and targets are copied.

### 9. Forms And Signatures

Add only after the object engine and page manipulation are strong.

- Inspect AcroForm fields. Initial read-only simple field inventory exists through `PdfDocumentInfo.FormFields`, including fully qualified field names, field types, simple values, alternate/mapping names, and raw flags.
- Read field values.
- Set field values. Initial `PdfFormFiller.FillFields(...)` support can update simple text/string-style values, choice values supplied as export values or `/Opt` display text when available, multi-select choice arrays through `PdfFormFieldValue.FromValues(...)`, and button name values by fully qualified field name from bytes, paths, or streams, generate simple text/choice-widget normal appearance streams and simple button-widget Off/selected appearance states for widgets with `/Rect`, mark `/NeedAppearances true`, return bytes from path inputs, write path inputs to paths or caller-owned output streams, and reject signed or active-content PDFs. Choice appearances use `/Opt` display text when available while keeping the stored export value. Rich widget behavior and full appearance regeneration remain roadmap work.
- Flatten fields. Initial `PdfFormFiller.FlattenFields(...)` and `FillAndFlattenFields(...)` support paints simple text-widget appearances, simple choice-widget text appearances with `/Opt` display text when available for scalar or array selected values, and simple button-widget normal appearance states into page content, generating minimal button appearances when needed, removes those widget annotations, removes the AcroForm tree, can return bytes from path inputs or write path inputs to paths/caller-owned output streams, and rejects signed or active-content PDFs; rich/custom appearances, JavaScript actions, and complex form preservation remain roadmap work.
- Create simple text fields, checkboxes, scalar choice fields, and multi-select choice fields. Initial generated text fields, check boxes, and choice fields exist through `PdfDoc.TextField(...)`, `PdfDoc.CheckBox(...)`, `PdfDoc.ChoiceField(...)`, and `PdfDoc.MultiSelectChoiceField(...)`, including flow placement, visible normal appearance streams, catalog `/AcroForm` registration, inspector/logical readback, and fill/fill-and-flatten compatibility; radio buttons, richer appearance styling, and row/column compose placement remain roadmap work.
- Preserve unsupported form structures when possible.
- Detect security, form, navigation, tagged-structure, and catalog metadata markers to avoid unsafe work. Initial encryption/signature/form/outline/catalog-view-setting/page-label/catalog-name-tree/named-destination/open-action/viewer-preference/tagged-structure/XMP-metadata/catalog-URI/output-intent/embedded-file/optional-content/active-content marker probing is exposed through `PdfInspector.Probe`; `PdfInspector.Preflight` reports wrapper-friendly read/rewrite capability, diagnostics, structured `ReadBlockers` through `PdfReadBlockerKind`, and structured `RewriteBlockers` through `PdfRewriteBlockerKind`; signature, form, complex outline, complex page-label, unsupported catalog name-tree, malformed or unsupported named-destination name-tree, complex open-action dictionary, complex viewer-preference, tagged-structure, complex XMP metadata, complex catalog URI, complex output-intent, complex embedded/associated-file, complex optional-content, and active-content marker detection is exposed through `PdfInspector` and rejects rewrite-style manipulation before copying, merging, editing, metadata rewriting, stamping, or watermarking page content, while simple direct catalog view settings, simple outlines including simple GoTo action outline entries, simple direct page labels, direct named destinations, simple destination name trees including leaf `/Kids`, destination-array open actions, simple GoTo open-action dictionaries, simple viewer preferences, simple catalog XMP metadata streams, simple catalog URI base dictionaries, simple output intents, simple embedded-file attachment trees, simple associated-file arrays, and simple optional-content metadata are preserved.
- Signature creation is a separate later decision.

Exit criteria:

- Existing form PDFs can be filled and flattened safely for common cases.

### 10. Office Exporters

Once the core layout engine is strong, make Office formats target it.

#### Word To PDF

- Map Word paragraphs, runs, lists, tables, images, headers, footers, sections, page setup, links, bookmarks, footnotes, and simple shapes into the logical PDF model.
- Preserve unsupported content with warnings.
- Add visual regression samples for Word-authored and OfficeIMO-authored documents.
- Replace QuestPDF-backed `OfficeIMO.Word.Pdf` when output quality is comparable for supported scenarios.

#### Excel To PDF

- Start with OfficeIMO-generated report sheets.
- Support page setup, margins, orientation, print area, headers/footers, tables, merged cells, styles, images, and chart snapshots.
- Add clear diagnostics for unsupported workbook features.

#### PowerPoint To PDF

- Start with OfficeIMO-generated decks.
- Map each slide to a PDF page.
- Support text boxes, shapes, images, tables, backgrounds, and chart snapshots.
- Treat animations/transitions as non-exportable metadata.

Exit criteria:

- OfficeIMO can export its own generated documents to credible PDFs.
- Real Office-authored documents are supported incrementally with warnings.

## API Principles

- Keep the simple path fluent.
- Keep the logical model inspectable and editable.
- Do not expose raw PDF operators as the default user experience.
- Provide low-level escape hatches for advanced users.
- Preserve unsupported content during manipulation when possible.
- Emit diagnostics rather than silently dropping content.
- Make PowerShell wrappers thin over reusable .NET APIs.

## Proposed Package Shape

Start with one package until boundaries become real:

- `OfficeIMO.Pdf`

Possible future splits:

- `OfficeIMO.Pdf.VisualTests` for test-only harness helpers.
- `OfficeIMO.Word.Pdf` as the Word exporter package.
- `OfficeIMO.Excel.Pdf` as the Excel exporter package.
- `OfficeIMO.PowerPoint.Pdf` as the PowerPoint exporter package.

Do not create package splits before the internal boundaries are useful.

## PowerShell Parity Targets

These are the eventual PSWriteOffice-facing operations needed to replace PSWritePDF:

- New PDF document.
- Add text.
- Add table.
- Add image.
- Add list. Initial bulleted and numbered list blocks exist in `OfficeIMO.Pdf`, including reusable `PdfListStyle` defaults and per-list overrides for Word-like typography, indentation, marker spacing, color, rhythm, keep-together, and keep-with-next page flow.
- Add page break.
- Save PDF.
- Read PDF text.
- Get PDF metadata.
- Get PDF page count and page sizes.
- Split PDF.
- Merge PDF.
- Extract PDF pages.
- Remove PDF pages.
- Rotate PDF pages.
- Add watermark.
- Add stamp.
- Inspect PDF form fields. Initial simple AcroForm field inventory exists through `PdfInspector.Inspect(...)` / `Preflight(...).DocumentInfo.FormFields`.
- Create PDF text fields, check boxes, scalar choice fields, and multi-select choice fields. Initial generated text field, check box, and choice field creation exists through `PdfDoc.TextField(...)`, `PdfDoc.CheckBox(...)`, `PdfDoc.ChoiceField(...)`, and `PdfDoc.MultiSelectChoiceField(...)`.
- Fill PDF form. Initial simple AcroForm value fill exists through `PdfFormFiller.FillFields(...)` / `FillFieldsToBytes(...)`, including path-to-output-stream helpers.
- Flatten PDF form. Initial simple text-widget, choice-widget, and button-widget flattening exists through `PdfFormFiller.FlattenFields(...)`, `FlattenFieldsToBytes(...)`, `FillAndFlattenFields(...)`, and `FillAndFlattenFieldsToBytes(...)`, including path-to-output-stream helpers.
- Convert Word to PDF.
- Convert Excel to PDF.
- Convert PowerPoint to PDF.

## Near-Term Issue Slices

Good first issues should be small and visual:

1. Add visual regression harness and baselines for current `OfficeIMO.Pdf` examples. Initial geometry/content-stream baselines exist, and the first repo-local Poppler raster comparison lane now covers the professional report, a two-page line-item statement fixture, a Word-like table style gallery with compact Accent1-6 swatches, a landscape showcase dashboard, plus compact hello-world, core-layout, style-cheatsheet, styled-runs, tabs-leaders, drawing-gallery, row-columns, links-rules, lists-tables, default-styles, three-page flow-dsl, and two-page headers-footers scenarios with diff artifacts; next step is expanding raster coverage across the remaining runnable example set.
2. Fix paragraph spacing so generated reports no longer look stretched.
3. Fix table cell overflow and add wrapping tests.
4. Update `OfficeIMO.Pdf/README.md` with real current features and roadmap link.
5. Add dependency guard test for `OfficeIMO.Pdf`.
6. Add support matrix doc for PDF create/read/manipulate/export. Initial matrix: `Docs/officeimo.pdf.support-matrix.md`.
7. Implement PDF page count and page size inspection APIs. Initial API: `PdfInspector.Inspect`.
8. Implement split by page range using object copying. Initial API: `PdfPageExtractor.ExtractPageRange`.
9. Implement merge using page import and resource collision handling. Initial API: `PdfMerger.Merge`.
10. Implement delete and rotate page helpers. Initial API: `PdfPageEditor`.
11. Implement metadata editing helpers. Initial API: `PdfMetadataEditor`.
12. Implement generated navigation anchors, links, and outlines. Initial APIs: `PdfOptions.CreateOutlineFromHeadings` for heading outlines, generic `Bookmark(...)` flow anchors for named destinations, and paragraph `LinkToBookmark(...)` runs for internal document navigation.
13. Add professional report example that becomes the visual quality reference. Initial professional report baseline exists in `PdfDocVisualBaselineTests`, runnable example coverage exists in `OfficeIMO.Examples/Pdf/Pdf.ProfessionalReport.cs`, and Poppler-rendered PNG approval exists in `PdfDocRasterVisualBaselineTests` alongside a two-page line-item statement fixture, a Word-like table style gallery with compact Accent1-6 swatches, a landscape showcase dashboard, compact smoke, core-layout, style-cheatsheet, styled-runs, tabs-leaders, drawing-gallery, row-columns, links-rules, lists-tables, default-styles, three-page flow-dsl, and two-page headers-footers approvals.

## Non-Goals For Now

- Full PDF/A compliance.
- Full tagged PDF accessibility.
- Full browser-grade HTML/CSS rendering.
- Full SVG engine.
- Full Office fidelity for arbitrary Word/Excel/PowerPoint files.
- Digital signature creation.
- OCR.
- JavaScript in PDFs.

These may become future tracks, but they should not block the core engine.

## Release Gates

Before claiming a feature publicly:

- Unit tests pass.
- Generated PDF opens in at least two readers during local/manual validation for visual features.
- Text extraction works where expected.
- Visual baseline exists for visible output.
- Unsupported inputs produce clear diagnostics.
- The feature has at least one realistic example.
- No new runtime dependency is added to `OfficeIMO.Pdf`.

## Product Principles

- Make boring business PDFs look excellent first.
- Prefer correctness over broad shallow coverage.
- Prefer inspectable models over black-box rendering.
- Prefer small verified slices over big rewrites.
- Do not hide weak visuals behind passing text-extraction tests.
- Treat PDF manipulation and PDF generation as one shared engine, not separate hacks.
- Keep the path open for PowerShell users from day one.
