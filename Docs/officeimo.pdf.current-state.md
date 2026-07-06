# OfficeIMO PDF Current State

This is the canonical PDF product state file. Keep it current and delete dated
review snapshots instead of adding another one.

## Product Direction

`OfficeIMO.Pdf` is the first-party PDF engine for OfficeIMO. The goal is a
dependency-light, MIT-licensed PDF platform that can create, inspect, read,
convert, and safely manipulate business PDFs without Microsoft Office,
commercial PDF engines, commercial PDF runtime dependencies, or runtime
rasterizer dependencies.

The reusable PDF engine should stay in `OfficeIMO.Pdf`. Source packages such as
Word, Excel, PowerPoint, Markdown, HTML, and Reader should remain thin adapters
that map their own models into shared PDF primitives or consume the shared PDF
read model.

## Current Areas

### Core PDF Creation

Status: useful and broad, with reusable Unicode typography groundwork, still
short of full Office/browser shaping parity.

Available now:

- Fluent document creation through `PdfDocument`.
- Page size, margins, orientation, page backgrounds, watermarks, and page
  borders.
- Headings, paragraphs, rich text runs, links, tabs, lists, rows/columns,
  panels, horizontal rules, tables, images, vector drawing, and foreground
  canvas content.
- Header/footer zones, page numbers, page labels, viewer preferences, metadata,
  XMP metadata groundwork, output intents, embedded/associated files, and simple
  AcroForm fields.
- Text annotations, link annotations, open actions, catalog page mode/layout,
  URI base dictionaries, print/viewer preferences, and page labels.
- Table styling, Word-like table style names, row splitting, repeated headers,
  spans, rich cell text, links, images, fills, borders, alignment, sizing, and
  visual table fixtures.
- Shared conversion warnings and text diagnostics for unsupported WinAnsi text
  before rendering.
- Deterministic generated TrueType font embedding with used-glyph subsetting,
  Type0 `/Identity-H` Unicode writing for generated standard-font slots,
  used-glyph `/ToUnicode` maps, extractable multilingual generated text, and a
  multilingual business-report proof artifact. Cached embedded-font programs
  reset glyph usage at each write so reused options and internal writer calls
  keep subset `/W` and `/ToUnicode` data local to the document being emitted.
- Word, Excel, Markdown, HTML, and PowerPoint conversion adapters now share a
  focused multilingual typography contract test that forwards embedded
  `PdfOptions`, verifies Type0 `/Identity-H`, `/ToUnicode`, and `/FontFile2`
  output, extracts Polish/Greek/Cyrillic text, and protects deterministic
  Markdown adapter subset output across repeated runs. The visual review gallery
  also treats the five converter-generated multilingual PDFs as declared
  manifest artifacts so reviewers can inspect the same proof files that the
  adapter tests generate.
- Shared text diagnostics for WinAnsi and embedded TrueType missing-glyph
  preflight, surfaced through `PdfConversionReport`. Generated writer encoding
  failures can now be routed through `PdfOptions.ReportDiagnosticsTo(...)` so
  converter adapters record missing-glyph warnings before the existing
  fail-closed exception escapes. Embedded TrueType preflight, measurement, and
  glyph-hex encoding now share the same internal Unicode-scalar glyph-run
  shaping boundary, including stable glyph ids, text indexes, advances, and
  missing-glyph diagnostics, so later OpenType/CFF or HarfBuzz-style providers
  have one core integration point instead of scattered renderer loops.
- Shared embedded-font diagnostics can preflight configured font bytes and
  surface unsupported OpenType/CFF, TrueType collection, unknown scaler-type, or
  malformed TrueType inputs through `PdfFontDiagnostics` and
  `PdfConversionReport`. Writer-time embedded-font parse failures routed through
  `PdfOptions.ReportDiagnosticsTo(...)` now record a stable warning before the
  existing fail-closed exception escapes.
- Dependency-free OpenType inspection can now parse real single-face
  OpenType/CFF font metadata, Unicode coverage, and layout-table evidence
  through `PdfOpenTypeFontInspector`, including advertised `GSUB` and `GPOS`
  feature tags. Parseable CFF fonts now participate in generated standard-font
  output through compact embedded `/FontFile3` Type0/CIDFontType0 dictionaries,
  `/Identity-H`, glyph-width arrays, and used-glyph `/ToUnicode` mappings.
  Covered Unicode ligature presentation-form scalars such as U+FB00 through
  U+FB04 can be written and extracted through that scalar path when the
  embedded font maps them, while automatic GSUB ligature substitution remains a
  separate shaping milestone.
  Writer output routed through `PdfOptions.ReportDiagnosticsTo(...)` records a
  stable `opentype-cff-charstrings-not-subset` warning with retained/unused CFF
  glyph counts, embedded table lists, removed layout-table evidence, and
  embedded font length, so current compact-table CFF embedding stays honest
  while CFF charstring subsetting remains future work.
  Malformed CFF inputs still surface stable `unsupported-opentype-cff-font`
  diagnostics before the fail-closed writer exception escapes.
- Shared advanced text layout diagnostics can preflight right-to-left text,
  complex-script shaping needs, combining-mark/joiner shaping, and
  script-specific line breaking through `PdfTextDiagnostics` and
  `PdfConversionReport`. Font-aware overloads can also inspect configured
  TrueType/OpenType-CFF bytes and report concrete OpenType feature gaps such as
  unsupported `GSUB` ligature substitution and `GPOS` mark positioning.
  Writer-time embedded-font text output routed through
  `PdfOptions.ReportDiagnosticsTo(...)` now records those stable
  simplified-content warnings before the existing encoding/missing-glyph checks
  continue. Provider-backed text shaping suppresses those warnings only for the
  exact embedded-font text run that the configured provider actually shaped, so
  declined provider runs still report the simplified shaping diagnostics.
- Embedded TrueType/OpenType-CFF fallback planning that splits text into
  contiguous candidate-font segments, reports uncovered scalars before
  rendering, and can convert fully covered plans into styled rich `TextRun`s
  assigned to generated font slots for the existing rich text renderer.
  `PdfEmbeddedFontFallbackSet` keeps the candidate list, generated font slots,
  styled-slot registration, and planned run generation together for thin
  converter adapters. It can also analyze planned fallback segments for
  selected-font OpenType layout feature warnings, and its report-aware
  `TryPlanTextRuns(...)` path now forwards those shaping diagnostics alongside
  incomplete-plan missing-glyph diagnostics.
  `PdfParagraphBuilder.FallbackText(...)` provides the fluent paragraph path for
  covered fallback plans while preserving the builder's current style, and
  `TryPlanTextRuns(...)` lets adapters surface incomplete-plan diagnostics
  through `PdfConversionReport` without exception-driven control flow.
- Registered fallback sets now participate in ordinary rich text wrapping and
  rendering: unsupported rich `TextRun` content is split into fallback-backed
  styled runs before measurement, so Word/Excel/Markdown/HTML/PowerPoint
  adapter rich text paths can benefit without hand-planning every run.
- Provider-backed shaping diagnostics now understand those automatically
  planned fallback runs: provider-owned fallback text suppresses the matching
  complex-script/bidi warnings only after actual writer shaping, while declined
  provider runs still report the simplified shaping diagnostics.
- The same fallback set now covers generated header/footer text, table captions,
  canvas text, text watermarks, generated free-text annotation appearances, and
  generated AcroForm text/choice field appearances, keeping measurement and
  emitted font resources aligned for those generated text surfaces.
- Parser-side form fill/flatten appearance regeneration can reuse an inherited
  embedded Type0 `/Helv` AcroForm font when its decoded `/ToUnicode` CMap covers
  every drawable scalar in the visible value, including simple multiline and
  comb text appearances, preserving extractable Unicode appearances for covered
  generated forms instead of falling back to Type1 Helvetica.
- Parser-side form fill/flatten appearance regeneration can also discover
  non-`/Helv` embedded Type0 fonts from inherited AcroForm default-resource
  `/Font` dictionaries when their `/ToUnicode` CMap covers the filled value,
  and regenerated appearances select the discovered resource name consistently
  before flattening/readback.
- Parser-side form fill/flatten appearance regeneration now also checks the
  widget's previous normal appearance resources and the widget page's inherited
  resources for covered embedded Type0 fonts, so external forms with usable
  Unicode appearance fonts outside AcroForm `/DR` can keep extractable Unicode
  appearances during fill/regeneration.
- Parser-side form fill/flatten can also synthesize new embedded Type0
  `/Identity-H` appearance font resources from caller-supplied
  `PdfFormFillerOptions` appearance fonts when the parsed source field has no
  reusable embedded appearance font, including subset `/FontFile2` for TrueType,
  compact `/FontFile3` `/Subtype /OpenType` streams for CFF, glyph-width
  arrays, `/ToUnicode`, metric-aware alignment, and extractable flattened
  Unicode text.
  CFF appearance fonts routed through
  `PdfFormFillerOptions.ReportDiagnosticsTo(...)` now surface the same
  `opentype-cff-charstrings-not-subset` warning used by generated document font
  output, and configured or fallback appearance plans also forward
  selected-font shaping warnings such as unsupported ligature substitution or
  mark positioning through both fill and fill-and-flatten flows.
  Explicit `PdfEmbeddedFontFallbackSet` appearance
  fallbacks can now be registered on the same options object; parsed text
  appearances use the shared TrueType/OpenType-CFF fallback plan to split
  covered values across synthesized Type0 font resources, using `/FontFile2` for
  TrueType and `/FontFile3` `/Subtype /OpenType` for CFF, and incomplete plans
  fail with stable missing-glyph diagnostic codes instead of falling back to
  Helvetica/WinAnsi.
  `PdfFormFillerOptions.ReportDiagnosticsTo(...)` can also send those
  configured-font and fallback-set failures into `PdfConversionReport` with
  field-aware source labels. The byte-returning, path, stream, output-stream,
  fluent form fill/flatten helpers, and safe fluent `Try...` operations now all
  have option-aware entrypoints so converters can keep appearance-font
  diagnostics without first dropping to manual byte handling.
- The form-appearance proof corpus now covers text, multiline text, comb text,
  scalar choice, multi-select choice, checkbox, and radio fields in one generic
  PDF fixture, including `/NeedAppearances false`, selected widget `/AS`
  values, generated normal appearance states, preserved radio appearance
  streams, per-cell comb appearances, multi-row choice appearances, and
  flattened widget visuals.
- Parser-side form flattening can synthesize missing rich text-widget
  appearances from `/RV` for the same controlled XHTML/CSS subset used by
  FreeText rich-content flattening, preserving bold/italic/underline/color spans
  through flattened visuals when no reusable widget `/AP` exists.
- Converter-supplied hyphenation hooks now let rich text wrapping try preferred
  UTF-16 token break points, with visible hyphenated chunks before the existing
  scalar fallback is used for long unspaced tokens.
- Caller-supplied non-hyphenating line-break hooks now let generated simple and
  rich text wrapping use dictionary/script break points for long unspaced
  tokens without adding visible hyphens, and covered script-specific
  line-breaking diagnostics are suppressed only when the callback supplies valid
  break points.
- Rich and simple generated text wrapping now share built-in multilingual
  breakpoints for CJK/Kana/Hangul-style tokens, avoiding invalid surrogate
  splits and common leading closing-punctuation breaks before the emergency
  scalar fallback is used.
- Fallback-set planning can analyze the selected embedded-font segments for
  OpenType feature warnings and report-aware `TryPlanTextRuns(...)` forwards
  selected-font shaping diagnostics with incomplete-plan missing-glyph
  diagnostics. Word, Excel, Markdown, HTML, and PowerPoint PDF adapters now have
  contract coverage proving configured embedded-font OpenType feature warnings
  reach their `PdfConversionReport`; the HTML bridge links its parent report to
  the selected nested Word/Markdown PDF report so late writer diagnostics remain
  visible for both `SaveAsPdf(...)` and `ToPdfDocument(...).ToBytes()` flows.

Important gaps:

- OpenType/CFF charstring subsetting. Compact OpenType table embedding exists,
  but CFF charstrings are still retained intact and reported as such.
- Broader annotation appearance semantics and complex/rich form appearance
  regeneration beyond generated free-text, multiline/comb text, scalar and
  multi-select choice, checkbox, and radio widget appearances.
- XFA rendering/filling remains unsupported; current support is deliberate
  detection and stable metadata for routing.
- Parsed font subset extension for values not already covered by an inherited
  appearance font or explicit fallback set, plus broader parser-side discovery
  beyond AcroForm default resources, prior normal appearance resources, and
  widget page resources.
- Automatic fallback selection inside ordinary text APIs, text shaping,
  ligatures, complex script handling, and full built-in Unicode line-breaking
  parity beyond the current callback and CJK/Kana/Hangul boundary rules.
- Built-in dictionary-driven hyphenation and stronger text measurement parity.

### PDF Reading And Inspection

Status: practical for born-digital/simple PDFs, conservative for complex PDFs.

Available now:

- Probe, inspect, preflight, text extraction, structured/logical readback,
  image extraction, attachment extraction, page metadata, outline/navigation
  readback, link annotations, form widget summaries, Reader/PDF typed form
  fields with widget geometry, passive open-action, catalog, page, and
  annotation action summaries plus
  active-content diagnostics, security/revision markers,
  signature metadata, DSS/VRI evidence summaries, tagged-structure summaries,
  optional-content summaries, catalog actions, page actions, XMP metadata,
  output-intent metadata, viewer metadata, FreeText annotation `/DA`, `/DS`,
  `/RC`, resolved font-size/color/alignment, and rich-content plain-text
  summaries, and diagnostics.
- Signature structure validation reports stable proof states, byte-range
  structure, LTV/DSS readiness markers, explicit non-cryptographic validation
  limits, and append-only mutation policy evidence for wrapper and automation
  surfaces.
- Fluent reader operations now include diagnostic-safe string page-range
  overloads for text, page text, Markdown, logical models, and text blocks, so
  malformed range text and blocked documents stay inside `PdfOperationResult`
  instead of forcing callers into ad hoc exception handling.
- Fluent inspection readback now exposes document info, Info-dictionary
  metadata, XMP metadata, output intents, tagged-content metadata,
  optional-content/layer metadata, attachment metadata, security/revision
  markers, page geometry, header/effective version, and diagnostic-safe
  `TryDocumentInfo(...)`, `TryMetadata(...)`, `TryXmpMetadata(...)`,
  `TryOutputIntents(...)`, `TryTaggedContent(...)`, `TryOptionalContent(...)`,
  `TryAttachmentMetadata(...)`, `TryPages(...)`, and `TrySecurity(...)` helpers
  through the reader facade.
- Fluent image extraction and placement-geometry readback now have the same
  selected-page and diagnostic-safe string page-range shape over the existing
  image extraction engine.
- Fluent navigation and link readback now exposes outlines, page labels, named
  destinations, catalog view metadata, open actions, viewer preferences, and
  URI/destination/named-action/remote-GoTo link filters without forcing callers
  to manually traverse the logical document model.
- Fluent annotation readback now exposes generic page annotations, subtype
  filters, and action-type filters through the reader facade, with
  diagnostic-safe `TryAnnotations(...)` variants over the existing inspector
  model.
- Fluent active-action readback now exposes catalog actions, page actions, and
  filters by action type, catalog source, page number, trigger key, and stable
  action path through the reader facade, with diagnostic-safe `Try...` variants
  over the existing inspector model.
- Fluent form field and widget readback now exposes the existing logical
  AcroForm model directly, including filters by field name, kind, and page plus
  diagnostic-safe `TryFormFields(...)` and `TryFormWidgets(...)` helpers.
- AcroForm XFA packets are detected and surfaced through
  `PdfLogicalDocument.AcroFormXfa`, `PdfDocumentInfo.AcroFormXfa`, and
  Reader/PDF `pdf.form.xfa` metadata. OfficeIMO does not render, execute, or
  fill XFA packets.
- Fluent page manipulation operations now mirror that diagnostic-safe string
  page-range shape for split, extract, delete, reorder, duplicate, move, and
  rotate.
- Standard security password handling can generate encrypted PDFs, require a
  password for read/decrypt, reject wrong passwords with typed exceptions, read
  text and security metadata with a valid `PdfReadOptions.Password`, and split
  password-opened pages into unencrypted outputs while encrypted rewrites remain
  blocked by preflight.
- Logical table readback exposes `PdfLogicalTableDiagnostics` with schema,
  cell-completeness, column-geometry, and overall confidence signals, while
  logical images expose first-placement convenience geometry for adapter and
  review surfaces.
- Reader/PDF can now read from file paths, streams, byte arrays, and already
  loaded logical models into the shared read-result envelope, JSON envelope,
  normalized chunks, logical tables, and table export bundles.
- Shared conversion reports now expose a reusable summary grouped by severity,
  converter, warning code, and source area, so adapters, proof packs, wrappers,
  and UIs can route accepted degradations without hand-rolling diagnostics.
- `PdfDocumentConversionResult` now exposes reusable conversion proof snapshots
  that capture generated artifact byte counts and SHA-256 hashes, verify
  generated PDFs are readable, required page counts and page sizes match,
  required document metadata matches, required outline titles and URI link
  targets are present, required AcroForm field names are present, required
  named destinations, page-label ranges, attachment file names, output-intent
  subtypes, output-condition identifiers, optional-content/layer group names,
  default configuration name/creator/base state, visible/hidden/locked/ordered
  layer membership, catalog language, catalog page mode/layout, document open
  actions, viewer preference values, XMP title, creator, description, producer,
  keywords, subjects, PDF/A and PDF/UA identification, tagged structure types,
  structure-element counts, and marked-content references are present, required
  text markers are extractable, required logical readback signals such as
  `page-geometry`, `metadata`, `named-destinations`, `page-labels`,
  `attachments`, `output-intents`, `optional-content`, `layers`,
  `catalog-view`, `open-action`, `viewer-preferences`, `xmp`, and
  `tagged-content` are present, expected warning codes/sources are present,
  unexpected warning codes can be rejected against an
  accepted-degradation allow list, pinned artifact hashes can be enforced, and
  post-processing kept the same captured diagnostics.
- Read and rewrite blockers for unsupported or risky inputs.
- Capability flags for wrapper dispatch, including text extraction, logical
  objects, images, attachments, page manipulation, simple form fill, and simple
  flattening.

Important gaps:

- Broader encrypted rewrite preservation beyond password-backed read/decrypt
  and page extraction.
- External cryptographic signature validation and broader signature-preserving
  mutation beyond the current append-only metadata/form/signature-prep policy.
- Tagged PDF preservation beyond readback.
- Optional content/layers preservation beyond simple metadata preservation.
- Broader real-world xref-stream/object-stream corpus coverage beyond current
  readback and rewrite-preservation marker proof, plus complex metadata, name
  tree, output intent, embedded-file, active-content, and richer-form coverage.
- OCR execution, which should remain outside the dependency-light core; the
  reader now exposes OCR handoff candidates and diagnostics for image-only PDF
  pages, and `OfficeIMO.Reader` can merge external OCR text back into generic
  read-result blocks/chunks with trace metadata.

### PDF Manipulation

Status: useful for safe simple documents.

Available now:

- Split, page range extraction, merge/import, delete, duplicate, move, reorder,
  rotate, metadata editing, text/image stamp, text/image watermark, simple form
  field fill with regenerated text appearances for alignment, multiline
  wrapping, comb widgets, multi-select choice list rows,
  inherited/configured embedded fonts, fallback font sets, inherited default
  appearance font sizes/resource names from form, widget appearance, and page
  resources, gray/RGB/CMYK default appearance text colors, widget appearance
  border widths from `/BS` and `/Border`, dashed border patterns from `/BS`,
  and underline/beveled/inset border styles from `/BS /S /U`, `/B`, and `/I`,
  text/choice/check-box/radio-widget flattening, text/path/stamp annotation
  flattening with `/BS` and `/Border` border-width preservation, dashed `/BS`
  border patterns, underline `/BS /S /U` borders, beveled/inset `/BS /S /B`
  and `/I` borders, gray/RGB/CMYK FreeText `/DA` text colors, FreeText `/DS`
  default-style font size/color/alignment fallback, FreeText `/RC` rich-content
  plain-text extraction plus synthesized rich-span appearance rendering for a
  controlled XHTML/CSS subset, square/circle and FreeText `/BE /S /C` cloudy border
  effects, `/CA` opacity preservation, and FreeText `/CL` callout lines with
  `/LE` line endings plus `/RD` inner text/border rectangles for synthesized
  appearance streams, preservation of invisible/hidden/no-view annotations
  during flattening, and simple catalog preservation for copied pages.
- Append-only form field updates for signed/append-sensitive PDFs, including
  scalar text/choice/button values, multi-select choice value arrays, and
  generated simple widget appearance streams that share the multiline/list-row
  and check-box/radio button behavior used by full form fill.
- Append-only metadata revisions remain available for tagged PDFs where full
  rewrite is blocked, and preservation proof verifies the tagged structure
  survives the incremental update.
- Preflight and validation surfaces expose the shared append-only mutation
  policy, including supported metadata/form/signature-preparation actions,
  blocked actions, blocker codes, caution warnings, and summaries derived from
  the same security markers used by the incremental updater.
- Fluent metadata updates preserve simple optional-content/layer catalog state
  when preflight allows rewrite, with rewrite-preservation proof covering layer
  names, visibility, locks, order, and usage metadata.
- Stream, path, and byte helper coverage with path validation and fail-closed
  preflight behavior.
- Generic rewrite preservation proof helpers can compare original and rewritten
  PDFs for page count, geometry, metadata, outlines, named destinations,
  page-label ranges, links, annotations, forms, attachments, XMP, output
  intents, optional content, tagged structure, catalog view settings, and
  required retained text markers. Navigation proof now checks named-destination
  names, target pages, destination modes/coordinates, and page-label start
  indexes, styles, prefixes, and start numbers instead of relying only on
  counts. Viewer/action proof now checks document open-action destination
  metadata, viewer preference values, catalog action names/types/sources, and
  page action page/trigger/type/path metadata. Source-structure proof now checks
  header/catalog/effective PDF versions plus previous-revision,
  incremental-update, xref-stream, object-stream, and startxref/revision-count
  markers, and shared page/metadata/stamp rewrite paths preserve the source
  header version instead of silently falling back to PDF 1.4.
- Reusable rewrite-preservation matrix helpers can execute named PDF rewrite
  scenarios and classify each row as rewrite-safe, preservation-failed, blocked
  by safety checks, or operation-failed, keeping the underlying preservation
  report or blocker message available for proof packs and CI summaries. The
  visual review artifact summary now exports these matrix rows so reviewers can
  distinguish expected safety blockers from preservation regressions. The
  shared matrix now covers metadata preservation, source-structure drift,
  optional-content/layer drift, targeted form-fill preservation, form rewrite
  blockers, tagged-content rewrite blockers, active-content rewrite blockers,
  and signature rewrite blockers. The
  `PdfDocument` proof surface exposes thin fluent matrix helpers for normal
  document rewrite operations.
- Generic redaction verification helpers can assert that removed text markers
  are no longer extractable, present in raw rewritten bytes, or hidden in common
  PDF string encodings such as escaped literal strings and hex strings, or
  retained inside decoded PDF streams such as Flate-compressed stream content,
  while retained markers remain readable; by default they now fail closed when a
  PDF stream cannot be decoded during removed-marker verification unless callers
  explicitly opt out of that weaker proof mode.
- Redaction planning now reports intersecting image XObject placements as a
  distinct match kind with a warning diagnostic. Fully covered page-level image
  placements and safe nested form-XObject image placements are removed from the
  relevant content streams and XObject resources before the redaction mark is
  painted. Shared form aliases are cloned before nested image removal so other
  visible placements remain intact. Simple isolated partial intersections over
  8-bit DeviceGray/DeviceRGB image streams are rewritten at the pixel level
  using the configured redaction fill color, including `/Decode` arrays and
  matching indirect 8-bit DeviceGray soft masks where the redacted region is
  made opaque. Reused safe image and form invocations are isolated by cloning
  the target resource and renaming only the matched content-stream invocation
  before rewriting pixels, so sibling placements remain intact. Transformed
  placements, color-key or explicit masks, JPEG payloads, and other complex
  image streams still fail closed by default;
  callers can explicitly opt into a visual-only image overlay through
  `AllowImagePlacementOverlays`, but that mode does not claim image pixels or
  resources were removed.

Important gaps:

- Incremental update strategy.
- Full safe redaction authoring with broader geometry/resource cleanup beyond
  current marker-based removal verification and whole page-level image
  placement removal, including color-key/explicit-mask/JPEG/transformed partial
  image pixel rewriting and broader complex repeated-resource cases.
- Broader rich form appearance regeneration beyond the current controlled `/RV`
  XHTML/CSS span subset used when synthesizing missing text-widget appearances
  for flattening.
- Broader FreeText rich-content coverage beyond the current controlled
  XHTML/CSS span subset, especially embedded-font and complex-script rich
  annotation appearances.
- More complex page/resource/catalog preservation.
- Broader real-world rewrite preservation corpus coverage without corrupting
  unsupported PDFs; the reusable matrix exists, but the fixture family still
  needs more signed, encrypted, tagged, optional-content, object-stream,
  attachment-heavy, active-content, and form-heavy examples.

### Office And Document Converters

Status: product-valuable, deliberately partial.

Available now:

- `OfficeIMO.Word.Pdf`: first-party Word-to-PDF path for common sections,
  paragraphs, lists, tables, images, links/bookmarks, headers/footers, simple
  content controls, simple form controls, page numbering, and warnings for
  unsupported content.
- `OfficeIMO.Excel.Pdf`: worksheet-to-PDF path for visible sheets, print areas,
  margins/orientation, hidden row/column filtering, repeated titles, simple
  headers/footers, images, supported charts, merged cells, common formats,
  basic styling, links, sizing, and warnings.
- `OfficeIMO.PowerPoint.Pdf`: slide-to-PDF path for page-sized slide canvases,
  backgrounds, text boxes, pictures, tables, supported charts, simple shapes,
  and warnings.
- PDF logical table extraction can now write editable document tables back into
  Excel worksheets, Word tables, and PowerPoint table slides for document and
  invoice-style workflows, including PowerPoint source-proportional column
  sizing, shared logical-table numeric parsing and text/numeric/mixed/empty
  column profiles for safely typed Excel cells, Word/PowerPoint numeric
  body-cell alignment, and row/column slide segmentation for wide or long
  tables, or emit table-only Markdown/semantic HTML with page ranges and row
  caps. The visual review gallery now carries a declared reverse-direction
  proof scenario with the source PDF plus editable DOCX, XLSX, and PPTX table
  artifacts so this table-only reconstruction boundary is reviewable.
- `OfficeIMO.Word.Pdf` also has a semantic PDF-to-Word import path over the
  first-party PDF logical model, preserving parser-supported metadata, page
  breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks,
  supported internal destination links as Word bookmark hyperlinks, complete
  image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware, soft-mask-capable simple
  `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic
  `ICCBased` N=1/3/4 streams, and Decode-aware, soft-mask-capable `Indexed` palette PDF image
  streams as native Word images when their filters are supported, unresolved masked JPEG pass-through warnings, form widget placeholders, and conversion
  diagnostics in an editable `.docx` package without claiming fixed-layout page
  recreation.
- `OfficeIMO.Markdown.Pdf`: Markdown-to-PDF path for headings, outlines, rich
  inline text, links, lists, task lists, tables, code/semantic panels, callouts,
  front matter, images, themes, and warnings.
- `OfficeIMO.Html.Pdf`: semantic/document HTML-to-PDF profiles and
  semantic/positioned-review PDF-to-HTML profiles over first-party pipelines,
  including inert PDF outline navigation for review HTML.
- `OfficeIMO.Reader.Pdf`: PDF ingestion registration for `DocumentReader`
  chunks with page-aware locations, Markdown text, detected tables, image
  placeholders, links, typed form fields, passive action summaries,
  active-content diagnostics, security and metadata summaries, hashes,
  split warnings, and table column profiles where the PDF read model can expose
  them.

Current conversion matrix:

| Path | Current support | Accuracy contract today | Next proof gate |
| --- | --- | --- | --- |
| Word -> PDF | Native first-party adapter through `WordDocument.ToPdfDocument(...)` and `SaveAsPdf(...)`. | Common document sections, text, lists, tables, images, links, headers/footers, controls, fields, and warnings for unsupported content. | Expand Word-origin visual fixtures for anchored/floating layout, revisions/fields, SmartArt fallbacks, and hard equation diagnostics. |
| Excel -> PDF | Native first-party adapter through `ExcelDocument.ToPdfDocument(...)` and `SaveAsPdf(...)`. | Visible/selected sheets, print areas, repeated titles, page breaks, styles, images, supported charts, links, and warnings for skipped/simplified workbook features. | Add print-layout fixtures for fit-to-height, automatic pagination, print scaling, locale formats, drawing placement, and additional conditional formats. |
| Markdown -> PDF | Native first-party adapter through Markdown string/file/document `ToPdfDocument(...)` and `SaveAsPdf(...)` APIs. | Structured Markdown maps to PDF headings, links, lists, task lists, tables, panels, front matter, images, themes, and warnings. | Add paginated nested-panel, long technical-document, image/resource policy, and front-matter/theme fixture families. |
| HTML -> PDF | Thin bridge with semantic and document profiles over Markdown/Word pipelines. | Good for structured HTML and practical trusted print HTML; intentionally not a browser-grade CSS renderer. | Publish a declared CSS/resource subset and add profile-specific fixtures for trusted/untrusted resources, page breaks, tables, images, and link preservation. |
| PowerPoint -> PDF | Native first-party adapter through `PowerPointPresentation.ToPdfDocument(...)` and `SaveAsPdf(...)`. | One slide per page, backgrounds, text boxes, pictures, tables, supported charts, simple shapes, group-shape geometry, and warnings. | Add master/layout inheritance, theme resolution, grouped transforms, richer table style, media placeholder, and SmartArt fallback fixtures. |
| PDF -> Markdown/Reader chunks | `OfficeIMO.Pdf` logical model plus `OfficeIMO.Reader.Pdf` page-aware chunks and `PdfLogicalDocument.ToMarkdown(...)`. | Born-digital/simple PDFs can expose page text, headings, paragraphs, lists, tables, image visual geometry, links, typed form fields with widget geometry and appearance states, XFA packet detection metadata, passive document-open/catalog/page/annotation action summaries without executable payloads, warnings, table diagnostics, active-content counters, table confidence aggregates, table/image geometry coverage, selected form-widget appearance coverage, OCR handoff candidates for image-only pages, generic downstream OCR text enrichment, and stable chunk/read-result diagnostics for source/security/form/image/OCR counters. | Expand the degradation corpus with broader real-world OCR/image-only and form files. |
| PDF -> HTML | `OfficeIMO.Html.Pdf` semantic and positioned-review profiles over the logical PDF model. | Semantic export for search/indexing and positioned review HTML for page-oriented inspection; image placeholders/data URIs, outline/bookmark navigation anchors, link/form hints, XFA notice metadata, unsafe URI inert rendering, and active catalog/page/annotation action counters are available without exposing executable payloads. | Add round-trip review fixtures that compare PDF logical objects, positioned HTML geometry, embedded image policy, and additional real-world active-content/name-tree cases. |
| PDF -> editable Word/Excel/PowerPoint | Word semantic import is supported for parser-supported logical PDF objects; Excel and PowerPoint table-focused imports are supported. | Word import preserves metadata, page breaks, headings, paragraphs, lists, logical tables, safe URI hyperlinks, supported internal destination links as Word bookmark hyperlinks, complete image-file payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, Decode-aware, soft-mask-capable simple `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams, basic `ICCBased` N=1/3/4 streams, and Decode-aware, soft-mask-capable `Indexed` palette PDF image streams as native Word images when their filters are supported, unresolved masked JPEG pass-through warnings, form placeholders, and diagnostics in an editable DOCX package. Excel/PPTX remain table-focused. This is semantic reconstruction, not fixed-layout page recreation. | Add full ICC color-managed transforms and complex/unsupported PDF image stream conversion, remote/cross-document link reconstruction, broader Word document reconstruction, and Excel/PowerPoint editable reconstruction beyond tables. |

Important gaps:

- Word: anchored/floating layout, richer table fidelity, field evaluation,
  tracked revisions, SmartArt, hard equations, and authored edge cases.
- Excel: fit-to-height, automatic pagination/scaling, richer chart fidelity,
  more conditional formats, locale-specific formats, richer drawing placement,
  and print-layout edge cases.
- PowerPoint: master/layout inheritance, theme resolution, grouped transforms,
  richer text layout, richer table styles, automatic imported-table fit-to-slide
  scaling, media, and SmartArt fallbacks.
- Markdown/HTML: stronger paginated panels, declared CSS subset maturity,
  resource-policy examples, and more visual fixture families.
- Reader/PDF: broader real-world OCR/image-only files, broader real-world form
  files, and broader real-world active-content/name-tree coverage.
- Cross-converter: a shared conversion proof contract that records source
  features, emitted PDF features, warnings, raster/text/logical checks, and
  accepted degradations for each scenario.

### Visual Proof

Status: strong test lane, now repeatable as an artifact.

Available now:

- Poppler raster baselines for core PDF authoring plus Word, Excel, Markdown,
  and PowerPoint scenarios.
- `Build/Export-PdfVisualReviewGallery.ps1` exports the current visual review
  PDFs and an index under `artifacts/pdf-visual-review` by default.
- `.github/workflows/pdf-visual-review-gallery.yml` generates the same gallery
  in CI for PDF/converter changes and uploads the review index, scenario
  manifest, proof summary, PDFs, positioned-review HTML, and editable
  reverse-conversion Office artifacts as a PR artifact. CI artifact generation
  skips strict Poppler raster comparison by default so reviewers still receive
  proof bundles when visual baselines drift; manual dispatch can opt into the
  strict raster lane.
- The gallery now includes a first invoice/statement proof generated through
  the Markdown-to-PDF adapter, with line-item and summary tables, right-aligned
  numeric columns, payment-term list items, and explicit readback markers.
- The dashboard proof is now covered by `excel-dashboard-report.pdf`, generated
  from a workbook with a chart snapshot, worksheet image, conditional
  formatting, header/footer, print area, and fit-to-page setup.
- The PowerPoint proof now includes `powerpoint-layout-theme-groups.pdf`, which
  exercises custom slide size, theme colors, a gradient background, text boxes,
  grouped shapes, and a group transform with raw PDF geometry checks.
- The readback proof now includes `pdf-logical-diagnostics-source.pdf` and a
  positioned review HTML companion, with explicit logical-image placement
  geometry, named numeric table-column profiles, and
  `PdfLogicalTableDiagnostics` confidence checks. Its summary artifact also
  records Reader/PDF image visual geometry, table confidence, and chunk-level
  table/image geometry coverage.
- The Reader/PDF degradation corpus now includes
  `pdf-reader-degradation-corpus.pdf` and a JSON accepted-degradation summary,
  proving readable page text, a safe URI link, typed form metadata, passive
  document-open, page, and annotation action summaries plus active-content
  diagnostics without emitting script
  payload text.
- The Reader/PDF hostile-action corpus now includes
  `pdf-reader-hostile-action-corpus.pdf`, positioned review HTML, and a JSON
  accepted-degradation summary, proving document-open, nested catalog
  JavaScript name trees, catalog additional actions, catalog chained
  ImportData/Movie actions, page JavaScript/Launch, annotation
  SubmitForm/Launch, and chained ImportData actions are counted as passive
  diagnostics with stable action paths and kept inert in Reader text, Reader
  JSON, and HTML output.
- The Reader/PDF hostile-layout corpus now includes
  `pdf-reader-hostile-layout-corpus.pdf` and a JSON accepted-degradation
  summary, proving close-column text readback, rotated text readback, and
  non-axis-aligned image geometry without claiming perfect reading order or
  editable layout reconstruction.
- The Reader/PDF hostile-table corpus now includes
  `pdf-reader-hostile-table-corpus.pdf` and a JSON accepted-degradation
  summary, proving headerless/jittered table-like bands can surface fallback
  columns, numeric-column hints, table geometry, and less-than-perfect
  confidence diagnostics without claiming editable spreadsheet reconstruction.
- The Reader/PDF OCR-handoff corpus now includes
  `pdf-reader-ocr-handoff-corpus.pdf` and a JSON accepted-degradation summary,
  proving image-only pages surface stable OCR candidates, linked image asset
  geometry, `ocr-needed` diagnostics, and `pdf.ocr` metadata counts without
  running OCR inside the dependency-light core. The same proof now applies a
  simulated external OCR response through `OfficeIMO.Reader`, showing callers can
  merge recognized text back as generic `ocr-text` blocks/chunks with
  `reader.ocr` trace metadata and resolved-candidate diagnostics.
- The Reader/PDF XFA form corpus now includes
  `pdf-reader-xfa-form-corpus.pdf` and a JSON accepted-degradation summary,
  proving AcroForm `/XFA` packet arrays surface packet names, stream counts,
  payload byte counts, and template/datasets flags through the logical model,
  inspector, and Reader/PDF metadata without claiming XFA rendering or filling.
- The Standard security proof now includes
  `pdf-standard-security-roundtrip.pdf`, an unencrypted extracted-page proof,
  and a JSON summary showing password-required read blockers, wrong-password
  failure, valid-password text extraction/security metadata, and fail-closed
  encrypted rewrite limits.
- The HTML proof now includes `html-css-resource-policy.pdf`, using the trusted
  document profile to allow a local stylesheet, embed a data URI image, and
  report a blocked remote stylesheet through the shared conversion report. Its
  summary artifact records the declared stylesheet/image resource policy.
- The HTML/PDF round-trip proof now includes
  `html-pdf-roundtrip-source.pdf`, semantic HTML, positioned review HTML, and a
  JSON export summary. This proves the declared profile contracts, logical
  PDF-to-HTML preservation, positioned image/link review hints, and explicit
  renderer boundaries without claiming browser-grade or pixel-perfect HTML
  rendering.
- The PDF-to-HTML positioned review proof now includes an XFA form source PDF,
  positioned review HTML, summary counters, and the `AcroFormXfaDetected`
  warning, proving XFA packets are exposed as inert review metadata without
  claiming XFA rendering or filling.
- The editable Office import proof now includes
  `pdf-semantic-import-word.docx` plus a JSON warning summary for semantic
  PDF-to-Word import, proving editable metadata, page boundaries, headings,
  paragraphs, lists, safe URI hyperlinks, internal bookmark hyperlinks, native
  Word image embedding for complete image payloads with transparency-mask fidelity metadata, supported `ImageMask` stencil streams, color-key masked simple and `Indexed` streams, simple
  Decode-aware, soft-mask-capable `DeviceGray`/`DeviceRGB`/basic-converted `DeviceCMYK` streams,
  basic `ICCBased` N=1/3/4 streams, and Decode-aware, soft-mask-capable `Indexed` palette image streams when their filters are
  supported, unresolved masked JPEG pass-through warnings, logical tables, and form placeholders alongside the table-focused DOCX/XLSX/PPTX
  artifacts.

Important gaps:

- More multilingual, compliance, form-heavy, additional dashboard,
  invoice/statement, technical document, slide, spreadsheet, and external
  corpus scenarios.
- A release workflow that attaches the same gallery reviewers inspect locally.

### Compliance Proof

Status: honest groundwork, no formal compliance claims yet.

Available now:

- `PdfOptions.ComplianceProfile` exposes planned intent for PDF/A, PDF/UA,
  Factur-X, and ZUGFeRD, but non-`None` profiles intentionally fail until formal
  generation is validator-backed.
- PDF/A, PDF/UA, and e-invoice metadata/output-intent/attachment/tagging
  primitives exist as groundwork.
- `PdfComplianceAnalyzer` and `PdfDocument.AssessCompliance(...)` report
  readiness and unsupported/missing requirements.
- `PdfDocument.AssessComplianceProof(...)` now combines generated-document
  readiness evidence with caller-supplied external validator results, so proof
  reports account for actual generated font/image/drawing/form evidence before
  any conformance claim can become claimable.
- `PdfComplianceGateTests` can run optional veraPDF, PDF/UA validator, and
  Mustang-style validator commands.
- `Build/Export-PdfComplianceProof.ps1` emits generated groundwork PDFs,
  validator diagnostics, expected-status metadata, a profile proof matrix,
  schema v3 `officeimo-profile-proof-contract.json`, `index.md`, and
  `proof.json`.
- Product proof contract rows now include `externalValidatorProofs` for every
  required validator family, with stable Missing/NotRun/Passed/Failed/Error
  status, satisfaction, claim-blocking, diagnostic, profile, and exit-code
  fields. The proof exporter overlays real validator diagnostics into those
  rows before writing the final proof pack.
- `.github/workflows/pdf-compliance-proof.yml` validates and uploads the proof
  pack for PDF compliance changes, with manual strict validator inputs. The CI
  assertion script now enforces the schema v3 product contract and the
  per-profile validator proof rows.

Important gaps:

- A selected CI PDF/UA validator package.
- Formal validator-passing PDF/A profile generation.
- Formal validator-passing PDF/UA generation.
- Formal Factur-X/ZUGFeRD generation with Mustang/schema/business-rule proof.
- User-facing conformance badges that are driven only by passing proof.

## Current Guardrails

- Do not claim PDF/A, PDF/UA, Factur-X, or ZUGFeRD conformance until internal
  readiness and external validator proof both pass.
- Keep `OfficeIMO.Pdf` runtime dependency-light. Rasterizers and validators
  belong in test/dev/proof lanes, not runtime packages.
- Keep converter packages thin. Move reusable layout, diagnostics, drawing,
  typography, proof, and parsing behavior into `OfficeIMO.Pdf` or
  `OfficeIMO.Drawing`.
- Prefer diagnostics and fail-closed behavior over silent corruption or silent
  content loss.
- Add visual proof for visible output, not only text/assertion tests.

## Missing Premium Work

The main premium gaps, in priority order:

1. Typography and text layout: OpenType/CFF, shaping, fallback, complex-script
   handling, multilingual fixtures beyond generated TrueType, and stronger
   extraction-safe text foundations.
2. Validator-backed compliance: convert readiness into proof for one narrow
   profile before enabling any formal conformance switch.
3. Visual proof productization: make review galleries and proof packs standard
   CI/release artifacts.
4. Real-world parser preservation: safely preserve more structures and expand
   manipulation only where rewrite proof exists.
5. Forms, annotations, and redaction: move beyond simple fill/flatten/stamp
   workflows toward richer appearances, stronger annotation behavior, and real
   redaction guarantees.
6. Converter fidelity: deepen Word, Excel, PowerPoint, Markdown, HTML, and
   Reader paths after shared typography/proof foundations improve.
7. Fluent processing ergonomics: keep growing the `PdfDocument.Open(...)`
   workflow into the one obvious path for read, inspect, split, merge, stamp,
   watermark, metadata, form-fill, flatten, and conversion proof hand-off.

## Proposed Goals

### 1. Typography Follow-Up

The first shared text foundation is now in place:

- deterministic generated TrueType embedded-font output,
- used-glyph TrueType subsetting for generated standard-font slots,
- Unicode writing beyond WinAnsi through Type0 `/Identity-H`,
- missing-glyph failures instead of silent replacement,
- embedded TrueType missing-glyph preflight diagnostics,
- embedded-font format diagnostics for unsupported OpenType/CFF and malformed
  TrueType inputs,
- dependency-free OpenType/CFF metadata and Unicode coverage inspection with a
  real Source Serif 4 OTF fixture,
- dependency-free OpenType/CFF font-program metrics, glyph-id encoding, compact
  embedded `/FontFile3` output, CIDFontType0 descendant dictionaries, and
  extractable `/ToUnicode` tests with the same real OTF fixture,
- report-visible compact OpenType/CFF embedding diagnostics while CFF
  charstring rewriting and subsetting remain roadmap work,
- advanced text layout diagnostics for right-to-left, complex-script,
  mark/joiner, and script-specific line-breaking inputs,
- embedded TrueType/OpenType-CFF fallback planning for converter-side run
  splitting,
- a renderable rich-run bridge for fully covered fallback plans,
- reusable fallback sets that register complete styled font slot families before
  planned runs render,
- a fluent paragraph fallback helper that preserves current run styling,
- try-plan fallback run generation with report-backed missing-glyph diagnostics,
- automatic rich-text fallback splitting at the shared measurement/rendering
  boundary,
- automatic generated header/footer text fallback splitting with matching
  measurement and page font resources,
- automatic table-caption, canvas text, text-watermark, generated free-text
  annotation appearance, and generated AcroForm text/choice field appearance
  fallback splitting with matching measurement and font resources,
- parser-side form fill/flatten reuse of inherited embedded Type0 AcroForm fonts
  for simple, multiline, and comb Unicode values covered by the existing
  `/ToUnicode` map,
- parser-side form fill/flatten discovery of non-`/Helv` embedded Type0
  AcroForm default-resource fonts whose `/ToUnicode` map covers the filled
  value,
- parser-side form fill/flatten discovery of covered embedded Type0 fonts from
  existing widget normal appearance resources and widget page inherited
  resources,
- parser-side form fill/flatten synthesis of configured embedded Type0
  appearance fonts for Unicode values not covered by the source PDF,
- parser-side form fill/flatten use of explicit embedded TrueType/OpenType-CFF
  fallback sets for covered Unicode values not handled by the preferred
  appearance font, including multi-resource appearance streams and missing-glyph
  diagnostic failures,
- converter-supplied rich text hyphenation hooks for preferred long-token break
  points before scalar fallback,
- caller-supplied non-hyphenating line-break hooks for dictionary/script token
  breaks in generated simple and rich text,
- shared rich/simple multilingual breakpoints for CJK/Kana/Hangul-style tokens
  before scalar fallback,
- internal glyph-run shaping boundary used by embedded TrueType preflight,
  measurement, and glyph encoding, plus a parallel OpenType/CFF program path for
  generated writer output and later HarfBuzz-style providers,
- multilingual business-report proof with extractable Polish, Greek, and
  Cyrillic text,
- shared text diagnostics surfaced through `PdfConversionReport`.

Next typography exit criterion: add parsed subset extension for values not
already covered by an inherited font, extend parser-side font discovery beyond
AcroForm default resources, prior normal appearance resources, and widget page
resources, extend fallback coverage to broader annotation/form appearance
semantics, add deterministic OpenType/CFF charstring subsetting, add
provider-backed HarfBuzz-style shaping/fallback and full Unicode line breaking
behind the glyph-run boundary, then prove one
complex-script business report without silent missing glyphs.

Exit criterion: a multilingual business report can be generated with embedded,
subsetted, extractable text, visual proof, and no silent missing glyphs.

### Goal 2. Deepen The Shared Layout Engine

Make the reusable document/layout model carry more of the work before converter
adapters grow:

- improve table measurement for mixed spans, oversized rows, repeated headers,
  footers, captions, and row/column-contained tables,
- harden keep-together, keep-with-next, widow/orphan, and column-flow behavior
  across paragraphs, headings, panels, lists, tables, images, drawings, and rows,
- add layout diagnostics that point to the exact block, row, column, or style
  that made a page impossible,
- improve canvas/layout interop so PowerPoint-like absolute content and
  document-flow content can share table, text-box, image, link, and shape
  primitives,
- make hyphenation and advanced measurement pluggable without adding a runtime
  dependency to the core package.

Exit criterion: the same shared primitives can render a dense business report,
a spreadsheet-like statement, and a slide-like page with fewer adapter-specific
layout branches.

### Goal 3. Expand Safe Parser And Rewrite Preservation

Grow manipulation only where preservation proof exists:

- add a curated corpus for signed, encrypted, tagged, optional-content,
  attachment-heavy, form-heavy, xref-stream, object-stream, and incremental PDFs,
- classify each fixture as read-only, rewrite-safe, blocked, or future,
- keep the current source-structure preservation proof as the gate for version,
  xref-stream, object-stream, and incremental-marker regressions while the corpus
  grows,
- route curated fixture checks through the reusable
  `PdfRewritePreservationMatrix` so proof packs can distinguish expected
  blockers from real preservation failures,
- preserve simple tagged/optional-content/output-intent/name-tree structures
  only when tests prove copied output remains valid,
- expand credential-aware encrypted read behavior and encrypted preservation
  proofs without weakening the dependency-light runtime boundary,
- keep fail-closed blockers for active content, signatures, and unsupported
  catalog structures until the engine can preserve them safely.

Exit criterion: rewrite helpers have an explicit fixture-backed preservation
matrix and fail closed for unsupported documents.

### Goal 4. Strengthen Forms, Annotations, And Redaction

Turn the current form/annotation groundwork into safer document-editing
capability:

- improve appearance regeneration for text, choice, check box, and radio fields,
- preserve or regenerate field resources predictably during fill/flatten flows,
- extend annotation creation and flattening through the same rendering helpers
  used by generated documents,
- add text highlight, strikeout, underline, caret, stamp, and free-text
  behavior only when geometry and extraction stay predictable,
- design safe redaction as a separate engine feature with resource cleanup and
  extraction proof, not as a visual-only overlay.

Exit criterion: common fill/flatten/annotate workflows round-trip through
OfficeIMO.Pdf without silent appearance loss, and redaction cannot ship until
removed text/images/resources are proven unrecoverable.

### Goal 5. Productize HTML/PDF And Reader/PDF Engine Contracts

Make the adapter lanes clear and safe:

- document semantic vs document HTML-to-PDF profiles,
- document semantic vs positioned-review PDF-to-HTML profiles,
- declare the supported CSS/resource subset,
- add trusted/untrusted examples,
- publish Reader.Pdf chunk metadata expectations,
- expand the small PDF/HTML corpus with accepted degradation notes,
- document security and active-content handling for Reader/PDF and PDF-to-HTML
  review workflows.

Exit criterion: users can choose the right HTML/PDF/Reader path and understand
where fidelity is guaranteed, simplified, or unsupported.

### 4. Cross-Converter Accuracy Proof

Make conversion accuracy observable instead of implied:

- define a shared conversion scenario manifest for Word, Excel, Markdown, HTML,
  PowerPoint, PDF logical readback, and PDF-to-HTML. The first manifest now
  lives in `Docs/pdf-conversion-scenarios.json`,
- record source feature coverage, expected simplifications, warnings, output
  hashes, extracted text, logical objects, raster pages, and optional validator
  evidence,
- make `Build/Export-PdfVisualReviewGallery.ps1` emit machine-readable scenario
  metadata beside the current PDF gallery and publish that gallery through the
  `PDF Visual Review Gallery` workflow so reviewers inspect the same PDFs that
  the tests generated,
- keep editable PDF-to-Office reconstruction out of scope until logical readback
  has enough table, coordinate, image, and form evidence to make it honest.

Exit criterion: every supported converter has at least one rich source fixture
with text, logical, warning, and visual proof, and unsupported/simplified
features are explicit in the proof metadata.

### 5. Fluent Processing And Building

Improve the end-user PDF workflow without duplicating converter logic:

- keep `PdfDocument.Open(...)` as the fluent processing entry point for existing
  PDFs,
- expose fluent read/inspect helpers from the existing logical model instead of
  adding another parser facade,
- route split, merge, page edit, metadata, stamping, watermarking, form fill,
  and flattening through the same opened-document workflow where possible,
- let converter packages return `PdfDocument` plus `PdfConversionReport` so
  callers can compose conversion, proof, and post-processing in one pipeline,
- add examples that chain conversion and processing, such as Word -> PDF ->
  metadata -> stamp -> save, while keeping the reusable implementation in
  `OfficeIMO.Pdf`.

Exit criterion: common PDF workflows can be written as one readable fluent
pipeline while preserving diagnostics and fail-closed rewrite behavior.

### 6. One Narrow Compliance Claim

After typography/proof improves, choose one narrow profile and make it pass:

- start with the smallest PDF/A profile that can be generated honestly,
- wire required validator evidence,
- map failures back to actionable requirements,
- flip formal profile generation only when proof is green.

Exit criterion: one generated OfficeIMO.Pdf profile can be claimed from internal
readiness plus passing external validation.

### Goal 7. Deepen Converter Fidelity Through Shared Primitives

After the shared engine work above, spend fidelity effort where it benefits all
adapters:

- Word: anchored/floating layout, richer table fidelity, field evaluation,
  tracked-revision policy, SmartArt fallback, and hard equation handling.
- Excel: fit-to-height, automatic pagination/scaling, richer conditional
  formats, locale-specific formats, chart fidelity, and drawing placement.
- PowerPoint: master/layout inheritance, theme resolution, grouped transforms,
  richer text layout, richer table styles, automatic imported-table fit-to-slide
  scaling, media, and SmartArt fallbacks.
- Markdown/HTML: declared CSS subset, stronger paginated panels, resource-policy
  examples, and broader visual fixture families.
- Reader/PDF: broader real-world OCR/image-only files, broader real-world form
  files, and broader real-world active-content/name-tree cases exposed in stable
  chunk contracts.

Exit criterion: each converter improvement either lands in `OfficeIMO.Pdf` or
`OfficeIMO.Drawing` as reusable behavior first, or documents why it is genuinely
adapter-specific.

## Documentation Rule

Keep this file as the single PDF roadmap/state document. Avoid dated review
files under `Docs/reviews` for PDF state. If the current state changes, update
this file and the relevant package README instead.
