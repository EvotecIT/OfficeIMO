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

Status: useful and broad, with first embedded TrueType/Unicode support, still
not premium typography.

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
  output through full embedded `/FontFile3` Type0/CIDFontType0 dictionaries,
  `/Identity-H`, glyph-width arrays, and used-glyph `/ToUnicode` mappings.
  Covered Unicode ligature presentation-form scalars such as U+FB00 through
  U+FB04 can be written and extracted through that scalar path when the
  embedded font maps them, while automatic GSUB ligature substitution remains a
  separate shaping milestone.
  Writer output routed through `PdfOptions.ReportDiagnosticsTo(...)` records a
  stable `opentype-cff-full-font-embedded` warning with glyph and font-length
  details so the current full-font CFF embedding limitation is visible instead
  of silent while CFF charstring subsetting remains future work.
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
  continue.
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
  full `/FontFile3` `/Subtype /OpenType` streams for CFF, glyph-width arrays,
  `/ToUnicode`, metric-aware alignment, and extractable flattened Unicode text.
  CFF appearance fonts routed through
  `PdfFormFillerOptions.ReportDiagnosticsTo(...)` now surface the same
  `opentype-cff-full-font-embedded` warning used by generated document font
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
- Converter-supplied hyphenation hooks now let rich text wrapping try preferred
  UTF-16 token break points, with visible hyphenated chunks before the existing
  scalar fallback is used for long unspaced tokens.
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

- OpenType/CFF subsetting.
- Broader annotation/form appearance semantics beyond generated free-text,
  text-field, and choice-field appearances.
- Parsed font subset extension for values not already covered by an inherited
  appearance font or explicit fallback set, plus broader parser-side discovery
  beyond AcroForm default resources, prior normal appearance resources, and
  widget page resources.
- Automatic fallback selection inside ordinary text APIs, text shaping,
  ligatures, complex script handling, and full Unicode
  line-breaking parity beyond the built-in CJK/Kana/Hangul boundary rules.
- Dictionary-driven hyphenation and stronger text measurement parity.

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
  output-intent metadata, viewer metadata, and diagnostics.
- Logical table readback exposes `PdfLogicalTableDiagnostics` with schema,
  cell-completeness, column-geometry, and overall confidence signals, while
  logical images expose first-placement convenience geometry for adapter and
  review surfaces.
- Read and rewrite blockers for unsupported or risky inputs.
- Capability flags for wrapper dispatch, including text extraction, logical
  objects, images, page manipulation, simple form fill, and simple flattening.

Important gaps:

- Encryption/decryption with supplied credentials.
- Signature validation and signature-preserving policy.
- Tagged PDF preservation beyond readback.
- Optional content/layers preservation beyond simple metadata preservation.
- Broader xref stream, object stream, complex metadata, name tree, output
  intent, embedded-file, active-content, and richer-form coverage.
- OCR, which should remain outside the dependency-light core.

### PDF Manipulation

Status: useful for safe simple documents.

Available now:

- Split, page range extraction, merge/import, delete, duplicate, move, reorder,
  rotate, metadata editing, text/image stamp, text/image watermark, simple form
  field fill, text/choice/button-widget flattening, text/path/stamp annotation
  flattening, and simple catalog preservation for copied pages.
- Stream, path, and byte helper coverage with path validation and fail-closed
  preflight behavior.

Important gaps:

- Incremental update strategy.
- Safe redaction with proof that removed content is not recoverable.
- Rich form appearance regeneration.
- More complex page/resource/catalog preservation.
- Broader real-world rewrite preservation without corrupting unsupported PDFs.

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
- `OfficeIMO.Markdown.Pdf`: Markdown-to-PDF path for headings, outlines, rich
  inline text, links, lists, task lists, tables, code/semantic panels, callouts,
  front matter, images, themes, and warnings.
- `OfficeIMO.Html.Pdf`: semantic/document HTML-to-PDF profiles and
  semantic/positioned-review PDF-to-HTML profiles over first-party pipelines.
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
| PDF -> Markdown/Reader chunks | `OfficeIMO.Pdf` logical model plus `OfficeIMO.Reader.Pdf` page-aware chunks and `PdfLogicalDocument.ToMarkdown(...)`. | Born-digital/simple PDFs can expose page text, headings, paragraphs, lists, tables, image visual geometry, links, typed form fields with widget geometry and appearance states, passive document-open/catalog/page/annotation action summaries without executable payloads, warnings, table diagnostics, active-content counters, table confidence aggregates, table/image geometry coverage, selected form-widget appearance coverage, and stable chunk diagnostics for source/security/form counters. | Expand the degradation corpus with hostile-layout cases, richer form appearance rendering semantics, and OCR hand-off boundaries. |
| PDF -> HTML | `OfficeIMO.Html.Pdf` semantic and positioned-review profiles over the logical PDF model. | Semantic export for search/indexing and positioned review HTML for page-oriented inspection; image placeholders/data URIs and link/form hints are available. | Add round-trip review fixtures that compare PDF logical objects, positioned HTML geometry, embedded image policy, and unsafe-action handling. |
| PDF -> editable Word/Excel/PowerPoint | Not a supported reconstruction path yet. | The current truthful path is PDF -> logical model/Markdown/HTML/Reader chunks, not editable Office package recovery. | Only add editable reconstruction after the logical model has stable table, coordinates, images, form metadata, and source diagnostics; start with Word-like document reconstruction before spreadsheet or slide reconstruction. |

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
- Reader/PDF: richer form appearance semantics, hostile-layout corpus cases,
  and OCR hand-off boundaries.
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
- The HTML proof now includes `html-css-resource-policy.pdf`, using the trusted
  document profile to allow a local stylesheet, embed a data URI image, and
  report a blocked remote stylesheet through the shared conversion report. Its
  summary artifact records the declared stylesheet/image resource policy.

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
- `PdfComplianceGateTests` can run optional veraPDF, PDF/UA validator, and
  Mustang-style validator commands.
- `Build/Export-PdfComplianceProof.ps1` emits generated groundwork PDFs,
  validator diagnostics, expected-status metadata, a profile proof matrix,
  `officeimo-profile-proof-contract.json`, `index.md`, and `proof.json`.
- `.github/workflows/pdf-compliance-proof.yml` validates and uploads the proof
  pack for PDF compliance changes, with manual strict validator inputs.

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
4. Forms, annotations, and redaction: move beyond simple fill/flatten/stamp
   workflows toward richer appearances, stronger annotation behavior, and real
   redaction guarantees.
5. Validator-backed compliance: convert readiness into proof for one narrow
   profile before enabling any formal conformance switch.
6. Converter fidelity: deepen Word, Excel, PowerPoint, Markdown, HTML, and
   Reader paths after shared typography/proof foundations improve.
6. Fluent processing ergonomics: keep growing the `PdfDocument.Open(...)`
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
- dependency-free OpenType/CFF font-program metrics, glyph-id encoding, full
  embedded `/FontFile3` output, CIDFontType0 descendant dictionaries, and
  extractable `/ToUnicode` tests with the same real OTF fixture,
- report-visible full-font OpenType/CFF embedding diagnostics while CFF
  charstring subsetting remains roadmap work,
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
semantics, add deterministic OpenType/CFF subsetting, add provider-backed
HarfBuzz-style shaping/fallback and full Unicode line breaking behind the
glyph-run boundary, then prove one
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
- preserve simple tagged/optional-content/output-intent/name-tree structures
  only when tests prove copied output remains valid,
- define credential-aware encrypted read behavior without weakening the
  dependency-light runtime boundary,
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
- Reader/PDF: richer form appearance rendering semantics and hostile-layout
  cases exposed in stable chunk contracts.

Exit criterion: each converter improvement either lands in `OfficeIMO.Pdf` or
`OfficeIMO.Drawing` as reusable behavior first, or documents why it is genuinely
adapter-specific.

## Documentation Rule

Keep this file as the single PDF roadmap/state document. Avoid dated review
files under `Docs/reviews` for PDF state. If the current state changes, update
this file and the relevant package README instead.
