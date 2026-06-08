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
- Shared conversion warnings and option-aware text diagnostics for unsupported
  WinAnsi text, embedded-font coverage gaps, and control characters. The
  document-level `AnalyzeTextEncoding()` preflight, generated writer, and
  non-throwing byte/save results use the same diagnostics for richer text
  encoding failures, including generated document locations and structured
  rich-text run indexes, table-cell row/column indexes, generated form field names,
  page numbers for page-scoped content, and machine-readable encoding/remediation
  for affected blocks, table cells, table captions, canvas items, canvas
  table captions, and header/footer text. Strict `ToBytes()` fails before rendering with the full
  preflight diagnostics payload, and non-throwing output returns the full
  preflight list before rendering generated documents.
- Optional standard-font ToUnicode maps and initial full-file TrueType
  embedding through `PdfOptions.EmbedStandardFont(...)`,
  `PdfOptions.UseFontFamily(...)`, and fluent `PdfDocument.UseFontFamily(...)`.
  Embedded font mappings feed generated text encoding, extraction, measurement,
  wrapping, headers/footers, forms, watermarks, and table sizing.

Important gaps:

- Standard-font generated text remains WinAnsi unless callers opt into embedded
  fonts.
- Font subsetting; embedded TrueType output currently embeds full font files.
- OpenType/CFF support.
- Text shaping, ligatures, glyph fallback, complex script handling, and
  multilingual line breaking.
- Hyphenation hooks and stronger text measurement parity.

### PDF Reading And Inspection

Status: practical for born-digital/simple PDFs, conservative for complex PDFs.

Available now:

- Probe, inspect, preflight, text extraction, structured/logical readback,
  image extraction, attachment extraction, page metadata, outline/navigation
  readback, link annotations, form widget summaries, security/revision markers,
  signature metadata, DSS/VRI evidence summaries, tagged-structure summaries,
  optional-content summaries, catalog actions, page actions, XMP metadata,
  output-intent metadata, viewer metadata, and diagnostics.
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
  caps.
- `OfficeIMO.Markdown.Pdf`: Markdown-to-PDF path for headings, outlines, rich
  inline text, links, lists, task lists, tables, code/semantic panels, callouts,
  front matter, images, themes, and warnings.
- `OfficeIMO.Html.Pdf`: semantic/document HTML-to-PDF profiles and
  semantic/positioned-review PDF-to-HTML profiles over first-party pipelines.
- `OfficeIMO.Reader.Pdf`: PDF ingestion registration for `DocumentReader`
  chunks with page-aware locations, Markdown text, detected tables, image
  placeholders, links, form summaries, security and metadata summaries, hashes,
  split warnings, and table column profiles where the PDF read model can expose
  them.

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
- Reader/PDF: richer coordinates, image placement metadata, table confidence,
  form metadata, source diagnostics, security/active-content policy examples,
  and a real-world PDF corpus.

### Visual Proof

Status: strong test lane, now repeatable as an artifact.

Available now:

- Poppler raster baselines for core PDF authoring plus Word, Excel, Markdown,
  and PowerPoint scenarios.
- `Build/Export-PdfVisualReviewGallery.ps1` exports the current visual review
  PDFs and an index under `artifacts/pdf-visual-review` by default.

Important gaps:

- CI artifact upload for the visual gallery.
- More multilingual, compliance, form-heavy, dashboard, invoice/statement,
  technical document, slide, spreadsheet, and hostile-layout scenarios.
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

1. Typography and text layout: harden embedded TrueType output with subsetting,
   fallback, OpenType/CFF planning, shaping boundaries, multilingual fixtures,
   and extraction-safe text.
2. Flow and layout engine depth: improve pagination, keeps, table/row-column
   measurement, canvas/layout interop, and shared primitives that all converters
   can reuse.
3. Parser and rewrite preservation: safely preserve more structures and expand
   manipulation only where rewrite proof exists.
4. Forms, annotations, and redaction: move beyond simple fill/flatten/stamp
   workflows toward richer appearances, stronger annotation behavior, and real
   redaction guarantees.
5. Validator-backed compliance: convert readiness into proof for one narrow
   profile before enabling any formal conformance switch.
6. Converter fidelity: deepen Word, Excel, PowerPoint, Markdown, HTML, and
   Reader paths after shared typography/proof foundations improve.

## Proposed Goals

These goals are based on current `master` after `6e1a4edd` / PR `#1894`
(`Expand OfficeIMO PDF capabilities`). They prioritize reusable engine
capability first. Visual/compliance proof remains required evidence for risky
changes, but it is not the product goal by itself.

### Goal 1. Harden Embedded-Font Typography

Build on the current full-file TrueType embedding instead of replacing it:

- add deterministic font subsetting for used glyphs,
- keep full-file embedding available as a diagnostic mode,
- expand generated-font diagnostics for missing glyphs and fallback choices,
- define a shaping boundary that can later support HarfBuzz-like behavior
  without adding a runtime dependency to the core package,
- add multilingual visual baselines for Latin extended, Greek, Cyrillic,
  symbols, right-to-left smoke text, and non-BMP characters,
- surface typography warnings consistently through conversion reports.

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
- build a small real-world PDF/HTML corpus with accepted degradation notes,
- document security and active-content handling for Reader/PDF and PDF-to-HTML
  review workflows.

Exit criterion: users can choose the right HTML/PDF/Reader path and understand
where fidelity is guaranteed, simplified, or unsupported.

### Goal 6. Make One Narrow Compliance Claim

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
- Reader/PDF: richer table confidence, image geometry, form metadata, source
  diagnostics, and security metadata exposed in stable chunk contracts.

Exit criterion: each converter improvement either lands in `OfficeIMO.Pdf` or
`OfficeIMO.Drawing` as reusable behavior first, or documents why it is genuinely
adapter-specific.

## Documentation Rule

Keep this file as the single PDF roadmap/state document. Avoid dated review
files under `Docs/reviews` for PDF state. If the current state changes, update
this file and the relevant package README instead.
