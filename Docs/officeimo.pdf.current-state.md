# OfficeIMO PDF Current State

This is the canonical PDF product state file. Keep it current and delete dated
review snapshots instead of adding another one.

## Product Direction

`OfficeIMO.Pdf` is the first-party PDF engine for OfficeIMO. The goal is a
dependency-light, MIT-licensed PDF platform that can create, inspect, read,
convert, and safely manipulate business PDFs without Microsoft Office,
commercial PDF engines, or runtime rasterizer dependencies.

The reusable PDF engine should stay in `OfficeIMO.Pdf`. Source packages such as
Word, Excel, PowerPoint, Markdown, HTML, and Reader should remain thin adapters
that map their own models into shared PDF primitives or consume the shared PDF
read model.

## Current Areas

### Core PDF Creation

Status: useful and broad, still not premium typography.

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
- Table styling, Word-like table style names, row splitting, repeated headers,
  spans, rich cell text, links, images, fills, borders, alignment, sizing, and
  visual table fixtures.
- Shared conversion warnings and text diagnostics for unsupported WinAnsi text
  before rendering.

Important gaps:

- Full Unicode writing beyond the current generated-font path.
- Font subsetting.
- OpenType/CFF support.
- Text shaping, ligatures, glyph fallback, complex script handling, and
  multilingual line breaking.
- Hyphenation hooks and stronger text measurement parity.

### PDF Reading And Inspection

Status: practical for born-digital/simple PDFs, conservative for complex PDFs.

Available now:

- Probe, inspect, preflight, text extraction, structured/logical readback,
  image extraction, attachment extraction, page metadata, outline/navigation
  readback, link annotations, form widget summaries, and diagnostics.
- Read and rewrite blockers for unsupported or risky inputs.
- Capability flags for wrapper dispatch, including text extraction, logical
  objects, images, page manipulation, simple form fill, and simple flattening.

Important gaps:

- Broader xref stream and object stream coverage.
- Encryption/decryption with supplied credentials.
- Signature inspection and signature-preserving policy.
- Tagged PDF preservation.
- Optional content/layers preservation.
- Complex metadata, name trees, output intents, embedded files, and richer form
  structures.
- OCR, which should remain outside the dependency-light core.

### PDF Manipulation

Status: useful for safe simple documents.

Available now:

- Split, page range extraction, merge/import, delete, duplicate, move, reorder,
  rotate, metadata editing, text/image stamp, text/image watermark, simple form
  field fill, and simple text/choice/button-widget flattening.
- Stream, path, and byte helper coverage with path validation and fail-closed
  preflight behavior.

Important gaps:

- Incremental update strategy.
- Safe redaction with proof that removed content is not recoverable.
- Rich form appearance regeneration.
- More complex page/resource preservation.
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
- `OfficeIMO.Markdown.Pdf`: Markdown-to-PDF path for headings, outlines, rich
  inline text, links, lists, task lists, tables, code/semantic panels, callouts,
  front matter, images, themes, and warnings.
- `OfficeIMO.Html.Pdf`: semantic/document HTML-to-PDF profiles and
  semantic/positioned-review PDF-to-HTML profiles over first-party pipelines.
- `OfficeIMO.Reader.Pdf`: PDF ingestion registration for `DocumentReader`
  chunks with page-aware locations, Markdown text, detected tables, image
  placeholders, links, form summaries, hashes, and split warnings where the PDF
  read model can expose them.

Important gaps:

- Word: anchored/floating layout, richer table fidelity, field evaluation,
  tracked revisions, SmartArt, hard equations, and authored edge cases.
- Excel: fit-to-height, automatic pagination/scaling, richer chart fidelity,
  more conditional formats, locale-specific formats, richer drawing placement,
  and print-layout edge cases.
- PowerPoint: master/layout inheritance, theme resolution, grouped transforms,
  richer text layout, richer table styles, media, and SmartArt fallbacks.
- Markdown/HTML: stronger paginated panels, declared CSS subset maturity,
  resource-policy examples, and more visual fixture families.
- Reader/PDF: richer coordinates, image placement metadata, table confidence,
  form metadata, source diagnostics, and a real-world PDF corpus.

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

1. Typography and text layout: subsetting, OpenType/CFF, Unicode writing,
   shaping, fallback, multilingual fixtures, and extraction-safe embedded fonts.
2. Validator-backed compliance: convert readiness into proof for one narrow
   profile before enabling any formal conformance switch.
3. Visual proof productization: make review galleries and proof packs standard
   CI/release artifacts.
4. Real-world parser preservation: safely preserve more structures and expand
   manipulation only where rewrite proof exists.
5. Converter fidelity: deepen Word, Excel, PowerPoint, Markdown, HTML, and
   Reader paths after shared typography/proof foundations improve.

## Next Focus

### 1. Typography Milestone

Build the shared text foundation first:

- deterministic embedded-font output,
- font subsetting,
- Unicode writing beyond WinAnsi,
- glyph fallback and missing-glyph diagnostics,
- shaping-ready abstraction,
- multilingual visual baselines,
- shared text diagnostics surfaced through `PdfConversionReport`.

Exit criterion: a multilingual business report can be generated with embedded
fonts, extractable text, visual proof, and no silent missing glyphs.

### 2. Proof Artifacts As Product Evidence

Make evidence easy to inspect:

- keep PDF compliance proof pack machine-readable,
- add visual review gallery upload in CI,
- record commit, runtime, rasterizer/validator configuration, scenario list,
  and expected vs observed status,
- keep examples and docs tied to proof results.

Exit criterion: every release candidate can attach current PDFs/proof metadata
that a reviewer can inspect without reconstructing a local worktree.

### 3. HTML/PDF And Reader/PDF Productization

Make the adapter lanes clear and safe:

- document semantic vs document HTML-to-PDF profiles,
- document semantic vs positioned-review PDF-to-HTML profiles,
- declare the supported CSS/resource subset,
- add trusted/untrusted examples,
- publish Reader.Pdf chunk metadata expectations,
- build a small real-world PDF/HTML corpus with accepted degradation notes.

Exit criterion: users can choose the right HTML/PDF/Reader path and understand
where fidelity is guaranteed, simplified, or unsupported.

### 4. One Narrow Compliance Claim

After typography/proof improves, choose one narrow profile and make it pass:

- start with the smallest PDF/A profile that can be generated honestly,
- wire required validator evidence,
- map failures back to actionable requirements,
- flip formal profile generation only when proof is green.

Exit criterion: one generated OfficeIMO.Pdf profile can be claimed from internal
readiness plus passing external validation.

## Documentation Rule

Keep this file as the single PDF roadmap/state document. Avoid dated review
files under `Docs/reviews` for PDF state. If the current state changes, update
this file and the relevant package README instead.
