# OfficeIMO PDF Premium Roadmap Review - 2026-06-06

## Scope

Reviewed `origin/master` at `87dbe736` from the isolated worktree
`C:\Support\GitHub\OfficeIMO-pdf-premium-roadmap-20260606` on branch
`codex/pdf-premium-roadmap-20260606`.

This review focuses on the premium PDF product question: what the current
implementation already owns, what is still missing, and what the next focus
should be if OfficeIMO.Pdf is to become a serious first-party PDF library rather
than a collection of converter helpers.

## Current State

OfficeIMO now has a real first-party PDF platform:

- `OfficeIMO.Pdf` owns the dependency-light core for authoring, layout, PDF
  writing, logical reading, inspection, page operations, stamping, metadata
  editing, simple forms, e-invoice/PDF/A/PDF/UA groundwork, compliance
  readiness, and optional external proof modeling.
- `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf`,
  `OfficeIMO.Markdown.Pdf`, and `OfficeIMO.Html.Pdf` are thin adapters that map
  their source models into reusable PDF primitives.
- `OfficeIMO.Reader.Pdf` now registers PDF ingestion with `DocumentReader`,
  emitting page-aware chunks with Markdown text, detected tables, image
  placeholders, links, and form widget summaries where the logical PDF model can
  expose them.
- The converter warning story is moving in the right direction: the shared
  `PdfConversionReport` and `PdfConversionWarning` model exists in
  `OfficeIMO.Pdf`, and the existing converters expose a shared report while
  preserving source-specific warning lists.
- The Poppler raster baseline lane covers core authoring, Word, Excel,
  Markdown, and PowerPoint scenarios, including the newer dense PowerPoint
  baseline and visual review gallery.

The architecture direction is correct. The premium work should continue to
strengthen `OfficeIMO.Pdf` and `OfficeIMO.Drawing` as shared engines, with
adapters staying thin. Avoid adding parallel PDF layout engines inside
Word/Excel/PowerPoint/HTML/Reader packages.

## What We Have

### Authoring

The authoring engine is already beyond a toy PDF writer. It has flow content,
rich paragraphs, headings, lists, row/column layout, panels, tables with spans
and styling, images, vector drawing, foreground canvas, headers/footers, page
labels, viewer preferences, metadata, attachments, form fields, and report-like
visual fixtures.

### Conversion

The converter graph now covers high-value inbound paths:

- Word to PDF.
- Excel to PDF.
- PowerPoint to PDF.
- Markdown to PDF.
- HTML to PDF through semantic and Word-document profiles.
- PDF to semantic or positioned-review HTML.
- PDF to `DocumentReader` chunks through `OfficeIMO.Reader.Pdf`.

The package layering is sensible: document-specific parsing stays in the owning
OfficeIMO package, and PDF layout/writing stays in `OfficeIMO.Pdf`.

### Reading And Manipulation

The parser/rewrite layer supports practical operations such as text extraction,
logical readback, image and attachment extraction, page ranges, split/extract,
merge/import, delete/duplicate/move/reorder/rotate, metadata edits, stamps, and
simple form fill/flatten. It also blocks unsafe or unsupported inputs with
structured diagnostics instead of silently corrupting PDFs.

### Proof Discipline

This is one of the strongest parts of the current implementation. Compliance
profiles are not falsely enabled; `ComplianceProfile != None` remains blocked
until validator-backed generation exists. Readiness and proof models are
separate, and optional validator gates exist for veraPDF, PDF/UA validator, and
Mustang-style checks.

## What Is Missing

### 1. Premium Typography

The biggest cross-cutting gap is text. The engine still needs font subsetting,
OpenType/CFF support, Unicode shaping, glyph fallback, script-aware line
breaking, hyphenation hooks, and multilingual fixtures. This affects every
surface: authoring quality, Office conversion, PDF/UA, PDF/A, HTML rendering,
Reader extraction, and PowerShell wrappers.

Until typography is strong, OfficeIMO.Pdf can produce attractive reports in
controlled cases, but it cannot credibly claim premium document generation for
global business documents.

### 2. Validator-Backed Compliance

The compliance groundwork is broad, but the product is still readiness-only for
PDF/A, PDF/UA, Factur-X, and ZUGFeRD. Formal claims need validator evidence,
machine-readable proof artifacts, and CI lanes that map failures back to
actionable PDF layout/content diagnostics.

Do not turn on formal profile generation until readiness and external proof are
both green for that profile family.

### 3. Real-World Parser Preservation

The current rewrite posture is conservative, which is good. The missing premium
step is preserving more real-world structures safely: xref streams, object
streams, encryption/decryption where allowed, signatures inspection, tagged
PDFs, optional content, complex metadata/name trees/output intents, richer forms,
incremental updates, and safe redaction.

This should be staged as preservation capability, not as "parse anything and
rewrite something" optimism.

### 4. Converter Fidelity

The converters are product-valuable but still partial:

- Word: anchored/floating content, complex tables, SmartArt, equations without
  extractable text, tracked revisions, field evaluation, and more authored
  layout edge cases.
- Excel: fit-to-height, automatic pagination/scaling, locale-specific formats,
  richer chart fidelity, worksheet drawing placement, print-title/page-break
  edge cases, and more conditional formatting coverage.
- PowerPoint: masters/layout inheritance, theme resolution, grouped transforms,
  SmartArt/media fallbacks, richer table styles, richer text layout, and full
  slide fidelity.
- HTML: declared trusted/untrusted profiles, CSS subset documentation,
  paged-media behavior, resource policy diagnostics, and visual baselines for
  invoice/article/table-heavy/untrusted fixtures.
- Reader/PDF: richer coordinates, image placement metadata, table confidence,
  form metadata, source diagnostics, and a corpus of born-digital, scanned,
  tagged, form-heavy, invoice, statement, and report PDFs.

### 5. Productized Visual Proof

The Poppler baseline lane is a strong start, but premium positioning needs a
repeatable artifact pack, not only tests. The visual review gallery currently
points to generated PDFs in a prior local worktree. The next version should
produce a stable artifact folder from a script/target so reviewers, release
notes, and CI artifacts all inspect the same output.

Visual proof should include reports, invoices/statements, dashboards, forms,
slides, spreadsheets, technical docs, multilingual documents, compliance
fixtures, and hostile layout edge cases.

### 6. File Structure Pressure

Several files are still above or near the structure threshold:

- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Layout.Context.Process.Row.cs` -
  967 lines.
- `OfficeIMO.Pdf/Rendering/PdfWriter.cs` - 1011 lines.
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Text.cs` - 905 lines.
- `OfficeIMO.Pdf/Reading/Core/ResourceResolver.cs` - 911 lines.
- `OfficeIMO.PowerPoint.Pdf/PowerPointPdfConverterExtensions.cs` - 831 lines.
- `OfficeIMO.Markdown.Pdf/MarkdownPdfVisualTheme.cs` - 751 lines.
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Drawing.cs` - 762 lines.

Before adding large new features in those areas, split by responsibility: text
layout, row flow, resource resolution, drawing, PowerPoint text/table/shape
mapping, and Markdown theme presets.

## Recommended Next Focus

### Focus 1 - Typography And Text Layout

Make text the next premium milestone. This is the shared bottleneck behind
authoring, conversion, accessibility, compliance, and extraction quality.

Current branch progress: `PdfTextDiagnostics.AnalyzeWinAnsiText(...)` and
`AnalyzeWinAnsiTextRuns(...)` now provide a shared preflight contract for
unsupported generated text. `PdfTextEncodingDiagnostic` records source, index,
code point, display text, control-character state, stable warning code, and can
produce a shared `PdfConversionWarning` with layout diagnostic details. This
does not finish premium typography; it gives adapters and wrappers a safe
diagnostic surface while embedded Unicode writing, shaping, fallback, and
subsetting are built.

Deliverables:

- font subsetting and deterministic embedded-font output,
- OpenType/CFF loading path,
- Unicode text writing beyond WinAnsi,
- glyph fallback and missing-glyph diagnostics,
- shaping-ready abstraction even if full complex-script shaping lands in
  stages,
- multilingual visual baselines for Latin, Central/Eastern European, CJK, RTL,
  and symbol-heavy documents,
- shared text measurement/render diagnostics surfaced through
  `PdfConversionReport`.

Exit criteria: a multilingual report can be generated with embedded fonts,
extractable text, visual baselines, and no silent missing glyphs.

### Focus 2 - Visual Proof As A Release Artifact

Turn the visual review gallery into a repeatable release artifact.

Deliverables:

- a script or test target that emits the current PDF visual gallery into a
  deterministic `artifacts/pdf-visual-review` folder,
- an index that records commit, runtime, rasterizer, and scenario list,
- CI upload of PDFs/PNGs/diffs for raster failures,
- review fixtures for core reports, invoices/statements, dashboards, forms,
  PowerPoint decks, Excel workbooks, Markdown/HTML documents, and multilingual
  typography.

Exit criteria: every release candidate can attach a PDF gallery that a human can
open and inspect without reconstructing a previous worktree.

### Focus 3 - HTML/PDF And Reader/PDF Productization

The new bridge and reader adapter should move from "exists" to "safe product
lane."

Current branch progress: `HtmlPdfProfileContracts`,
`PdfHtmlProfileContracts`, and `ReaderPdfProfileContracts.OfficeIMO` now expose
stable product contracts for the HTML/PDF and Reader/PDF adapter lanes. The
contracts describe profile identifiers, first-party pipelines, intended use,
fidelity guarantees, safety behavior, and unsupported scope without adding a
second renderer/parser or wrapper-local policy layer.

Deliverables:

- documented semantic vs document HTML-to-PDF profiles,
- documented PDF-to-HTML semantic vs positioned-review profiles,
- CSS subset and unsupported-feature diagnostics,
- trusted/untrusted resource policies with examples,
- Reader.Pdf capability manifest examples and chunk metadata contract,
- a real-world PDF/HTML corpus with accepted degradation notes.

Exit criteria: users can choose the right HTML/PDF/Reader path and understand
where fidelity is guaranteed, simplified, or unsupported.

### Focus 4 - Compliance Proof

Keep compliance honest and make proof easier to consume.

Current branch progress: `PdfComplianceGateTests` now has an opt-in
`OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT` artifact hook, and
`Build/Export-PdfComplianceProof.ps1` turns the current PDF/A-3b, PDF/UA-1, and
e-invoice groundwork gates into a repeatable proof pack with generated PDFs,
validator diagnostics, expected-status metadata, profile proof matrix rows,
`officeimo-profile-proof-contract.json`, `index.md`, and machine-readable
`proof.json`.
`.github/workflows/pdf-compliance-proof.yml` now publishes that pack as a PR/CI
artifact, summarizes observed and expected validator statuses plus the profile
matrix and product proof contract in the job summary, and can be manually
dispatched with strict validator requirements plus optional veraPDF, PDF/UA
validator, and Mustang path and argument overrides.
This still does not claim formal conformance; it makes the current
validator-backed gap visible and release-reviewable.

Deliverables:

- machine-readable proof artifacts from validator gate tests,
- veraPDF lane for PDF/A,
- generic PDF/UA validator lane wired, with the concrete CI validator package
  still to be selected,
- Mustang lane for Factur-X/ZUGFeRD,
- examples that display conformance badges only after `AssessProof(...)`
  succeeds,
- diagnostics that map validator failures back to missing output intent, fonts,
  language, tagged structure, alt text, embedded XML, or XMP metadata.

Exit criteria: at least one narrow PDF/A profile can be claimed from generated
OfficeIMO.Pdf output with internal readiness plus passing external validation.

### Focus 5 - Safe Manipulation Expansion

After text/proof work, expand parser/rewrite support around preservation.

Deliverables:

- xref stream and object stream support,
- incremental-update preservation strategy,
- signature inspection and non-destructive handling,
- encryption read/decrypt where keys are supplied,
- safe redaction with proof that removed content is not recoverable,
- broader form appearance regeneration and flattening.

Exit criteria: common real-world business PDFs can be inspected and safely
rewritten, and unsupported documents still fail closed with actionable blockers.

## Immediate Low-Risk Hardening

`OfficeIMO.Reader.Pdf` is now part of the PDF product surface and currently has
only project references. Add it to the dependency-light package guardrail so the
adapter does not quietly acquire runtime package dependencies later.

The visual proof lane should also be repeatable from the repo root instead of
depending on a prior local worktree. `Build/Export-PdfVisualReviewGallery.ps1`
now wraps the existing Poppler raster-baseline scenario builders and writes the
generated PDFs plus an `index.md` manifest to `artifacts/pdf-visual-review` by
default.

## Roadmap Shape

The recommended sequencing is:

1. Typography and text diagnostics.
2. Repeatable visual proof gallery.
3. HTML/PDF and Reader/PDF product contracts.
4. Validator-backed compliance claim for one narrow profile.
5. Broader safe manipulation and preservation.
6. Deeper converter fidelity once shared text/proof infrastructure is solid.

This keeps the work reusable-core first: fix the shared engine and proof lanes,
then let Word, Excel, PowerPoint, Markdown, HTML, Reader, PSWriteOffice, and
future Visio/EPUB paths benefit without adapter-specific workarounds.

## Validation Performed

- Source inventory across `OfficeIMO.Pdf`, PDF converters, `OfficeIMO.Html.Pdf`,
  `OfficeIMO.Reader.Pdf`, tests, and current PDF docs.
- File-size structure scan for PDF-related C# files.
- Dependency guardrail updated to include `OfficeIMO.Reader.Pdf`.
- Repeatable visual review gallery export path added through
  `Build/Export-PdfVisualReviewGallery.ps1`.
- Shared text encoding preflight diagnostics added through
  `PdfTextDiagnostics` / `PdfTextEncodingDiagnostic`.
- Repeatable compliance proof export path added through
  `Build/Export-PdfComplianceProof.ps1`; local smoke run generated three
  groundwork PDFs, three not-run validator diagnostics,
  `officeimo-profile-proof-contract.json`, `index.md`, and `proof.json` with
  validator configuration, expected-status metadata, product proof contract
  rows, and profile-level proof rows.
- PDF compliance proof workflow added with proof-pack validation and artifact
  upload, including manual strict-run validator path/argument inputs for
  veraPDF, PDF/UA, and Mustang.
