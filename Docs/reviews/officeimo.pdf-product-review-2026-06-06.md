# OfficeIMO PDF Product Review - 2026-06-06

## Scope

Reviewed current `origin/master` from the isolated worktree
`C:\Support\GitHub\OfficeIMO-pdf-capability-review-20260606` on branch
`codex/pdf-capability-review-20260606`.

This review builds on the existing PDF support matrix and the 2026-06-05 PDF
implementation review. It focuses on the next product jump: making PDF a
premium OfficeIMO product, identifying concrete bugs and risks, and filling the
missing converter graph, especially HTML to PDF and PDF to HTML.

Implementation update in this branch: Phase 0 now introduces a shared
`PdfConversionReport`/`PdfConversionWarning` surface across the existing Office
PDF converters and backs PowerPoint text/table clipping warnings with actual
`OfficeIMO.Pdf` render-pass diagnostics. Phase 1 now starts with
`OfficeIMO.Html.Pdf`, a thin bidirectional adapter that exposes semantic and
document HTML-to-PDF profiles plus semantic and positioned-review PDF-to-HTML
profiles over existing OfficeIMO ingestion and PDF logical/read/render layers.

## Current State

OfficeIMO now has a real first-party PDF stack, not just a wrapper:

- `OfficeIMO.Pdf` is the shared dependency-light PDF engine for create, read,
  inspect, split, merge, page edit, stamp, metadata, forms groundwork, logical
  extraction, and compliance-readiness diagnostics.
- `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf`, and
  `OfficeIMO.Markdown.Pdf` convert their owning document models into
  `OfficeIMO.Pdf` primitives.
- `OfficeIMO.Html.Pdf` now provides the first named bidirectional HTML/PDF
  adapter surface, with semantic/document HTML-to-PDF profiles and
  semantic/positioned-review PDF-to-HTML profiles.
- `OfficeIMO.Word.Html` is bidirectional for Word and HTML.
- `OfficeIMO.Markdown.Html` and `OfficeIMO.Reader.Html` provide HTML ingestion
  into Markdown/Reader-shaped content.
- Poppler visual baselines already cover core PDF, Word, Excel, Markdown, and
  PowerPoint PDF scenarios.

The architecture direction is good: PDF layout and syntax stay in
`OfficeIMO.Pdf`; source-specific packages map their own models into that engine.
The next work should keep strengthening that shared core rather than adding
independent converter mini-engines.

## Bugs And Product Risks

### P0 - HTML/PDF bridge exists but needs fidelity maturity

The baseline had no `OfficeIMO.Html.Pdf`, `OfficeIMO.Pdf.Html`, or
`OfficeIMO.Reader.Pdf` adapter package. This branch adds `OfficeIMO.Html.Pdf` as
the single HTML/PDF bridge rather than splitting directionality across two
packages. It now covers:

- PDF to semantic HTML for search/indexing,
- PDF to layout-preserving HTML for visual review,
- HTML to PDF through semantic and document profiles.

Remaining gaps:

- PDF to Reader chunks with page/coordinate/image/table metadata.

The new bridge should now mature into a documented product lane:
trusted/untrusted HTML profiles, a declared CSS subset, paged-media behavior,
visual baselines, richer PDF image placement metadata, and clear diagnostics for
unsupported CSS/media or unsupported PDF structures.

### P1 - Converter diagnostics needed a shared report contract

Each converter has its own warning type:

- `OfficeIMO.Word.Pdf.PdfExportWarning`
- `OfficeIMO.Excel.Pdf.ExcelPdfExportWarning`
- `OfficeIMO.PowerPoint.Pdf.PowerPointPdfExportWarning`
- `OfficeIMO.Markdown.Pdf.MarkdownPdfExportWarning`

Only some warnings carry `PdfLayoutDiagnostic`. A premium conversion product
needs one shared diagnostic contract for skipped content, simplified content,
clipping, unsupported CSS/layout, missing images, unsafe resources, validator
status, and confidence level. Without that, wrappers and PowerShell cmdlets will
either lose useful details or need converter-specific branches.

This branch starts that contract with shared report/warning types in
`OfficeIMO.Pdf` and converter options that expose `ConversionReport` while
preserving existing warning lists.

### P1 - PowerPoint overflow diagnostics needed render-pass backing

The baseline `PowerPointPdfConverterExtensions.RenderTextBox` estimated
paragraph height before calling the shared PDF text-box renderer, then warned if
the estimated height exceeded the text area. Table overflow diagnostics
similarly used `EstimateDiagnosticTextWidth`, which multiplied character count
by a fixed font-size factor.

That is a useful first warning, but it is not the product contract. It can miss
real clipping or report false positives for fonts, rich runs, hyperlinks, long
tokens, line breaks, and theme-resolved typography. The shared PDF layout engine
should return actual text/table layout diagnostics from the render pass, and
PowerPoint should forward those.

This branch moves PowerPoint text-box and table-cell clipping warnings onto
actual canvas render-pass diagnostics from `OfficeIMO.Pdf`.

### P1 - Formal compliance is readiness-only

PDF/A, PDF/UA, Factur-X, and ZUGFeRD are correctly exposed as readiness and
proof concepts, not as enabled conformance profiles. The current validator
fixtures are expected to fail. This is honest, but the product risk is marketing
or wrappers accidentally claiming compliance because the metadata primitives
exist.

Keep `ComplianceProfile != None` blocked until internal readiness plus external
validator evidence is green in CI for each profile.

### P1 - Reading/manipulation still blocks many real-world PDFs

The parser and rewrite pipeline intentionally block encrypted PDFs, signed PDFs,
tagged PDFs, active content, complex name trees, complex metadata, complex
output intents, complex embedded files, optional content, and richer forms.
That is the right safety behavior, but it means the product is still far from a
general-purpose PDF manipulation engine.

The next read/manipulation milestones should focus on preserving or safely
round-tripping these structures, not only detecting them.

### P2 - Font and text shaping are not premium yet

The engine has standard fonts, ToUnicode groundwork, and TrueType embedding for
generated standard-font slots, but the support matrix still calls out planned
OpenType/CFF embedding, subsetting, Unicode shaping, and glyph fallback.

This affects every product lane: authoring quality, Office conversion fidelity,
PDF/UA/PDF/A readiness, extraction quality, and multilingual documents.

### P2 - Large files are slowing the next wave

Several files are already above the structure threshold:

- `OfficeIMO.PowerPoint.Pdf/PowerPointPdfConverterExtensions.cs` - 1330 lines
- `OfficeIMO.Pdf/Rendering/PdfWriter.cs` - 1157 lines
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Layout.Context.Process.Row.cs` - 1076 lines
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Text.cs` - 1035 lines
- `OfficeIMO.Pdf/Reading/Core/ResourceResolver.cs` - 940 lines

Before adding HTML/PDF conversion, richer layout, or parser preservation, split
by responsibility: source mapping, text layout, table layout, image handling,
resource resolution, diagnostics, and PDF syntax writing.

### P2 - Converter graph is uneven

Current strong paths:

- Word to PDF
- Excel to PDF
- PowerPoint to PDF
- Markdown to PDF
- Word to/from HTML
- HTML to Markdown/Reader
- HTML to PDF through `OfficeIMO.Html.Pdf`
- PDF to semantic/positioned-review HTML through `OfficeIMO.Html.Pdf`
- PDF to text/Markdown-like logical extraction

Missing or not productized:

- PDF to Reader package registration
- PDF to Word using the logical model
- Reader-anything to PDF through a shared document model
- Visio to PDF
- EPUB to PDF
- Excel/PowerPoint to HTML
- converter graph examples that show safe multi-hop flows and their expected
  fidelity loss

## Implementation Phases

### Phase 0 - Product Contract Hardening

Goal: make existing PDF work safer to wrap and easier to prove.

1. Introduce a shared `PdfConversionWarning` / `PdfConversionReport` contract in
   `OfficeIMO.Pdf` with source, severity, kind, page/slide/sheet, coordinates,
   message, and optional original converter-specific details.
2. Have Word/Excel/PowerPoint/Markdown PDF options expose the shared report while
   preserving existing warning lists as compatibility adapters if needed.
3. Move PowerPoint text/table overflow reporting to actual shared layout
   diagnostics returned by the PDF renderer.
4. Keep expanding visual baselines, especially dense PowerPoint, long tables,
   multilingual text, embedded fonts, tagged-output groundwork, and form
   flattening.
5. Split the largest PDF files before adding new behavior in those areas.

Exit criteria: converters produce a consistent report, no silent clipping in
known visual gates, and focused tests cover the shared report contract.

Branch status: items 1-3 are implemented with focused tests. Items 4-5 remain
ongoing product work.

### Phase 1 - First-Class HTML To PDF

Goal: add the missing high-value inbound converter without pretending to be a
full browser.

1. Add `OfficeIMO.Html.Pdf` as the public thin adapter. Internally it can offer
   semantic and document profiles that reuse the existing `OfficeIMO.Markdown.Html`
   and `OfficeIMO.Word.Html` ingestion layers before rendering through
   `OfficeIMO.Pdf`.
2. Support two input modes:
   - semantic HTML profile: HTML -> Markdown/Reader AST -> PDF,
   - document HTML profile: HTML -> WordDocument -> PDF.
3. Reuse existing untrusted/trusted HTML resource policies: node limits,
   stylesheet limits, image limits, URI allow-lists, and diagnostics.
4. Declare the CSS subset: headings, paragraphs, links, lists, tables, images,
   captions, block quotes, simple colors, borders, spacing, page breaks, and
   basic `@page` margins.
5. Add raster baselines for an invoice, article, table-heavy report, and
   untrusted HTML with blocked resources.

Exit criteria: `html.SaveAsPdf(...)` is documented, tested, warns clearly about
unsupported CSS/media, and produces good business documents.

Branch status: items 1-2 are implemented as the first thin adapter. Items 3-5
remain the next maturity step.

### Phase 2 - PDF Reader Integration And HTML Fidelity

Goal: make PDF extraction useful for humans and pipelines.

1. Add `OfficeIMO.Reader.Pdf` registration so `DocumentReader` can read PDFs
   into chunks with page number, coordinates, text blocks, headings, links,
   images, tables, and form fields where available.
2. Mature `OfficeIMO.Html.Pdf` PDF-to-HTML exports:
   - semantic HTML from the logical model,
   - positioned review HTML with page wrappers and absolutely positioned text,
     tables, links, form widgets, and image-placement hints when the reader
     exposes them.
3. Keep OCR as an optional adapter, not part of the dependency-light core.
4. Build a corpus of born-digital, scanned, tagged, form-heavy, invoice,
   statement, and report PDFs.

Exit criteria: PDF to HTML works for born-digital PDFs with useful structure,
and unsupported/unsafe files produce clear read diagnostics instead of garbage.

Branch status: semantic and positioned-review PDF-to-HTML exports are now in
`OfficeIMO.Html.Pdf`; image placeholders remain page-scoped until the logical
PDF model exposes image placement geometry.

### Phase 3 - Premium Authoring Engine

Goal: make `OfficeIMO.Pdf` excellent for direct report/document generation.

1. Strengthen layout primitives: sections, repeated running elements, columns,
   keep-together, keep-with-next, widows/orphans, page templates, floats,
   absolute layers, clipping, transforms, and reusable themes.
2. Finish table excellence: repeated headers, complex spans, auto-fit, fixed
   layout, nested block content, long-cell splitting, row grouping, table
   captions, and visual table style galleries.
3. Improve typography: font discovery, subsetting, Unicode shaping, fallback,
   hyphenation hooks, OpenType/CFF support, text measurement parity, and
   multilingual fixtures.
4. Add polished templates for invoices, statements, reports, dashboards,
   certificates, forms, and technical documents.

Exit criteria: direct OfficeIMO PDFs look polished by default, have stable
visual proof, and expose ergonomic APIs for common business-document patterns.

### Phase 4 - Office Converter Fidelity

Goal: make source-document conversion feel native rather than approximate.

1. Word: improve anchored/floating objects, equations, SmartArt fallbacks,
   tracked revisions, complex section breaks, footnote/endnote placement, and
   advanced table cases.
2. Excel: implement fit-to-height, automatic pagination/scaling, page-break
   previews, locale-aware formats, richer conditional formatting, charts,
   drawing placement, and workbook print settings.
3. PowerPoint: improve theme color/font resolution, master/layout inheritance,
   group transforms, text layout, SmartArt/media fallbacks, table themes, and
   chart fidelity.
4. Markdown/HTML: improve paginated panels, code blocks, callouts, CSS-like
   spacing, page-break policy, repeated table headers, and image/resource
   policies.
5. Visio: add Visio to PDF through the existing first-party drawing/SVG/raster
   path instead of requiring desktop Visio.

Exit criteria: every converter has representative visual baselines, shared
diagnostics, and a documented support matrix with fidelity claims tied to tests.

### Phase 5 - Robust PDF Reading And Manipulation

Goal: move from useful parser to dependable PDF tooling.

1. Expand parser coverage: xref streams, object streams, filters, CMaps, font
   encodings, optional content, tagged structure, outlines, name trees, embedded
   files, output intents, and annotations.
2. Preserve complex catalog/page structures through safe rewrite operations.
3. Add encrypted-read support where passwords are supplied.
4. Add signature inspection and signature-preserving blocked/allowed rewrite
   policies.
5. Add safe redaction with content removal, annotation removal, and raster proof
   gates.
6. Improve form filling/flattening for richer widgets and appearance streams.

Exit criteria: real-world PDF operations either preserve the document safely or
fail with precise blockers and repair guidance.

### Phase 6 - Compliance, Packaging, And Public Proof

Goal: make formal claims only when independently proven.

1. Add CI lanes for veraPDF, a chosen PDF/UA validator, and Mustang.
2. Emit machine-readable proof artifacts from compliance tests.
3. Only enable formal profile generation after generated fixtures validate.
4. Add performance and memory benchmarks for large reports, large Excel exports,
   page extraction, merging, and PDF text extraction.
5. Add PowerShell/PSWriteOffice wrappers that stay thin and expose the shared
   conversion report.
6. Publish website docs and examples from the exact supported matrix.

Exit criteria: product claims match validator evidence, package contents, docs,
examples, and wrapper behavior.

## Recommended Immediate PRs

1. Shared conversion report contract and adapter mapping for existing warnings.
2. PowerPoint actual-layout diagnostics for text boxes and tables.
3. `OfficeIMO.Html.Pdf` alpha as the bidirectional bridge: HTML-to-PDF over
   existing Markdown/Word ingestion and PDF-to-HTML over `PdfLogicalDocument`.
4. `OfficeIMO.Reader.Pdf` alpha around `PdfLogicalDocument`.
5. Large-file splits for PowerPoint PDF mapping and core PDF writer text/layout.
6. Visual baseline expansion for HTML-to-PDF, PDF-to-HTML review output,
   multilingual font fallback, and dense slides.

## Non-Goals To Keep Explicit

- Browser-grade HTML/CSS rendering in the dependency-light core.
- OCR in the core package.
- Compliance claims without validator evidence.
- Rewrite of encrypted/signed/tagged/active PDFs without explicit preservation
  policy.
- Large converter-specific layout engines outside `OfficeIMO.Pdf`.
