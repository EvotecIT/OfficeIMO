# OfficeIMO RTF Market Assessment - 2026-06-18

Branch: `codex/rtf-market-assessment-20260618`

Base: `origin/master` at `0c5b047d` (`Harden PowerPoint package XML loading (#1974)`)

## Current State

OfficeIMO already has a real first-party RTF stack, not just examples:

- `OfficeIMO.Rtf` is the reusable engine for dependency-free parsing, syntax preservation, semantic binding, fluent document construction, lossless editing, and deterministic writing.
- `OfficeIMO.Word.Rtf` maps between `WordDocument` and the reusable `RtfDocument` model.
- `OfficeIMO.Html` contains bidirectional HTML/RTF conversion on top of `OfficeIMO.Rtf`.
- `OfficeIMO.Rtf.Markdown` maps directly between `RtfDocument` and `MarkdownDoc` without detouring through Word, HTML, or PDF.
- `OfficeIMO.Rtf.Pdf` contains bidirectional RTF/PDF conversion through the first-party PDF model.
- `OfficeIMO.Reader.Rtf` registers an RTF adapter for `OfficeIMO.Reader`.

The RTF-related surface currently spans 329 implementation files and 66 RTF test files. The focused RTF test slice passes across all currently relevant test frameworks:

```powershell
dotnet test OfficeIMO.Rtf.Tests\OfficeIMO.Rtf.Tests.csproj -c Release --filter "FullyQualifiedName~Rtf"
```

Result:

- `net10.0`: 551 passed, 0 failed.
- `net8.0`: 551 passed, 0 failed.
- `net472`: 551 passed, 0 failed.

## Current Conversion Graph

Direct first-party routes:

- RTF read/write: `RtfDocument.Read`, `RtfDocument.Load`, `RtfDocument.ToRtf`, `RtfReadResult.ToRtfLossless`, `RtfLosslessEditor`.
- RTF <=> Word: `WordDocument.ToRtfDocument`, `WordDocument.ToRtf`, `RtfDocument.ToWordDocument`, `LoadFromRtf`.
- RTF <=> HTML: `RtfDocument.ToHtml`, `html.ToRtfDocument`, `html.ToRtf`, file/stream/async variants.
- RTF <=> Markdown: `RtfDocument.ToMarkdownDocument`, `RtfDocument.ToMarkdown`, Markdown string to `RtfDocument`, Markdown string to RTF.
- RTF <=> PDF: `RtfDocument.ToPdfDocument`, RTF string/bytes/stream/file to PDF, PDF read/logical documents to RTF, file/stream/async variants.
- RTF => Reader chunks: `DocumentReaderRtfRegistrationExtensions.RegisterRtfHandler`, `ReadRtfFile`, `ReadRtf`, `ReadRtfDocument`.

Indirect routes that work only through another model today:

- RTF <=> DOC/DOCM/ODT is not a first-class OfficeIMO story in this stack.

## Market Bar

Commercial competitors sell broad format conversion as the product promise: Aspose.Words, GemBox.Document, Syncfusion DocIO, Telerik WordsProcessing, TX Text Control, and SautinSoft all position around document load/edit/save/conversion across DOCX/RTF/HTML/PDF and related formats. Open-source RTF options such as RtfPipe focus mostly on RTF to HTML/Markdown-style extraction, not a full first-party document engine.

To be best on the market, OfficeIMO should not only match "loads RTF and saves X". It should differentiate on:

- dependency-free, cross-platform, deterministic conversion;
- lossless syntax preservation and targeted lossless editing;
- diagnostics and declared degradation instead of silent content loss;
- composable semantic model shared by Word, HTML, Markdown, PDF, and Reader;
- public compatibility corpus and round-trip scorecards.

## Gap Assessment

Strong foundations:

- Dedicated engine package with parser, syntax tree, semantic model, writer, byte-preserving/lossless paths, and diagnostics.
- Rich semantic model coverage for fonts, colors, styles, paragraphs, runs, sections, notes, page setup, headers/footers, tables, images, objects, shapes, fields, revisions, document variables, user properties, and code pages.
- Clean adapter split: Word, HTML, PDF, and Reader do not own RTF parsing.
- Good multi-target coverage including `netstandard2.0`, `net8.0`, `net10.0`, and Windows-only `net472`.
- Focused RTF suite passes locally.

Material gaps before "best on market":

- Markdown bridge is now first-class, but still needs more edge-case fixtures for footnotes, metadata, raw HTML, and media embedding callbacks.
- No formal feature matrix that maps RTF controls/destinations to Word, HTML, Markdown, PDF, and Reader outputs.
- No benchmark/compatibility corpus against WordPad/Word, LibreOffice, Google Docs import/export, and competitor libraries.
- PDF import is correctly described as semantic extraction, not visual reconstruction; this needs an explicit product positioning and test corpus.
- HTML and Word conversions need public "loss budget" diagnostics for each degraded feature. Phase 3 has started this by making RTF-to-HTML skipped images diagnostic and by preserving leading Word footnote references instead of dropping them.
- Reader/AI extraction is good, but should become a first-class RTF ingestion story with chunk provenance, tables, images, notes, and hidden/revision handling documented.
- RTF security posture should be packaged as a promise: bounded input, depth limits, binary payload limits, object/file destination policy, URL/image policy, safe defaults for untrusted content.

## Proposed Plan

### Phase 1 - Product Contract And Scorecard

- Add `Docs/officeimo.rtf-support-matrix.md` as the source-of-truth matrix for RTF syntax, semantic model, Word, HTML, Markdown, PDF, Reader, lossless edit, and diagnostics.
- Define conversion classes: lossless, semantic-preserving, visual-preserving, extractive, and unsupported-with-diagnostic.
- Add golden corpus buckets: basic text, Unicode/code pages, styles, paragraph layout, sections/page setup, lists, tables, images, fields, notes, headers/footers, revisions, shapes/objects, metadata, pathological input.
- Publish a current "RTF conversion scorecard" from tests rather than marketing prose.

### Phase 2 - RTF <=> Markdown First-Class Bridge

- Add `OfficeIMO.Rtf.Markdown`, depending on `OfficeIMO.Rtf` and `OfficeIMO.Markdown`.
- Map RTF semantic blocks to Markdown AST/output directly: paragraphs, headings/styles, emphasis, links, lists, tables, code-like styles, images, and warnings.
- Add Markdown to RTF through the Markdown AST instead of detouring through HTML.
- Preserve unsupported rich constructs as placeholders or diagnostics depending on options.

### Phase 3 - Word/HTML Fidelity Hardening

- Preserve leading Word footnote/endnote references by mapping them to generated RTF note reference markers when there is no preceding text run to attach to.
- Report skipped RTF-to-HTML images through diagnostics when embedding is disabled, image data is missing, or the image format cannot be emitted as HTML.
- Continue expanding Word bridge coverage around styles, numbering, tracked changes, comments, section inheritance, headers/footers, floating objects, shape text, and table edge cases.
- Continue expanding HTML bridge around CSS coverage, images/resources, lists, table layout, directionality, embedded metadata, and accessible output.
- Keep PDF implementation out of this phase to avoid colliding with active PDF work in another branch/worktree.

### Phase 4 - Lossless Editing As The Differentiator

- Extend `RtfLosslessEditor` from targeted text/metadata/settings edits toward safe structural edits: insert/remove paragraphs, update styles, rewrite fields, replace images, add/remove headers and footers, and modify table text without normalizing untouched syntax.
- Add "preserve unknown destinations" tests for every editor operation.
- Add a diff-friendly mode that emits stable RTF while preserving opaque binary/object payloads.

### Phase 5 - Safety, Performance, And Scale

- Add explicit untrusted input profiles for RTF read, HTML-to-RTF, RTF-to-PDF, and Reader ingestion.
- Enforce configurable limits for group depth, token count, binary payload bytes, image count/bytes, object destinations, external resource loading, output size, and conversion time.
- Add streaming/large-document benchmarks and memory ceilings.
- Add fuzzing/property tests for parser recovery, lossless round trip, and malformed input diagnostics.

### Phase 6 - PDF Integration Last

- Integrate with the active PDF work only after that workstream has settled.
- Make PDF export visibly better for pagination, page boxes, header/footer repetition, notes, tables, images, and link/bookmark mapping.
- Keep PDF import positioned as structured extraction; improve logical grouping and provenance rather than claiming lossless reconstruction.

### Phase 7 - Market Proof And Developer Experience

- Add CLI/sample workflows for:
  - `rtf -> docx`
  - `docx -> rtf`
  - `rtf -> html`
  - `html -> rtf`
  - `rtf -> markdown`
  - `markdown -> rtf`
  - `rtf -> pdf`
  - `pdf -> rtf` semantic extraction
- Add NuGet readmes with one-screen examples and explicit capability tables.
- Add visual/golden artifact tests for real-world RTF samples generated by Microsoft Word, LibreOffice, Google Docs export, and common EHR/CRM/helpdesk systems.
- Publish comparison pages that emphasize OfficeIMO advantages: MIT, dependency-free, deterministic, lossless edit path, diagnostics, cross-platform targets, and first-party OfficeIMO integration.

## Recommended Ownership

Keep the current architecture:

- `OfficeIMO.Rtf` owns parser, syntax tree, lossless editor, semantic model, writer, diagnostics, and shared feature contracts.
- `OfficeIMO.Word.Rtf`, `OfficeIMO.Html`, `OfficeIMO.Rtf.Pdf`, and `OfficeIMO.Rtf.Markdown` own only adapter mappings.
- `OfficeIMO.Reader.Rtf` owns extraction/chunking only.
- Docs, examples, CLI wrappers, and PSWriteOffice surfaces should stay thin and call the engine/adapters.
