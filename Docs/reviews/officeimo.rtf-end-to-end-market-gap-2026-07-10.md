# OfficeIMO RTF end-to-end and market gap audit - 2026-07-10

Branch: `codex/rtf-end-to-end-market-gap-20260710`

Base: `origin/master` at `28384ffa8` (`bump`)

## Verdict

OfficeIMO has a serious first-party RTF engine. It is already broader than the usual open-source RTF converter: it can parse and preserve source syntax, bind a rich semantic document model, write deterministic RTF, apply targeted lossless edits, and convert through first-party Word, HTML, Markdown, PDF, and Reader adapters.

It is not yet a proper end-to-end market implementation.

The main blocker is not the number of modeled formatting properties. The missing product contract is safe ingestion, explicit degradation, real interoperability proof, and complete workflows over the model. The current suite proves many focused contracts, but nine small synthetic corpus files cannot prove that files from Word, LibreOffice, Outlook, Google Docs, TextEdit, or competing libraries survive realistic round trips.

OfficeIMO can credibly lead the open-source RTF category after closing the P0 safety and diagnostic gaps. Competing with commercial document suites requires additional workflow and rendering depth, but OfficeIMO should not copy their WYSIWYG editor or printing products into `OfficeIMO.Rtf`.

## Current evidence

The RTF implementation currently contains:

| Surface | C# source files | Role |
| --- | ---: | --- |
| `OfficeIMO.Rtf` | 208 | Parser, syntax tree, semantic model, writer, diagnostics, lossless editor. |
| `OfficeIMO.Word.Rtf` | 16 | Word/RTF bridge. |
| `OfficeIMO.Html/Rtf` | 86 | HTML/RTF bridge. |
| `OfficeIMO.Rtf.Markdown` | 8 | Markdown/RTF bridge. |
| `OfficeIMO.Rtf.Pdf` | 13 | PDF/RTF bridge. |
| `OfficeIMO.Reader.Rtf` | 6 | Reader ingestion adapter. |
| `OfficeIMO.Rtf.Tests/Rtf` | 68 | Focused RTF tests. |

All five RTF NuGet surfaces are published: `OfficeIMO.Rtf` 0.1.10, `OfficeIMO.Word.Rtf` 0.1.10, `OfficeIMO.Rtf.Markdown` 0.1.8, `OfficeIMO.Rtf.Pdf` 0.1.10, and `OfficeIMO.Reader.Rtf` 0.0.11. Published versions were checked through the corresponding `api.nuget.org/v3-flatcontainer/<package>/index.json` endpoints.

Current validation:

```powershell
dotnet test OfficeIMO.Rtf.Tests\OfficeIMO.Rtf.Tests.csproj -c Release --logger "console;verbosity=minimal"
```

| Target | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| `net10.0` | 547 | 0 | 0 |
| `net8.0` | 547 | 0 | 0 |
| `net472` | 547 | 0 | 0 |

The suite has 490 test methods before theory expansion. The focused tests are valuable, but the corpus has only nine `.rtf` fixtures plus its README, and `RtfGoldenCorpusTests` currently verifies only that those fixtures parse without errors.

## What end to end should mean

An RTF implementation is end to end when a caller can safely ingest an arbitrary RTF, understand what was preserved or degraded, edit it, run document workflows, convert it, save it, and prove that the output reopens correctly in the applications users actually have.

| Stage | Current state | What is still missing |
| --- | --- | --- |
| Install and target .NET | Strong | Packages are published and target `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. |
| Read string, bytes, file, and stream | Functional but unbounded | Only group depth is bounded in the core. Input, token, text, binary, image, object, and output sizes are not. |
| Parse and preserve source syntax | Strong | Unknown content is preserved by the syntax/lossless path. There is no formal control-word coverage catalog. |
| Bind a semantic model | Broad but partial | Several RTF 1.9.1 families are preserve-only or unrecognized semantically, and skipped ignorable destinations are not reliably diagnosed. |
| Create and write RTF | Strong for the modeled surface | No compatibility profiles for target producers/readers and no externally verified writer corpus. |
| Targeted lossless editing | Useful but partial | Metadata, settings, tables, fonts, colors, page setup, text replacement, and append are covered; structural paragraph/table/image/header/footer edits are not. |
| General semantic editing | Partial | No document clone/merge, block insert/remove/move, bookmark range operations, or cross-run rich replacement in the RTF model. |
| RTF to/from Word | Broad but incomplete | Styles and list definitions are not carried as structures; shapes/objects and unsupported Word elements can disappear; no conversion report exists. |
| RTF to/from HTML | Broad but unsafe for untrusted output | RTF hyperlinks are emitted without an output URL policy, and round-trip object metadata can carry binary data into HTML attributes. |
| RTF to/from Markdown | Good core bridge | RTF notes and headers/footers are omitted with diagnostics; images need a complete extraction/embedding callback workflow. |
| RTF to/from PDF | Useful semantic export/import | It is not Word-grade pagination. Shapes/objects fall back to plain text without a warning, notes are appended rather than laid out as page notes, and only PNG/JPEG embed directly. |
| Reader/AI ingestion | Strong extraction surface | Safety depends partly on Reader input limits rather than a bounded core RTF profile. Provenance and diagnostics should be demonstrated on real producer files. |
| Mail merge, fields, compare, merge | Available through `OfficeIMO.Word` | No documented or result-bearing RTF workflow wraps the Word engine, so conversion loss is not visible to the caller. |
| Performance and resilience | Unproven | No RTF benchmark project, allocation budget, fuzzing lane, or large-document ceiling was found. |
| Interoperability proof | Weak | No checked-in Word, LibreOffice, Outlook, Google Docs, macOS TextEdit, or competitor-generated compatibility matrix. |

## P0 gaps: fix before calling untrusted RTF safe

### 1. The core parser is not resource-bounded

`RtfReadOptions` exposes only `MaxDepth`. `RtfDocument.Load` reads the complete file or stream into a string, `RtfTokenizer` materializes a token list, and binary payloads are copied into arrays. Direct Word, HTML, Markdown, and PDF routes can therefore accept inputs that consume unbounded memory before an adapter can apply policy.

`OfficeIMO.Reader.Rtf` does enforce `ReaderOptions.MaxInputBytes`, but that protects only the Reader surface. It does not make `OfficeIMO.Rtf`, `OfficeIMO.Word.Rtf`, or `OfficeIMO.Rtf.Pdf` safe ingestion boundaries.

Required core contract:

- `RtfReadOptions.CreateUntrustedProfile()` with conservative defaults.
- Maximum input bytes/chars, token count, text chars, group count, per-payload and total binary bytes, image count/bytes, object count/bytes, and semantic block count.
- A stable limit exception and diagnostic code for every limit.
- Cancellation checks during tokenization, tree construction, semantic binding, and conversion loops.
- Explicit object, file-reference, and external-link policies. Parsing should never fetch external resources.
- Tests that prove limits are enforced before large allocations where practical.

Streaming should follow measurements. A bounded tokenizer is the first contract; a streaming parser should not be built speculatively.

### 2. RTF-to-HTML needs a safe output profile

`HtmlToRtfOptions` already has `HtmlUrlPolicy` and an untrusted HTML profile. `RtfToHtmlOptions` has no equivalent. `RtfHtmlWriter` writes `run.Hyperlink` and field hyperlink targets directly into `href` after attribute encoding. Encoding prevents attribute injection, but it does not reject executable schemes such as `javascript:`.

The writer also serializes RTF object data and other OfficeIMO round-trip metadata into `data-officeimo-*` attributes. That is useful for a trusted fidelity round trip, but it is the wrong default for rendering untrusted RTF on a web page and can greatly inflate output.

Required HTML contract:

- Reuse `HtmlUrlPolicy` on RTF-to-HTML output.
- Make the default profile web-safe while the packages are still 0.x, or require callers to choose between `CreateWebSafeProfile()` and `CreateRoundTripProfile()`.
- Reject or neutralize unsafe run and field links with diagnostics.
- Add explicit policies for object data, file references, custom metadata, and embedded image data.
- Add image extraction/resource callbacks so disabling data URIs does not mean silently losing the image.
- Cover `javascript:`, `file:`, UNC, oversized data URI, object payload, and hostile field-instruction cases.

### 3. Semantic loss is not consistently reported

`RtfReadOptions.WarnOnUnsupportedDestinations` sounds broader than it is. The semantic reader skips ignorable destinations, but `RtfDestinationRegistry.IsUnsupportedSemanticDestination` currently identifies only registered object sub-destinations other than `object`. An arbitrary unknown ignorable destination can therefore be preserved in the syntax tree, omitted from the semantic model, and produce no warning.

Diagnostics are also fragmented:

- Core parsing uses `RtfDiagnostic`.
- HTML uses `HtmlRtfConversionDiagnostic`.
- Markdown uses `RtfMarkdownConversionDiagnostic`.
- PDF uses `PdfConversionReport`.
- Reader maps diagnostics into strings and chunk diagnostics.
- Word has no conversion report.

Required product contract:

- A shared `RtfConversionReport` in `OfficeIMO.Rtf`, because every RTF adapter already depends on the core.
- Stable severity, code, source path, feature/control word, action (`preserved`, `flattened`, `omitted`, `blocked`), and optional count/details.
- Result-returning APIs for Word conversion and equivalent report access from HTML, Markdown, PDF, and Reader without removing their useful native reports.
- A strict mode such as `RequireNoLoss()` for workflows that cannot accept silent degradation.
- Diagnostics for every ignored `IRtfBlock`, `IRtfInline`, Word element, destination family, unsafe resource, and fallback conversion.

## P1 gaps: interoperability and workflow completeness

### 4. The compatibility corpus is not market proof

The current corpus covers useful categories, but the actual files are small synthetic examples. There is no producer manifest, application/version metadata, visual baseline, or reopen result.

The corpus should include legally redistributable files created by:

- current Microsoft Word and at least one older Word generation;
- LibreOffice Writer;
- Outlook HTML-encapsulated RTF;
- Google Docs export/import;
- macOS TextEdit RTF/RTFD where applicable;
- common EHR/CRM/helpdesk exports;
- representative output from commercial libraries when licensing permits redistribution;
- malformed and adversarial generators.

Each fixture needs a manifest containing origin, version, license/provenance, expected syntax preservation, expected semantic features, expected diagnostics, and adapters under test. The harness should test lossless bytes, semantic snapshots, normalized rewrite, conversion diagnostics, target reopen, and visual output where the target is visual.

### 5. Important RTF 1.9.1 families are missing semantically

The [Microsoft RTF 1.9.1 reference](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxrtfcp/85c0b884-a960-4d1a-874e-53eeee527ca6) includes more than the current semantic model. Source search found no modeled controls for these high-value groups:

- Outlook/Exchange HTML encapsulation: `fromhtml`, `htmltag`, `htmlrtf`, and `mhtmltag`.
- East Asian/DBCS code pages such as 932, 936, 949, and 950, composite fonts, and related run controls.
- Nested table controls such as `nesttableprops`, `nestrow`, `nestcell`, and `itap`.
- Theme data, color scheme mapping, quick/latent styles, and table styles.
- Custom XML, smart tags, and data-store destinations.
- Move revisions and protection exception ranges.
- Index and table-of-contents entry destinations.

These do not all need editable object models immediately. The minimum honest implementation is: recognize the family, preserve it losslessly, classify it in the support catalog, and emit a semantic-degradation diagnostic. Outlook HTML encapsulation, DBCS text, and nested tables should be implemented before the lower-demand metadata families.

The open-source [RtfPipe](https://github.com/erdomke/RtfPipe) explicitly advertises HTML encapsulation and nested table support, so those are not merely commercial-suite expectations.

### 6. The Word bridge is the largest functional adapter gap

The bridge covers paragraphs, runs, rich character/paragraph formatting, tables, images, notes, comments, fields, revisions, headers/footers, sections, page setup, metadata, and async I/O. That is a strong base.

However:

- RTF stylesheet entries are not projected as Word styles.
- RTF list definitions and overrides are not projected as Word numbering structures.
- Top-level and inline RTF shapes/objects are not handled by the Word bridge switch statements.
- Word elements outside paragraphs/tables and recognized image paragraphs can be ignored.
- There is no diagnostic result to tell the caller what disappeared or flattened.

This also blocks the clean answer to commercial workflow parity. `OfficeIMO.Word` already owns mail merge, template preflight, find/replace, field update reports, document comparison, and document append/merge. `OfficeIMO.Rtf` must not duplicate those engines. `OfficeIMO.Word.Rtf` should provide thin result-bearing workflows:

1. Load RTF with a bounded profile.
2. Convert to `WordDocument` with a conversion report.
3. Call the existing Word operation.
4. Convert back to RTF with a conversion report.
5. Return the output plus combined diagnostics.

### 7. PDF export is implemented, but the old matrix still says deferred

`OfficeIMO.Rtf.Pdf` is now a published bidirectional package. It supports a meaningful semantic export surface including page setup, sections, text, lists, tables and merges, images, bookmarks, links, headers/footers, metadata, and notes. PDF import is correctly positioned as semantic extraction.

The remaining fidelity gaps are:

- Objects and shapes are reduced to `ToPlainText()` without a conversion warning.
- Notes are appended after the document instead of placed on their source pages.
- Section/header/footer variants and page-border behavior are simplified.
- EMF, WMF, and DIB images are diagnosed as unsupported instead of routing reusable conversion through `OfficeIMO.Drawing`.
- Font substitution/embedding and pagination decisions are not exposed as a complete RTF conversion report.
- There is no public RTF-to-page-image or thumbnail workflow, even though commercial suites commonly expose it.

Image decoding/conversion belongs in `OfficeIMO.Drawing`; pagination and PDF diagnostics belong in `OfficeIMO.Pdf`; the RTF adapter should remain mapping only.

### 8. HTML and Markdown need complete media and profile workflows

Markdown conversion has good diagnostics and strong list/table coverage. Markdown-to-RTF supports footnotes, but RTF-to-Markdown still reports run-attached notes as omitted instead of emitting Markdown footnote definitions. Headers/footers are omitted, and image conversion can create paths but does not provide a complete caller callback for consuming the image bytes.

HTML has a broad round-trip representation, including OfficeIMO metadata attributes. It needs two explicit products rather than one mixed default:

- clean, web-safe semantic HTML for display and publishing;
- OfficeIMO round-trip HTML that carries private metadata and binary payloads when explicitly requested.

## P2 gaps: editing, scale, and developer experience

### 9. Structural editing remains shallow

`RtfLosslessEditor` can replace visible text, append a paragraph, and update document settings, metadata, fonts, colors, styles, page setup, revisions, file references, XML namespaces, and related tables while preserving untouched syntax. It cannot yet insert/remove/move arbitrary paragraphs or tables, replace images, edit header/footer content, or operate on bookmark ranges without normalizing the document.

The semantic model also needs block-level insert/remove/move, document clone/merge, and cross-run rich replacement. These are native RTF model operations. Mail merge, comparison, and field evaluation should continue to route through `OfficeIMO.Word`.

### 10. No performance or resilience baseline exists

No dedicated `OfficeIMO.Rtf.Benchmarks` project was found. The parser duplicates input into source text, raw token text, syntax nodes, semantic objects, and binary arrays. That may be acceptable for normal files, but there is no evidence for large clinical notes, generated reports, embedded-image documents, or hostile inputs.

Add benchmarks for parse, lossless round trip, semantic rewrite, Word/HTML/Markdown/PDF conversion, and Reader extraction. Track throughput, allocations, peak working set, and output size for small, medium, and large documents. Add fuzz/property tests for malformed groups, control parameters, Unicode fallback, binary lengths, and lossless preservation.

### 11. The product surface is fragmented in documentation

The package READMEs are concise, but there is no single current conversion graph, no safe-ingestion example, no strict-diagnostic example, and no end-to-end workflow example. The living support matrix was stale enough to mark the published PDF package as deferred.

The support matrix should be generated or validated from a machine-readable capability manifest and test evidence. A small CLI/sample is useful only after the engine contracts are safe and truthful.

## Competitive position

The following comparison uses vendor-documented public capabilities, not benchmark claims.

| Library | Vendor-documented position | OfficeIMO implication |
| --- | --- | --- |
| OfficeIMO | MIT, dependency-free RTF syntax and semantic engine; lossless source path; Word, HTML, Markdown, PDF, and Reader adapters. | Strong open-source architecture. Needs safe defaults, unified diagnostics, external corpus, and workflow polish. |
| [Aspose.Words](https://docs.aspose.com/words/net/) | [Loads and saves RTF plus many Word, web, fixed-layout, image, and e-book formats](https://docs.aspose.com/words/net/supported-document-formats/); documents rendering, printing, mail merge, and reporting. | OfficeIMO should not chase every output format. It must make its smaller graph dependable and diagnosable. |
| [GemBox.Document](https://www.gemboxsoftware.com/document/docs/introduction.html) | Unified model for RTF, DOC/DOCX, PDF, HTML, Markdown, ODT, text, and XPS; mail merge, rendering, and printing. | The largest gap is a polished single-document workflow, not the RTF parser itself. |
| [Syncfusion DocIO](https://help.syncfusion.com/document-processing/word/word-library/net/feature-matrix) | Detailed RTF feature matrix, mail merge, rich find/replace, merge/split, PDF/PDF-A/PDF-UA, and image export. | OfficeIMO needs an evidence-backed matrix and thin RTF workflows over existing Word capabilities. |
| [Telerik WordsProcessing](https://docs.telerik.com/devtools/document-processing/libraries/radwordsprocessing/formats-and-conversion/rtf/rtfformatprovider) | RTF import/export over a flow document, with explicit operation timeouts; separate [PDF](https://docs.telerik.com/devtools/document-processing/libraries/radwordsprocessing/formats-and-conversion/pdf/pdfformatprovider) and [mail-merge](https://docs.telerik.com/devtools/document-processing/libraries/radwordsprocessing/editing/mail-merge) features. | Bounded operations and a unified flow workflow are expected product behavior. |
| [TX Text Control](https://docs.textcontrol.com/textcontrol/asp-dotnet/article.aspnet.introduction.htm) | Server document processing, reporting/mail merge, RTF/DOCX/PDF conversion, and browser editor/viewer products. | WYSIWYG UI is a separate product category and should not enter the RTF core roadmap. |
| [SautinSoft.Document](https://sautinsoft.com/products/document/help/net/developer-guide/document.php) | Create/read/write/convert/merge PDF, DOCX, RTF, HTML, text, and images with templates/mail merge. | OfficeIMO should expose its existing Word workflow cleanly through the RTF bridge. |
| [RtfPipe](https://github.com/erdomke/RtfPipe) | MIT RTF-to-HTML/text conversion with tables, nested tables, lists, hyperlinks, pictures, headings, and HTML encapsulation. | OfficeIMO is broader, but still needs Outlook encapsulation and nested-table parity. |

Commercial documentation also shows that security is part of the product. For example, [Aspose documents resource-loading risks and callbacks](https://docs.aspose.com/words/net/web-applications-security-when-loading-external-resources/), and Telerik's RTF import/export APIs accept operation timeouts. OfficeIMO needs equivalent first-party boundaries rather than leaving them entirely to hosts.

## Recommended delivery order

### Phase 0A - Safe RTF-to-HTML output

- Add `HtmlUrlPolicy` to `RtfToHtmlOptions`.
- Define web-safe and round-trip profiles.
- Block unsafe run and field links with stable diagnostics.
- Gate object data, file references, private metadata, and embedded images by policy.
- Prove hostile-link and payload cases.

Acceptance: untrusted RTF cannot produce executable HTML links or unbounded private payload attributes under the default web profile.

### Phase 0B - Bounded core parsing

- Add the untrusted RTF read profile and core limits.
- Thread limits and cancellation through tokenizer, syntax parser, binder, and all file/stream routes.
- Reuse that profile from Reader, Word, HTML, Markdown, and PDF entry points.

Acceptance: every public ingestion surface can be bounded before excessive allocation and reports the exact limit that stopped conversion.

### Phase 0C - Diagnostic truth

- Report unknown ignorable destinations and every semantic omission.
- Add `RtfConversionReport` and strict mode.
- Add result-returning Word APIs first, then align the other adapters.

Acceptance: a caller can require no loss and receive a deterministic failure whenever content is flattened, omitted, or blocked.

### Phase 1 - Real interoperability corpus

- Add producer manifests and real files.
- Add semantic, normalized, target-reopen, and visual assertions.
- Publish the scorecard from test data.

Acceptance: every `Full` claim in the matrix points to a public API, focused test, and real producer fixture where applicable.

### Phase 2 - Word bridge and workflow parity

- Preserve styles and numbering structures.
- Map or diagnose objects, shapes, and unsupported Word elements.
- Add thin RTF mail-merge, find/replace, field-update, compare, and merge samples/results over `OfficeIMO.Word`.

Acceptance: common RTF automation workflows return output plus a combined conversion/workflow report without duplicating Word engines.

### Phase 3 - High-value RTF specification gaps

Implement in this order:

1. Outlook HTML encapsulation.
2. DBCS/East Asian text and composite fonts.
3. Nested tables.
4. Theme, quick-style, and table-style semantics.
5. Move revisions, protection exceptions, index/TOC entries.
6. Custom XML and smart-tag classification/modeling where real fixtures justify it.

Acceptance: unsupported families are at least recognized, preserved, classified, and diagnosed; the top three are semantically usable.

### Phase 4 - Adapter fidelity

- RTF notes to Markdown footnotes and complete media callbacks.
- Clean versus round-trip HTML profiles.
- PDF shape/object warnings, reusable image conversion, note layout, font diagnostics, and image/page workflows.
- Reader provenance and safety examples over the real corpus.

### Phase 5 - Editing, performance, and product polish

- Structural semantic and lossless edits.
- Benchmark and fuzz lanes with budgets.
- Current NuGet examples, conversion graph, safe/strict recipes, and optional CLI/sample workflows.

## Ownership boundaries

| Capability | Owner |
| --- | --- |
| RTF limits, tokenizer/parser, syntax tree, semantic model, writer, lossless edit, and shared RTF conversion report | `OfficeIMO.Rtf` |
| Word mapping and thin access to mail merge/find/replace/fields/compare/merge | `OfficeIMO.Word.Rtf`, calling `OfficeIMO.Word` |
| URL/resource policy and clean versus round-trip HTML | `OfficeIMO.Html` |
| Markdown AST mapping | `OfficeIMO.Rtf.Markdown` |
| Image decoding/conversion shared by document formats | `OfficeIMO.Drawing` |
| Pagination, PDF warnings, and PDF rendering | `OfficeIMO.Pdf`, with mapping in `OfficeIMO.Rtf.Pdf` |
| Chunking and ingestion projection | `OfficeIMO.Reader.Rtf`, calling bounded `OfficeIMO.Rtf` APIs |
| CLI, examples, websites, and PowerShell | Thin consumers only |

## Strategic non-goals

- Do not put a WYSIWYG editor, browser viewer, or printing subsystem in `OfficeIMO.Rtf`.
- Do not implement separate RTF mail merge, comparison, or field engines while `OfficeIMO.Word` already owns them.
- Do not claim semantic support for every RTF control word. Preserve obscure syntax and diagnose the boundary until real workflows justify a model.
- Do not call PDF import lossless or visual reconstruction.
- Do not add streaming architecture until bounded benchmarks show where it pays off.

The best market position is not “free Aspose.” It is a transparent, MIT, cross-platform RTF engine with a lossless source path, safe defaults, explicit degradation, and first-party OfficeIMO workflows that can be proven against real documents.
