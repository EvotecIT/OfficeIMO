# OfficeIMO.Pdf Library Cleanup Plan

Date: 2026-06-03
Branch: `codex/pdf-library-cleanup-20260603`
Base: `origin/master` / `master` at `e472a4b2`

## Current State

Local `master` is up to date with `origin/master`. The current head is `e472a4b2`, `Merge pull request #1882 from EvotecIT/codex/officeimo-pdf-font-system-family`.

`OfficeIMO.Pdf` is already broad enough to be the PDF foundation for PSWriteOffice and PSWritePDF replacement work. The current source has:

- dependency-light PDF generation with `OfficeIMO.Drawing` as the shared drawing/font/image layer;
- fluent `PdfDocument` / `PdfOptions` / compose APIs for generated documents;
- Word, Excel, and Markdown adapters that map source semantics into `OfficeIMO.Pdf` primitives;
- PDF read/probe/preflight, text extraction, structured/logical readback, images, attachments, forms, links, outlines, page labels, catalog metadata, and capability diagnostics;
- page extraction, split-style extraction, merge, import, reorder, duplicate, delete, rotate, metadata edit, text/image stamp/watermark, simple AcroForm fill, and simple flattening;
- `PdfPageSelection` as the shared page-selection vocabulary for fluent page operations and readback;
- `PdfOperationResult<T>` as the shared preflight-backed result vocabulary for page, merge, readback, stamp, metadata, and form workflows that need structured diagnostics instead of direct exceptions;
- `PdfSaveResult` as the shared save/output result vocabulary for byte counts, output paths, and I/O diagnostics without requiring PDF parse preflight;
- Word, Excel, and Markdown PDF adapters expose `TrySaveAsPdf` entry points that return the core `PdfSaveResult` instead of inventing adapter-specific result types;
- Poppler-backed raster visual baselines plus text/structure tests. Raster assertions now compare decoded pixels, not raw PNG bytes, and allow a tiny default 32-pixel Poppler/font edge-noise budget while still failing geometry and material layout changes. The raster baseline tests are now split into the test catalog, reusable raster support, native/adapter scenarios, core routing, simple core scenarios, composition scenarios, and report scenarios.
- Read-layout smoke tests are now split by contract: extension API smoke coverage, structured-by-page extraction, page-range/table-wrapper extraction, and shared fixtures.
- The visual-quality test family is now split by contract instead of carried by two giant catch-all files: document options/metadata/decorations, rich text, default styles/themes, table layout/paint/borders/styles/spans/validation, links/bookmarks, typography, flow keep behavior, panels, rules/media, row/report fixtures, vector primitives, and shared support helpers all have named partial files.
- Compliance analyzer tests are now split by proof family: PDF/A readiness, Factur-X overview/header/line/settlement/tax/payment/attachment checks, PDF/UA readiness, option overloads, and focused CII/ICC support fixtures.
- Read-stream and catalog rewrite tests are now split by contract: stream/path validation, outline rewriting, trailer-root catalog selection, catalog metadata, named destinations, open actions, viewer/metadata structures, embedded files, optional/active content, and supporting hand-built PDF fixtures.
- Inspector tests are now split by contract: core readback, catalog/preflight checks, destinations, embedded/optional content, links/bookmarks, malformed content, outlines, security/forms, viewer metadata, and focused fixture/assertion helpers.
- Compose-page option tests are now split by contract: page setup and validation, scoped defaults/flow, background flow, dictionary construction, header/footer rendering, footer segment layout, header/footer variants, validation, and shared test support.
- Reader/footer regression tests are now split by contract: syntax/page boxes, footer font and segment behavior, text parser and marked-content handling, page-tree/content-array traversal, text operators and metadata strings, stream filters/decode parameters, form XObjects, and focused fixture helpers.
- Excel PDF export tests are now split by adapter contract: basic sheet selection, worksheet images, page setup/print areas/page breaks, header/footer mapping, cell styles and number formats, hyperlinks, row/column/layout behavior, chart snapshots, options/warnings, and shared geometry/image support helpers.
- Page editor tests are now split by operation contract: delete, duplicate, move, reorder, rotate, wrapper path/stream pipelines, validation/output-target failures, and shared PDF/stream fixtures.
- External document compatibility tests are now split by external-PDF contract: text/font extraction, split/merge of external producer files, xref/incremental-reference handling, object stream rewrite/preflight behavior, and focused text/xref/object-stream fixture helpers.
- Compliance assessment tests are now split by readiness/structure contract: PDF/A font readiness, PDF/UA readiness, tagged text/images/links, tagged lists/tables, page chrome/drawings, validation failures, and shared fixture/assertion helpers.
- Word section PDF export tests are now split by adapter contract: basic section/header smoke coverage, explicit page setup, Word section columns, first/even header-footer variants, rich header-footer content, body/content-control export, and shared PDF text/image helper assertions.
- Logical document tests are now split by readback contract: logical model construction, Markdown projection, page-range filtering, layout fixtures, AcroForm readback, navigation/link elements, and shared generated-PDF fixture helpers.
- Page extractor tests are now split by operation contract: extraction ordering/ranges, range parsing, wrapper pipeline round-trips, stream pipelines, path pipelines, split outputs, resource/link/reference preservation, validation failures, and shared hand-built PDF/stream fixtures.
- Word paragraph/table PDF export tests are now split by adapter contract: basic paragraph/table output, paragraph formatting and layout, headings/TOC/links, native list handling, run/background styling, table rich runs, and shared count helpers.
- Form filler tests are now split by form workflow contract: fill/value appearance, flattening, choice fields, stream/path wrapper pipelines, validation/signed-PDF failures, and shared hand-built form PDF fixtures.
- Stamper tests are now split by stamp workflow contract: text stamping/watermarking, text path/stream pipelines, image stamping, image/watermark pipelines, placement/layering, options/validation, and shared PDF/image/stream fixtures.
- Text extractor page tests are now split by readback contract: basic page/range extraction, stream-position handling, layout/span metrics, all-text wrapper pipelines, file output, Markdown output, validation failures, and shared PDF/stream fixtures.
- Word table-style PDF export tests are now split by adapter contract: cell border/fill rendering, Word/table-level styles, margins and spacing, row/cell layout, repeating headers, placement/preferred width, native style mapping, and shared table fixture helpers.
- Rich paragraph wrapping tests are now split by writer contract: font selection, WinAnsi/diagnostic handling, monospace and hard line breaks, glyph-width wrapping, rich tab layout, rendered paragraph tabs, and shared reflection helpers.

The problem is not that there should be a separate advanced or alternate library. There should be one `OfficeIMO.Pdf` library with one coherent API and a higher quality bar.

## Active Branch Context

There are no open PDF PRs in `EvotecIT/OfficeIMO` as of this check.

Local PDF worktrees still matter:

- `C:\Support\GitHub\OfficeIMO-pdf-flow-quality-gates` is already an ancestor of `origin/master`; treat it as merged history.
- `C:\Support\GitHub\OfficeIMO-pdf-external-pdf-ops` points at a merged PR name but has local commits not on `origin/master`; it is stale/superseded history and should only be mined carefully.
- `C:\Support\GitHub\OfficeIMO-pdf-compliance-gates` points at merged PR #1881 but has many local commits not on `origin/master`, including embedded-font, tagged-structure, form-appearance, and CII readiness follow-ups. This branch is the most important shelf to triage before new compliance/font work.

## Cleanup Direction

1. Keep one public document type.

Use `PdfDocument.Create(...)` for generated documents and `PdfDocument.Open(...)` for existing PDFs. Do not maintain a separate legacy creation path and a separate loaded-document wrapper.

2. Keep low-level helpers as implementation and wrapper support.

Static helpers such as `PdfMerger`, `PdfPageExtractor`, `PdfPageEditor`, `PdfTextExtractor`, and `PdfInspector` are still useful for thin wrappers and focused low-level calls. They should not be the main experience users have to learn first.

3. Make read, merge, split, edit, stamp, and forms first-class workflows.

The main user-facing shape should be:

```csharp
PdfDocument.Create()
    .H1("Report")
    .Paragraph(p => p.Text("Generated content"))
    .MergeWith("appendix.pdf")
    .Stamp.Text("Reviewed")
    .Save("final.pdf");

PdfDocument.Open("final.pdf")
    .Pages.Extract("1-2,4")
    .Save("selection.pdf");

string text = PdfDocument.Open("final.pdf").Read.Text();
IReadOnlyList<PdfDocument> pages = PdfDocument.Open("final.pdf").Pages.Split();
```

4. Replace overload growth with shared models.

Add `PdfPageSelection`, operation results, export results, and structured diagnostics so callers get one predictable vocabulary across generation, readback, manipulation, and adapters.

5. Treat compliance as proof, not marketing.

PDF/A, PDF/UA, Factur-X/ZUGFeRD, output intents, metadata, tags, alternate text, and forms have strong groundwork, but formal claims should stay disabled until validator-backed artifacts exist.

6. Put font and text quality in the core.

Unicode writing, TrueType subsetting, fallback chains, and extraction-preserving ToUnicode coverage belong in `OfficeIMO.Pdf` / `OfficeIMO.Drawing`, not in Word/Excel/Markdown adapters.

7. Split large tests as they are touched.

The largest PDF test files are carrying too many scenarios. New cleanup should split by feature or contract so failures stay readable.

## Immediate Next Steps

1. Keep the `PdfDocument` unification intact: generated documents use `PdfDocument.Create(...)`, loaded PDFs use `PdfDocument.Open(...)`, and source, examples, docs, and tests stay on that single public document type.
2. Keep contract tests around create/open and fluent read/manipulation workflows on the same `PdfDocument` type.
3. Push `PdfPageSelection` into any remaining wrapper entry points that genuinely need page ranges instead of adding ad hoc string-only overloads. The Word, Excel, Markdown, and Reader PDF adapters do not currently own user-facing page-range workflows; range-heavy static APIs remain in `OfficeIMO.Pdf` low-level read/manipulation helpers.
4. Reconcile the local compliance-gates branch before deeper compliance/font changes.
5. Keep direct fluent APIs chainable while using result APIs for wrapper/automation surfaces that need structured outcomes.
6. Continue splitting large PDF test files by feature/contract as they are touched. Raster visual baselines, read-layout extraction, visual quality, compliance analyzer, read-stream/catalog rewrite, inspector, compose-page option, reader/footer regression, Excel PDF export, page-editor, external compatibility, compliance assessment, Word section export, logical-document, page-extractor, Word paragraph/table export, form-filler, stamper, text-extractor page, Word table-style, and rich paragraph wrapping tests are split; the next cleanup candidates are `PdfDocumentRasterVisualBaselineCoreCompositionScenarios.cs`, `PdfDocumentVisualQualityThemeTests.cs`, `PdfDocumentImageValidationTests.cs`, and `PdfFormCreationTests.cs`.

## Validation

Minimum validation for the cleanup branch:

- `dotnet build .\OfficeIMO.Pdf\OfficeIMO.Pdf.csproj -c Release /nr:false`
- `dotnet build .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~PdfDocumentWorkflowTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~PdfDocumentRasterVisualBaselineTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~PdfDocumentVisualQualityTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfComplianceAnalyzerTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfReadStreamTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfInspectorTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfComposePageOptionsTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfReaderAndFooterRegressionTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~SaveAsPdf_ExcelWorkbook" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfPageEditorTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfExternalDocumentCompatibilityTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfDocumentComplianceAssessmentTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~SaveAsPdf_OfficeIMOEngine|FullyQualifiedName~Test_WordDocument_SaveAsPdf" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfLogicalDocumentTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfPageExtractorTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~SaveAsPdf_OfficeIMOEngine|FullyQualifiedName~Test_WordDocument_SaveAsPdf|FullyQualifiedName~SaveAsPdf_Renders_Paragraphs|FullyQualifiedName~SaveAsPdf_Renders_Tables" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfFormFillerTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfStamperTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~PdfTextExtractorPageTests" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~Test_WordDocument_SaveAsPdf_TableStyles|FullyQualifiedName~SaveAsPdf_OfficeIMOEngine_Renders_Table_|FullyQualifiedName~SaveAsPdf_OfficeIMOEngine_Maps_Table_" /nr:false`
- `dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~RichParagraphWrappingTests" /nr:false`
- `git diff --check`
