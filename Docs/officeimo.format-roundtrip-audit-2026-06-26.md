# OfficeIMO Format Roundtrip Audit

Date: 2026-06-26

Baseline branch: `codex/excel-image-premium-drawing-consolidation`

Worktree: `C:\Support\GitHub\_worktrees\OfficeIMO-format-roundtrip-audit-20260626`

## Short Answer

OfficeIMO already has most of the right package boundaries for HTML, Word, Markdown, PDF, Excel, PowerPoint, Reader, and semantic authoring. The missing work is not another converter package by default. The missing work is a shared roundtrip/proof contract that makes every lane report the same feature inventory, visual theme vocabulary, diagnostics, asset manifest, and visual evidence story without forcing every format through one universal document model.

The strongest existing owners are:

- `OfficeIMO.Html` for bounded HTML parsing, URL/resource policy, normalized HTML, shared diagnostics, and `HtmlConversionDocument`.
- `OfficeIMO.Markdown` for the Markdown AST and `MarkdownVisualTheme`.
- `OfficeIMO.Pdf` for PDF generation, logical readback, table extraction, conversion diagnostics, layout primitives, and PDF proof.
- `OfficeIMO.Drawing` for reusable raster/SVG/image/text/shape/chart rendering primitives.
- `OfficeIMO.Markup` for reading and writing markdown-like Office authoring content into specific target formats.
- Format adapters such as `OfficeIMO.Word.Html`, `OfficeIMO.Markdown.Html`, `OfficeIMO.Html.Pdf`, `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, and `OfficeIMO.PowerPoint.Pdf`.

The main risk is letting those adapters continue to improve independently until "nice roundtrip" means something different for each package.

## HTML End-To-End Goal

The HTML work should be tackled as seven connected lanes, all on top of the Excel image consolidation branch so visual/image reuse keeps flowing through `OfficeIMO.Drawing`.

| Point | Decision | First implementation/proof slice |
| --- | --- | --- |
| 1. HTML profiles | Add explicit profiles instead of one ambiguous HTML mode: `Semantic`, `Document`, `HighFidelityPrint`, and `PositionedReview`. | `OfficeIMO.Html` now exposes all four shared profile contracts, including `PositionedReview` for geometry/review HTML. |
| 2. Shared HTML owner | Keep HTML parsing, resource policy, diagnostics, profile declarations, asset manifests, and HTML proof in `OfficeIMO.Html`. | Do not create a new converter package for this until a public cross-package CLR contract is truly needed. |
| 3. HTML -> Word -> HTML | Treat this as the primary editable document roundtrip. | `OfficeIMO.Word.Html` now has a public `SaveHtmlCapabilityGallery` path that records source HTML, valid DOCX, round-trip HTML, assets, diagnostics, score, and the source-specific `WordDocumentRoundTrip` office profile. |
| 4. HTML -> Markdown -> HTML | Treat this as the semantic Markdown roundtrip, not a visual layout lane. | Default style-less Markdown HTML/Word/PDF exports now resolve through the shared `MarkdownVisualTheme.Default()` Word-like profile, with explicit opt-outs for plain output. |
| 5. HTML/PDF profiles | Make HTML-to-PDF and PDF-to-HTML profile choices explicit: semantic/document input to PDF, semantic/positioned-review HTML from PDF. | Reuse `OfficeIMO.Pdf` reports and `PdfLogicalDocument`; add fixture manifests before claiming editable reconstruction. |
| 6. Excel/PowerPoint HTML | Add thin adapter packages because direct references from `OfficeIMO.Html` back into Excel/PowerPoint would invert ownership. | `OfficeIMO.Excel.Html` and `OfficeIMO.PowerPoint.Html` now emit semantic HTML and visual-review HTML while mapping to the shared `OfficeIMO.Html` contracts. |
| 7. Proof/gallery manifest | Standardize a schema-first proof manifest across HTML lanes. | `OfficeIMO.Html` gallery manifests now write profile contract details and typed roundtrip expectations for preserved, simplified, blocked, omitted, reported, text-proof, and visual-proof outcomes. |

### First Slice Completed

The first code slice addresses the "style-less output looks bad" problem without adding another package or theme brain:

- `OfficeIMO.Markdown` now exposes `MarkdownVisualTheme.Default()` and `MarkdownVisualTheme.ResolveOrDefault(...)` as the shared Markdown-owned resolver.
- Markdown HTML rendering applies the shared default visual theme for `Clean`/`Word` styled output when no explicit `VisualTheme` is supplied.
- `HtmlOptions.ApplyDefaultVisualTheme = false` preserves plain/legacy HTML output.
- Markdown-to-Word applies the shared default theme when no explicit `Theme` is supplied.
- `MarkdownToWordOptions.ApplyDefaultTheme = false` preserves plain Word output.
- Markdown-to-PDF fallback now flows through `MarkdownVisualTheme.Default()` before mapping into `MarkdownPdfVisualTheme`, keeping HTML, Word, and PDF palettes aligned.

Validation:

```text
dotnet test OfficeIMO.Markdown.Tests\OfficeIMO.Markdown.Tests.csproj --filter "FullyQualifiedName~Markdown_VisualTheme_Tests|FullyQualifiedName~Markdown_Html_Render_Tests|FullyQualifiedName~MarkdownToWord_AppliesDefaultSharedVisualThemeWhenThemeIsOmitted|FullyQualifiedName~MarkdownToWord_CanDisableDefaultSharedVisualTheme"
dotnet test OfficeIMO.Pdf.Tests\OfficeIMO.Pdf.Tests.csproj --filter "FullyQualifiedName~MarkdownSaveAsPdfVisualTests"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 132 filtered tests.

### Second Slice Completed

The second code slice closes the profile vocabulary gap for point 1:

- `OfficeIMO.Html.HtmlConversionProfile` now includes `PositionedReview`.
- `HtmlConversionProfileContracts` now advertises a `Positioned Review` contract for page wrappers, positioned text/images, link frames, form field frames, review-safe CSS, resource handling, and explicit no-editable-reconstruction diagnostics.
- Existing adapter behavior remains unchanged; this is a shared contract label that PDF-to-HTML and future visual-review lanes can report against.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlEnginePlatform|FullyQualifiedName~HtmlPdf_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~PdfHtml_ProfileContracts_CoverSupportedProfiles"
dotnet test OfficeIMO.Markdown.Tests\OfficeIMO.Markdown.Tests.csproj --no-restore --filter "FullyQualifiedName~Markdown_VisualTheme_Tests|FullyQualifiedName~Markdown_Html_Render_Tests|FullyQualifiedName~MarkdownToWord_AppliesDefaultSharedVisualThemeWhenThemeIsOmitted|FullyQualifiedName~MarkdownToWord_CanDisableDefaultSharedVisualTheme"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 45 filtered tests.

### Third Slice Completed

The third code slice starts the schema-first proof/gallery manifest without adding a new package:

- Added `HtmlCapabilityGalleryExpectation` and `HtmlCapabilityGalleryExpectationOutcome` under `OfficeIMO.Html`.
- `HtmlCapabilityGalleryManifest` can now carry typed expectations for preserved, simplified, blocked, omitted, reported, text-proof, and visual-proof outcomes.
- `HtmlCapabilityGalleryManifestWriter` now emits a `Profile Contract` section and `Roundtrip Expectations` section, so gallery outputs expose supported HTML/CSS/resource/diagnostic contracts and explicit evidence expectations in the same artifact.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlEnginePlatform|FullyQualifiedName~HtmlPdf_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~PdfHtml_ProfileContracts_CoverSupportedProfiles"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 15 filtered tests.

### Fourth Slice Completed

The fourth code slice prevents the HTML/PDF adapters from becoming a second profile brain:

- `HtmlPdfProfileContract` now exposes the shared `OfficeIMO.Html.HtmlConversionProfile` each HTML-to-PDF adapter profile maps to.
- `HtmlPdfProfile.Semantic` maps to the shared `Semantic` profile and `HtmlPdfProfile.Document` maps to the shared `Document` profile.
- `PdfHtmlProfileContract` now exposes the shared profile for PDF-to-HTML output lanes.
- `PdfHtmlProfile.Semantic` maps to the shared `Semantic` profile and `PdfHtmlProfile.PositionedReview` maps to the shared `PositionedReview` profile.
- Adapter-specific profile enums remain in the adapter package, but the public contract can now be compared against one shared HTML vocabulary.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlPdf_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~PdfHtml_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~HtmlEnginePlatform" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 15 filtered tests.

### Fifth Slice Completed

The fifth code slice makes the shared Markdown visual theme choices discoverable instead of hidden behind hardcoded factory methods:

- Added `MarkdownVisualThemePreset` as a typed descriptor for built-in theme choices.
- `MarkdownVisualTheme.Presets` now exposes the stable built-in preset catalog: `Plain`, `WordLike`, `TechnicalDocument`, `GitHubLike`, `Compact`, and `Report`.
- `MarkdownVisualTheme.ColorSchemes` now exposes the built-in accent/color choices: `Default`, `Blue`, `Emerald`, `Indigo`, `Rose`, `Amber`, and `Slate`.
- Added `MarkdownVisualTheme.Create(kind, colorScheme)` and `TryCreate(name, colorScheme, out theme)` so front matter, APIs, and UI pickers can choose the same named visual preset plus palette.
- Existing `TryCreate(name, out theme)` behavior remains compatible and now resolves through the typed preset catalog.

Validation:

```text
dotnet test OfficeIMO.Markdown.Tests\OfficeIMO.Markdown.Tests.csproj --no-restore --filter "FullyQualifiedName~Markdown_VisualTheme_Tests|FullyQualifiedName~Markdown_Html_Render_Tests|FullyQualifiedName~MarkdownToWord_AppliesDefaultSharedVisualThemeWhenThemeIsOmitted|FullyQualifiedName~MarkdownToWord_CanDisableDefaultSharedVisualTheme" --logger "console;verbosity=minimal"
dotnet test OfficeIMO.Pdf.Tests\OfficeIMO.Pdf.Tests.csproj --no-restore --filter "FullyQualifiedName~MarkdownSaveAsPdfVisualTests" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 134 filtered tests.

### Sixth Slice Completed

The sixth code slice declares the Excel and PowerPoint HTML lanes in the shared HTML owner without starting new converter packages:

- Added `OfficeHtmlConversionProfile` for `ExcelSemanticTables`, `ExcelVisualReview`, `PowerPointSemanticSlides`, and `PowerPointVisualReview`.
- Added `OfficeHtmlConversionProfileContract` and `OfficeHtmlConversionProfileContracts` under `OfficeIMO.Html`.
- Semantic Excel/PowerPoint lanes map to the shared `HtmlConversionProfile.Semantic` contract.
- Visual-review Excel/PowerPoint lanes map to the shared `HtmlConversionProfile.PositionedReview` contract.
- Visual-review lanes explicitly name `OfficeIMO.Drawing` as the reusable visual primitive owner, so future adapters should use the Excel image branch drawing work instead of building separate renderers.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlEnginePlatform|FullyQualifiedName~HtmlPdf_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~PdfHtml_ProfileContracts_CoverSupportedProfiles" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 16 filtered tests.

### Seventh Slice Completed

The seventh code slice connects the real HTML -> Word -> HTML gallery proof to the shared manifest system:

- `HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml` now uses `HtmlCapabilityGalleryManifestWriter` instead of a local test-only manifest writer.
- The Word roundtrip gallery now records the shared `Document` profile contract, roundtrip score, resource manifest, artifacts, diagnostics, and typed expectations in one manifest.
- The Word roundtrip gallery now also records the source-specific `WordDocumentRoundTrip` office profile contract, keeping Word evidence on the same `officeProfiles` manifest path as Excel and PowerPoint.
- The scenario carries expectations for preserved headings, table sections, form controls, and reported comments.
- This keeps the primary editable HTML roundtrip lane on the same proof vocabulary as the shared HTML platform tests and future Excel/PowerPoint/PDF lanes.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml|FullyQualifiedName~HtmlEnginePlatform" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 15 filtered tests.

### Eighth Slice Completed

The eighth code slice replaces the previous contract-only Excel/PowerPoint HTML story with real first-party adapters and the first rich-content proof slice:

- Added `OfficeIMO.Excel.Html` as a thin adapter over `OfficeIMO.Excel`, `OfficeIMO.Html`, and `OfficeIMO.Drawing`.
- Excel semantic HTML exports workbook/worksheet sections, used-range tables, formulas, comments, chart inventory, and image inventory/previews through the shared OfficeIMO HTML shell.
- Excel visual-review HTML uses the existing Excel SVG image exporter from the image consolidation branch and marks the visual owner as `OfficeIMO.Drawing`; the proof fixture shows table values, formula text, visible comment callout/list proof, a visible image, and a rendered chart.
- Added `OfficeIMO.PowerPoint.Html` as a thin adapter over `OfficeIMO.PowerPoint`, `OfficeIMO.Html`, and shared visual contracts.
- PowerPoint semantic HTML exports slide text, tables, picture inventory/previews, chart snapshot inventory, and slide-aligned extraction proof for notes/tables without creating a separate interchange model.
- PowerPoint visual-review HTML exports positioned slide canvases from public shape geometry, embeds pictures as data URIs, and renders supported chart snapshots through the shared `OfficeIMO.Drawing` SVG chart renderer while keeping an explicit placeholder diagnostic fallback for unsupported charts.
- Added `OfficeHtmlDocumentShell`, `OfficeHtmlDocumentOptions`, `OfficeHtmlDocumentThemeKind`, and `OfficeHtmlText` to `OfficeIMO.Html` so generated adapter output shares one CSS/theme shell.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --filter "FullyQualifiedName~HtmlOfficeAdapters|FullyQualifiedName~HtmlEnginePlatform" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 19 filtered tests.

Visual proof artifacts:

- `output/html-visual-proof/Generator/artifacts/excel-semantic.html`
- `output/html-visual-proof/Generator/artifacts/excel-visual.html`
- `output/html-visual-proof/Generator/artifacts/powerpoint-semantic.html`
- `output/html-visual-proof/Generator/artifacts/powerpoint-visual.html`
- Rich screenshots under `output/html-visual-proof/screenshots/*-rich.png`

Broader regression validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml|FullyQualifiedName~HtmlEnginePlatform|FullyQualifiedName~HtmlPdf_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~PdfHtml_ProfileContracts_CoverSupportedProfiles|FullyQualifiedName~HtmlOfficeAdapters" --logger "console;verbosity=minimal"
dotnet test OfficeIMO.Markdown.Tests\OfficeIMO.Markdown.Tests.csproj --no-restore --filter "FullyQualifiedName~Markdown_VisualTheme_Tests|FullyQualifiedName~Markdown_Html_Render_Tests|FullyQualifiedName~MarkdownToWord_AppliesDefaultSharedVisualThemeWhenThemeIsOmitted|FullyQualifiedName~MarkdownToWord_CanDisableDefaultSharedVisualTheme" --logger "console;verbosity=minimal"
dotnet test OfficeIMO.Pdf.Tests\OfficeIMO.Pdf.Tests.csproj --no-restore --filter "FullyQualifiedName~MarkdownSaveAsPdfVisualTests" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 156 filtered tests.

### Ninth Slice Completed

The ninth code slice makes the shared proof/gallery manifest machine-readable without introducing a new converter package or cross-format CLR contract assembly:

- Added `HtmlCapabilityGalleryManifestJsonWriter` in `OfficeIMO.Html`.
- The JSON payload carries a stable schema id/version, scenario metadata, shared profile contract, roundtrip expectations, artifacts with hashes, roundtrip score and metrics, resource inventory, and diagnostics.
- The existing Markdown manifest remains for human review; JSON is the schema-shaped proof artifact for hosts, CI, galleries, and future non-HTML adapters to map into.
- `HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml` now writes both `quarterly-report.manifest.md` and `quarterly-report.manifest.json`.
- Tests parse the generated JSON with `JsonDocument` on all target frameworks, so the contract is not just string-shaped output.

Validation:

```text
dotnet test OfficeIMO.Html.Tests\OfficeIMO.Html.Tests.csproj --no-restore --filter "FullyQualifiedName~HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml|FullyQualifiedName~HtmlEnginePlatform_ConnectsProfilesIrStylesResourcesScoringDiagnosticsAndGallery" --logger "console;verbosity=minimal"
```

Result: passed on `net10.0`, `net8.0`, and `net472` with 2 filtered tests.

## Current Capabilities

| Lane | Current owner | Current shape | Roundtrip quality |
| --- | --- | --- | --- |
| HTML to Word | `OfficeIMO.Word.Html` over `OfficeIMO.Html` | Strong document import: blocks, inline formatting, lists, tables, images/SVG, stylesheets, resource limits, diagnostics, forms, notes, comments, headers/footers. | Best current editable HTML roundtrip lane. Not browser-layout complete. |
| Word to HTML | `OfficeIMO.Word.Html` | Exports semantic document HTML with optional CSS, metadata, notes, comments, lists, tables, images, figures, bookmarks, forms. The shared HTML catalog now declares `WordSemanticDocument`, `WordDocumentRoundTrip`, and `WordPrintReview` lanes, and the package exposes `SaveHtmlCapabilityGallery` for HTML -> Word -> HTML proof artifacts. | Good for document HTML. The editable roundtrip lane is now adapter-owned and manifest-backed; print-like output still needs stronger visual proof. |
| HTML to Markdown | `OfficeIMO.Markdown.Html` over `OfficeIMO.Html` | Converts document fragments/full documents into Markdown text or `MarkdownDoc`; supports custom block/inline converters. | Useful, but current architecture still has some post-render and AST-extension gaps from the Markdown readiness review. |
| Markdown to HTML | `OfficeIMO.Markdown` / renderer packages | `MarkdownDoc` renders HTML with options and visual themes; renderer shells add host behavior. | Good for Markdown-owned content. Needs AST-first extension model and neutral preset cleanup. |
| Markdown to Word | `OfficeIMO.Word.Markdown` | Converts Markdown text or `MarkdownDoc` into Word, including shared `MarkdownVisualTheme` use. | Good semantic bridge. Lossy for Word-only features and Markdown-only extensions. |
| Word to Markdown | `OfficeIMO.Word.Markdown` | Converts Word to Markdown and offers Word to HTML via Markdown. | Good for common document content. Needs explicit loss metadata for unsupported Word features. |
| HTML to PDF | `OfficeIMO.Html.Pdf` | Two profiles: semantic HTML via Markdown to PDF, document HTML via Word to PDF. | Correct split. Needs profile proof and stronger CSS/resource contract per profile. |
| PDF to HTML | `OfficeIMO.Html.Pdf` over `OfficeIMO.Pdf` logical model | Semantic HTML and positioned-review HTML from PDF logical readback. | Review/search lane, not editable reconstruction. Needs geometry/image/link/form fixture proof. |
| Word to PDF | `OfficeIMO.Word.Pdf` over `OfficeIMO.Pdf` and `OfficeIMO.Drawing` | Rich native PDF path with many layout/table/typography tests. | Strong, with remaining gaps in typography, complex layout, compliance, and visual proof productization. |
| Excel to PDF | `OfficeIMO.Excel.Pdf` over `OfficeIMO.Pdf` and `OfficeIMO.Drawing` | Exports sheets/ranges/print areas/page setup/tables/images/charts and imports PDF tables to Excel. | Good first-party adapter; premium fidelity depends on shared Drawing and Excel visual snapshot burn-down. |
| Excel to HTML | `OfficeIMO.Excel.Html` over `OfficeIMO.Excel`, `OfficeIMO.Html`, and `OfficeIMO.Drawing` | Semantic workbook/worksheet table HTML plus formulas, comments, chart data tables, image inventory/previews, and visual-review HTML from existing Excel SVG image export. | Explicit rich first slice now exists. Semantic chart categories/series/values can import back into native charts, SVG review covers visible charts/images when supported by the Excel image exporter, and comment bodies surface as dependency-free callout/list proof. Premium depth still needs merged-cell/conditional-format/pivot/filter/slicer/object diagnostics, Excel-exact comment popovers, and richer editable reconstruction rules. |
| Excel to image | `OfficeIMO.Excel` plus `OfficeIMO.Drawing` on this branch | PNG/SVG range/worksheet/workbook work is advanced and actively consolidated into Drawing. | Good architecture, not premium-complete. Tracked approximations remain explicit. |
| PowerPoint to PDF | `OfficeIMO.PowerPoint.Pdf` over `OfficeIMO.Pdf` and `OfficeIMO.Drawing` | Slide-to-page export for backgrounds, text boxes, pictures, tables, charts, shapes; PDF tables back to PPT tables. | Useful first-party lane. Needs master/theme/layout inheritance and richer visual proof. |
| PowerPoint to HTML | `OfficeIMO.PowerPoint.Html` over `OfficeIMO.PowerPoint` and `OfficeIMO.Html` | Semantic slide HTML with extraction proof, picture previews, chart snapshot data tables, and positioned-review slide canvases from public shape geometry. | Explicit rich first slice now exists. Pictures render in positioned review; supported chart snapshots render through shared `OfficeIMO.Drawing` SVG output, and semantic chart categories/series/values can be imported back into native charts. Full master/theme/animation/media/SmartArt fidelity remains future depth. |
| PDF to Word/Excel/PowerPoint | Partial table extraction to Excel/PowerPoint; no editable Word reconstruction found | `PdfLogicalDocument` supports tables/text/images/forms metadata and table import helpers. | Honest partial reconstruction. Full editable PDF-to-Office should stay out of scope until logical proof is stronger. |
| Reader ingestion | `OfficeIMO.Reader` and `OfficeIMO.Reader.*` | Reads many formats into chunks and summaries. | Good AI/search ingestion path. Not a layout roundtrip engine. |
| Semantic authoring | `OfficeIMO.Markup` | Markdown-inspired AST exported to Word/Excel/PowerPoint via target packages. | Good authoring lane. Do not turn it into the universal conversion interchange model. |

## Existing Shared Brains

### `OfficeIMO.Html`

This is already the right shared owner for HTML ingestion. It owns URL policy, base URI resolution, AngleSharp parsing helpers, DOM traversal limits, image discovery, data URI parsing, diagnostics/gallery contracts, and `HtmlConversionDocument`.

What it should own next:

- One shared HTML capability manifest shape for import/export paths.
- One resource manifest and asset-policy contract reused by Word, Markdown, PDF, Reader, and future Excel/PowerPoint HTML lanes.
- Profile-level CSS support declarations, not duplicated adapter notes.

What it should not own:

- Word layout decisions.
- Markdown AST semantics.
- PDF layout/rendering.
- Browser-grade CSS layout.

### `OfficeIMO.Pdf`

This is already the right shared owner for PDF creation, logical readback, table extraction, conversion warnings, typography diagnostics, page/canvas primitives, and proof artifacts.

What it should own next:

- A first-class cross-converter proof result shape: source features, warnings, logical objects, text extraction, image/form/link metadata, hashes, raster pages, and optional validator output.
- A stable positioned/logical model that HTML/PDF/Reader/PDF-to-table adapters can all report against.
- More shared layout primitives before adapters add private pagination, table, text-box, and drawing decisions.

What it should not own:

- Excel workbook semantics.
- PowerPoint master/layout semantics.
- Word style model semantics.
- HTML/CSS browser layout.

### `OfficeIMO.Drawing`

This branch proves Drawing is the correct owner for reusable visual output: raster/SVG rendering, text-block plans, shape presets, image info, chart/drawing primitives, quality analysis, PNG/SVG helpers, and nonblank/visual-baseline support.

What it should own next:

- Common image/SVG/raster proof helpers for all visual converters.
- Generic shape, text-box, image, chart, gradient, stroke, clipping, and transform primitives used by Excel, Visio, Word/PDF, PowerPoint/PDF, and future image/HTML preview lanes.
- Visual theme projection primitives that are not Markdown-specific.

What it should not own:

- Excel cell style interpretation.
- PowerPoint slide/master inheritance.
- Word paragraph/table style resolution.
- PDF object model/rewrite safety.

### `OfficeIMO.Markdown` and `OfficeIMO.Markup`

`OfficeIMO.Markdown` has the real Markdown AST and shared `MarkdownVisualTheme`. `OfficeIMO.Markup` has a markdown-inspired Office authoring AST that can read/write markdown-like content and export to Word, Excel, and PowerPoint.

`OfficeIMO.Markup` should not become the universal interchange model. That would blur its job and create another dependency center that every converter would be tempted to reference.

If the goal is nice roundtrips across HTML, Word, Markdown, Excel, PowerPoint, and PDF, `MarkdownDoc` is too narrow and `OfficeMarkupDocument` is intentionally authoring-oriented. The missing shared place is not a new authoring AST; it is a small set of conversion contracts that existing owners can populate while keeping their native models.

## What Is Missing

### 1. A Roundtrip Contract

Today each lane has tests, warnings, and docs, but there is no single cross-format contract that says:

- which source features were seen;
- which features were preserved, simplified, omitted, or blocked;
- which assets were embedded, linked, rewritten, or dropped;
- which visual theme/profile was applied;
- which diagnostics are expected;
- what logical text/tables/images/forms/links survived;
- what visual proof exists.

Recommendation: avoid a new `.XXX` package unless the dependency graph proves it is cleaner than the alternatives. Do not put these contracts in `OfficeIMO.Html`, `OfficeIMO.Pdf`, or `OfficeIMO.Markup` if non-HTML, non-PDF, or non-markup converters must consume them; that would create wrong-direction dependencies.

Also do not treat the current source-ingested `OfficeIMO.Shared` folder as a public contract assembly. It is fine for internal helper code that each package compiles privately, but public shared types compiled into multiple assemblies would have different CLR type identities. That is acceptable for private implementation helpers and dangerous for APIs that callers pass between packages.

Prefer one of these shapes:

1. Keep proof/report contracts internal to existing owners until at least two independent owners need the exact same type.
2. If the same shape is needed only for artifact manifests, define a dependency-free JSON/schema contract and let each package export to it from its native report type.
3. If callers must pass the same public CLR type across Word, Excel, PowerPoint, HTML, Markdown, PDF, Reader, and host wrappers, create or promote a tiny packageable contract assembly with no dependencies. This is the point where a new package may be cleaner than source-ingested duplicates.

Start with contracts, not engines:

- `OfficeConversionScenario`
- `OfficeConversionProfile`
- `OfficeConversionReport`
- `OfficeConversionFeature`
- `OfficeConversionAsset`
- `OfficeConversionArtifact`
- `OfficeConversionRoundtripExpectation`

Then let existing adapters populate it without moving their format-specific engines or making every converter reference every other converter.

#### Dependency Check

Current answer: no new shared converter dependency is needed yet.

Evidence from the current project graph:

- `OfficeIMO.Shared` is not a project or package. It is source-included by `Directory.Build.props`, and its current shared types are internal implementation helpers. That is safe for private code reuse, but it should not become a public cross-package CLR contract layer.
- PDF-producing adapters already share the correct owner dependency: `OfficeIMO.Word.Pdf`, `OfficeIMO.Markdown.Pdf`, `OfficeIMO.Html.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf`, and `OfficeIMO.Rtf.Pdf` all reference `OfficeIMO.Pdf`. Their public options expose `PdfConversionReport` from that owner, which is acceptable because PDF output already requires the PDF package.
- HTML-consuming adapters already share the correct owner dependency: `OfficeIMO.Word.Html`, `OfficeIMO.Markdown.Html`, and `OfficeIMO.Html.Pdf` reference `OfficeIMO.Html`. HTML diagnostics, resource manifests, profile contracts, and gallery contracts can stay there while the workflow is HTML-specific.
- Reader already owns a separate `OfficeDocumentReadResult` envelope in `OfficeIMO.Reader`, which is appropriate for ingestion/readback and should not force conversion adapters to depend on Reader unless they are reader workflows.
- The Blazor converter host already references the concrete converter packages it needs and can aggregate native reports at the host boundary without requiring all converter packages to share one public CLR type.

So the next step should be schema-first:

1. Define a portable manifest shape for conversion/proof artifacts in docs or tests.
2. Let each owner map its native report into that manifest when producing galleries or host output.
3. Add a real packageable contract assembly only if a public API must accept or return the same `OfficeConversionResult` / `OfficeConversionReport` type across unrelated converter packages.

Until that public type-identity requirement exists, a new shared dependency would add package complexity before it removes real duplication.

### 2. A Format-Neutral Report Vocabulary, Not A Universal Model

The repo currently has several partially-overlapping models:

- `HtmlConversionDocument` for HTML.
- `MarkdownDoc` for Markdown.
- `OfficeMarkupDocument` for semantic authoring.
- `PdfLogicalDocument` for PDF readback.
- Word/Excel/PowerPoint native object models.
- Drawing visual primitives.

These are valid owners. They should not be collapsed into one giant model.

Recommendation: do not create a giant universal object model. Instead, create a small report vocabulary that describes what happened during conversion:

- document, section, heading, paragraph, inline runs, list, table, figure, image, code, quote, callout, form field, note/comment, chart, drawing object, page/slide/sheet region;
- common style tokens: typography, color, spacing, borders, fills, alignment, language, link target, alt/title, source reference;
- stable source anchors and degradation records.

This vocabulary is for diagnostics, manifests, proof galleries, and loss reporting. It should not replace `WordDocument`, `ExcelDocument`, `PowerPointPresentation`, `MarkdownDoc`, `HtmlConversionDocument`, `PdfLogicalDocument`, or `OfficeMarkupDocument`.

### 3. A Shared Visual Theme System Beyond Markdown

`MarkdownVisualTheme` is useful and already bridges Markdown HTML, PDF, and Word. It should not become the accidental theme brain for Excel and PowerPoint unless it is renamed/generalized.

Recommendation:

- Promote the neutral subset into an `OfficeVisualTheme` or `OfficeDocumentTheme` concept.
- Keep `MarkdownVisualTheme` as a Markdown-friendly facade over the neutral theme.
- Map Word styles, Markdown themes, PDF themes, PowerPoint themes, and future HTML profiles through that neutral theme.
- Keep format-specific theme details in adapters when the source format has richer semantics.

### 4. Excel/PowerPoint HTML Lanes

The branch now has dedicated first-class Excel-to-HTML and PowerPoint-to-HTML adapter packages. The remaining work is premium fidelity, not basic package existence.

Recommended profiles:

- `ExcelHtmlProfile.SemanticTables`: workbook/sheet/table/cell values as accessible HTML tables with workbook metadata, formulas/comments/chart/image annotations, and assets through the shared asset manifest.
- `ExcelHtmlProfile.VisualReview`: worksheet/range/page view using Drawing/SVG/positioned HTML, not editable spreadsheet semantics.
- `PowerPointHtmlProfile.SemanticSlides`: slide outlines, notes, tables, image previews, chart snapshot metadata, and alt text as accessible HTML.
- `PowerPointHtmlProfile.VisualReview`: slide pages as positioned-review HTML with directly embedded pictures, shared Drawing SVG chart rendering for supported snapshots, and honest chart placeholders when chart drawing cannot be mapped.

Do not make one "Excel to HTML" mean both editable semantic data and visual screenshot parity.

### 5. PDF-To-Office Reconstruction Honesty

PDF-to-HTML and PDF-table-to-Excel/PowerPoint exist. Full PDF-to-Word/Excel/PowerPoint editable reconstruction is not yet honest as a general claim.

Recommendation:

- Keep PDF-to-Office scoped to named extraction workflows first: tables, text outline, images/assets, forms, annotations.
- Require `PdfLogicalDocument` proof before any editable reconstruction claim.
- Use degradation records instead of silent reconstruction.

### 6. Cross-Converter Visual Proof

PDF has a scenario manifest and raster gallery direction. Excel image export has strong visual baseline gates. HTML Word has an artifact gallery. These need to become one evidence system.

Recommendation:

- One manifest shape across HTML, Word, Markdown, Excel, PowerPoint, PDF, and images.
- `OfficeIMO.Html` now has both human-readable Markdown and deterministic JSON for the HTML capability-gallery manifest. The next step is mapping other owners into the same schema-shaped payload from their native reports, not forcing them to reference a universal converter model.
- Excel and PowerPoint HTML adapters now write semantic HTML, visual-review HTML, and the same Markdown/JSON gallery manifest through thin `SaveHtmlCapabilityGallery` APIs. Their manifests derive expectations from real workbook/deck content, so formulas, comments, charts, images, pictures, text, and placeholders are reported only when the source actually contains them.
- Every supported lane gets at least one rich source fixture with:
  - source feature inventory;
  - expected simplifications;
  - generated artifacts;
  - hashes;
  - extracted text/logical readback;
  - diagnostics;
  - visual baseline or nonblank rendered proof where applicable.

### 7. Adapter-Specific Fidelity Burn-Downs

Known high-value gaps:

- Word/HTML: browser layout, richer CSS, floating/anchored layout, advanced forms/widgets, and visual proof across real documents.
- Markdown/HTML: AST-first extension seam, de-IX generic renderer cleanup, HTML-to-Markdown built on AST rather than parallel conventions.
- HTML/PDF: clearer semantic vs document profiles, CSS/resource subset docs, trusted/untrusted examples, PDF-to-HTML proof.
- PDF: typography, shared layout engine, parser preservation, forms/annotations/redaction, compliance proof.
- Excel: chart fidelity, conditional formatting, exact text metrics, pagination/tiling, headers/footers, grouped objects, broader image effects.
- PowerPoint: master/layout/theme inheritance, grouped transforms, richer text/table/chart fidelity, media/SmartArt fallbacks.

## Recommended Build Order

1. Define the shared conversion/proof contract in the smallest shape that keeps dependencies correct: schema-only first, source-ingested internal helpers second, and a tiny packageable contract assembly only if public CLR type identity is required.
2. Make existing adapters populate the contract without changing output behavior.
3. Generalize `MarkdownVisualTheme` into a neutral theme core, then keep Markdown as a facade.
4. Add Excel/PowerPoint semantic HTML profiles over existing models.
5. Add Excel/PowerPoint visual-review HTML profiles over Drawing/SVG/page surfaces.
6. Expand the cross-converter scenario manifest and visual proof gallery.
7. Only then deepen fidelity in adapters, moving reusable layout/rendering pieces into `OfficeIMO.Pdf` or `OfficeIMO.Drawing`.

## Tenth Slice Completed

This slice connected the rich Excel and PowerPoint HTML lanes to the shared proof system instead of leaving their evidence as separate adapter-specific assertions.

Added:

- `HtmlCapabilityGalleryArtifact.WriteTextFile` as the shared way to write text artifacts and hash them.
- `ExcelHtmlCapabilityGalleryOptions` and `SaveHtmlCapabilityGallery` in `OfficeIMO.Excel.Html`.
- `PowerPointHtmlCapabilityGalleryOptions` and `SaveHtmlCapabilityGallery` in `OfficeIMO.PowerPoint.Html`.
- Rich adapter tests that create workbooks/presentations with formulas, comments, charts, images, pictures, notes, text, visual chart SVG output, and fallback-placeholder diagnostics, then assert parseable manifest JSON under the shared `officeimo.html.capability-gallery` schema.

This does not claim full PowerPoint chart parity or full Excel object parity. It makes those states explicit in the same manifest contract as the rest of the HTML proof pipeline.

## Eleventh Slice Completed

This slice tightened the proof contract so source-specific Office-to-HTML lanes are visible in the same manifest as the shared HTML profile.

Added:

- `HtmlCapabilityGalleryManifest.OfficeProfiles` for source-specific lane contracts.
- Markdown and JSON manifest output for `OfficeHtmlConversionProfileContracts`, including source format, shared profile, visual primitive owner, supported HTML, resource guarantees, and diagnostic guarantees.
- Excel gallery manifests now report both `ExcelSemanticTables` and `ExcelVisualReview`.
- PowerPoint gallery manifests now report both `PowerPointSemanticSlides` and `PowerPointVisualReview`.
- The rich PowerPoint gallery test now includes an actual table so table preservation is covered by the same manifest evidence as pictures, text, notes, charts, and placeholders.

This keeps one shared proof schema while still letting Excel and PowerPoint own their native semantics and lane-specific fidelity claims.

## Twelfth Slice Completed

This slice removes the placeholder-only PowerPoint chart visual proof for supported chart snapshots without creating another chart renderer.

Added:

- `OfficeIMO.PowerPoint.Html` maps `PowerPointChartSnapshot` into the existing `OfficeIMO.Drawing` chart snapshot model.
- PowerPoint visual-review HTML now embeds shared Drawing SVG output for supported charts and marks the visual owner with `data-officeimo-visual-owner="OfficeIMO.Drawing"`.
- The PowerPoint gallery manifest reports `PowerPointChartVisualReviewRendered` when the shared renderer path is used and keeps `PowerPointChartVisualPlaceholder` only for fallback cases.
- Rich PowerPoint adapter tests now require `officeimo-chart-rendered`, SVG output, and the Drawing owner marker instead of accepting a chart placeholder as success.

This improves visual consistency through the existing Drawing owner while staying honest about remaining PowerPoint gaps such as theme inheritance, grouped transforms, animation, media, SmartArt, and full chart parity.

## Thirteenth Slice Completed

This slice strengthens the Excel annotation proof so the rich fixture is no longer only a red comment indicator plus a diagnostic.

Added:

- `OfficeIMO.Excel.Html` now enables the existing Excel/Drawing comment-body approximation when visual-review HTML is exported with default visual options.
- Excel visual-review HTML emits a readable `data-officeimo-visual-proof="comment-callout"` comment section beside the shared Drawing SVG, so comments are available to HTML consumers even when the SVG renderer reports approximation boundaries.
- The Excel capability gallery marks comments as `VisualProof` and emits `ExcelCommentVisualReviewRendered` when the visual proof is present.
- Rich Excel adapter tests now require the visual comment proof marker, comment text, and dependency-free approximation diagnostic across `net472`, `net8.0`, and `net10.0`.

This uses the existing Excel model and Drawing visual snapshot path rather than creating another annotation renderer. It still does not claim Excel-exact threaded-comment editing or native popover parity.

## Fourteenth Slice Completed

This slice brings Word back into the same source-specific HTML profile vocabulary as Excel and PowerPoint instead of leaving the primary editable roundtrip lane described only by the generic shared `Document` profile.

Added:

- `OfficeHtmlConversionProfile` now declares `WordSemanticDocument`, `WordDocumentRoundTrip`, and `WordPrintReview`.
- `OfficeHtmlConversionProfileContracts` now documents Word semantic, editable roundtrip, and print-review lanes with shared profile mappings, supported HTML, resource guarantees, and diagnostic guarantees.
- The existing HTML -> Word -> HTML artifact gallery manifest now records the `WordDocumentRoundTrip` office profile in both Markdown and JSON.
- Tests assert that the Word roundtrip manifest carries `officeProfiles[0].id = WordDocumentRoundTrip`, `sourceFormat = Word`, and `sharedProfile = Document`.

This does not claim browser-layout or print-fidelity parity. It makes the lane explicit and puts its evidence in the same manifest path as the richer Excel and PowerPoint galleries.

## Fifteenth Slice Completed

This slice moves the Word HTML proof path out of test-only code and into the owning adapter package, matching the Excel and PowerPoint gallery shape.

Added:

- `WordHtmlCapabilityGalleryOptions` in `OfficeIMO.Word.Html`.
- `SaveHtmlCapabilityGallery(this string html, ...)` in `OfficeIMO.Word.Html`.
- The Word gallery API writes source HTML, generated DOCX, roundtrip HTML, Markdown manifest, and JSON manifest.
- The generated manifest records `WordDocumentRoundTrip` in `officeProfiles`, source-derived expectations for headings/tables/forms/comments/images when present, and a `WordOpenXmlPackageValid` diagnostic when the DOCX passes OpenXML validation.
- The existing artifact gallery test now consumes the public Word gallery API instead of owning a private manifest writer path.

This improves the "one proof brain" goal: Word, Excel, and PowerPoint now expose adapter-owned gallery APIs that all write through the shared `OfficeIMO.Html` manifest schema.

## Sixteenth Slice Completed

This slice upgrades the Word proof from a basic document sample to a richer roundtrip fixture with visual evidence.

Added:

- The Word gallery test now includes table sections, checkboxes, text/select form controls, an embedded data-URI image, and skipped HTML comments as diagnostic evidence.
- The test requires image evidence in the generated roundtrip HTML and records image preservation in the shared capability manifest expectations.
- The visual proof generator now emits a Word roundtrip gallery beside the Excel and PowerPoint galleries using the public `OfficeIMO.Word.Html` API.
- Generated proof artifacts include the source HTML, generated DOCX, roundtrip HTML, Markdown manifest, JSON manifest, and Playwright screenshots for source and roundtrip rendering.
- Focused validation passes across `net472`, `net8.0`, and `net10.0` for the Word gallery and profile-contract tests.

This still does not claim Word-perfect layout parity. The manifest score and diagnostics intentionally expose losses such as figure-signature, grid, and form-state fidelity instead of hiding them behind a successful file write.

## Seventeenth Slice Completed

This slice fixes a real Markdown image roundtrip blocker found by the broader rich-content validation pass.

Added:

- The shared Markdown reader now preserves Windows drive paths such as `C:\Support\GitHub\_worktrees\...\OfficeIMO.png` as local paths instead of treating `C:` like a URI scheme.
- Markdown destination unescaping now keeps backslashes for Windows drive paths, so `\_` in a directory name is not collapsed into `_`.
- `DisallowFileUrls` still blocks Windows local paths when that security policy is requested.
- A regression test covers both preserving and blocking Windows local image destinations.
- Markdown-to-Word image tests and visual Markdown fixture tests now pass across `net472`, `net8.0`, and `net10.0`.

This belongs in the shared Markdown reader rather than the Word adapter because every Markdown consumer needs the same local-path policy. Word remains the thin consumer that decides whether local images are allowed through `MarkdownToWordOptions.AllowLocalImages`.

## Eighteenth Slice Completed

This slice records the current evidence boundary after the richer HTML/Markdown validation pass.

Verified:

- The focused Markdown image-path regression, HTML saved-image-path regression, and natural-size Markdown-to-Word image test pass across `net472`, `net8.0`, and `net10.0`.
- The broader `FullyQualifiedName~Html|FullyQualifiedName~Markdown` test lane passes 4,974 tests per target framework with 4 existing skips.
- The only failure in that broader lane is `PdfDocumentRasterVisualBaselineTests.MarkdownTechnicalDocument_MatchesPopplerRasterBaseline` on each target framework.
- The PDF raster delta is large enough to stay as an open visual-baseline decision rather than being treated as noise: 113,190 different pixels out of 484,704, max channel delta 240, allowed different pixels 32.

This means the Windows local-image roundtrip blocker is fixed, but the PDF visual contract is not closed. The next PDF step is either approving and refreshing the intended baseline after visual review or fixing the renderer/theme change that caused the delta.

## Nineteenth Slice Completed

This slice closes the stale-base and stale-proof gaps that were hiding behind the previous validation result.

Completed:

- Fast-forwarded the worktree onto `origin/codex/excel-image-premium-drawing-consolidation`, including the latest shared Drawing and Excel image-rendering fixes.
- Reviewed the failing `MarkdownTechnicalDocument_MatchesPopplerRasterBaseline` artifact pair. The generated PDF raster represented the new shared Markdown technical-document theme, so the approved raster baseline was refreshed instead of weakening the comparison.
- Rebuilt and reran focused rich HTML/PDF checks across `net472`, `net8.0`, and `net10.0`: 10 passed on each target framework.
- Reran the broad `FullyQualifiedName~Html|FullyQualifiedName~Markdown` lane after the baseline refresh: 4,975 passed and 4 skipped on each target framework, with 0 failures.
- Regenerated the rich HTML visual proof artifacts from the fast-forwarded branch and captured fresh Playwright screenshots for Excel semantic, Excel visual review, PowerPoint semantic, PowerPoint visual review, Word source HTML, and Word roundtrip HTML.
- Updated the proof generator so Excel and PowerPoint use their adapter-owned `SaveHtmlCapabilityGallery` APIs, matching Word and writing shared Markdown/JSON manifests instead of only ad-hoc HTML files.

Current generated manifest proof:

- `excel-rich.manifest.json` declares `ExcelSemanticTables` and `ExcelVisualReview`, with formulas and chart data preserved, comments represented with explicit visual proof, and images represented through semantic data URIs plus `OfficeIMO.Drawing` visual proof.
- `powerpoint-rich.manifest.json` declares `PowerPointSemanticSlides` and `PowerPointVisualReview`, with text boxes and tables preserved and pictures/charts represented as visual proof through `OfficeIMO.Drawing`.
- `word-roundtrip.manifest.json` declares `WordDocumentRoundTrip`, with headings, tables, table sections, form controls, images, and DOCX package validity preserved, while skipped HTML comments are reported honestly.

This closes the stale visual evidence problem. It does not close the larger end-to-end goal: Excel and PowerPoint currently have export/visual-review proof lanes, not full editable HTML import back into native Excel/PowerPoint packages.

## Twentieth Slice Completed

This slice adds the first editable semantic HTML import foundation for Excel and PowerPoint instead of stopping at rich export screenshots.

Added:

- `ExcelHtmlLoadOptions`, `ExcelHtmlLoadResult`, and `LoadExcelFromHtmlWithResult` in `OfficeIMO.Excel.Html`.
- `PowerPointHtmlLoadOptions`, `PowerPointHtmlLoadResult`, and `LoadPowerPointFromHtmlWithResult` in `OfficeIMO.PowerPoint.Html`.
- Both importers reuse `OfficeIMO.Html` parsing and existing adapter models; no new parser package or cross-format interchange package was added.
- Excel semantic import rebuilds worksheet tables, formulas, comments, embedded data-URI images, and native charts from semantic chart data when present.
- PowerPoint semantic import rebuilds slide text, tables, embedded data-URI pictures, speaker notes, and native chart inventory.
- Import result objects expose counts and diagnostics so callers can see exactly which semantic lanes were restored and which were approximate.
- Focused rich adapter tests now prove Excel and PowerPoint can export rich semantic HTML, import it back into native package models, and export the imported package back to semantic HTML.

Verified:

- `ExcelHtml_LoadsSemanticRichWorkbookBackToNativeWorkbook` asserts native Excel values, formula cells, comments, images, charts, and second-pass semantic HTML.
- `PowerPointHtml_LoadsSemanticRichPresentationBackToNativePresentation` asserts native PowerPoint text boxes, table cells, pictures, charts, speaker notes, and second-pass semantic HTML.
- The focused adapter lane passes across `net472`, `net8.0`, and `net10.0`: 4 passed on each target framework.

Current honest limitation:

- Excel semantic HTML in this slice can rebuild native chart categories, series names, and values when exported by the current adapter. Older or minimal semantic HTML without chart data still falls back to chart creation from the restored sheet range when a usable range is available.
- PowerPoint semantic HTML in this slice can rebuild native chart categories, series names, and values when exported by the current adapter. Older or minimal semantic HTML without chart data still falls back to reconstructed placeholder values and records a diagnostic.

This closes the "basic table only" proof gap for Excel and PowerPoint. It does not claim perfect Office fidelity yet; it establishes adapter-owned import lanes, shared diagnostics, and tests that expose the remaining rich-data gaps instead of pretending screenshots are editable roundtrips.

## Twenty-First Slice Completed

This slice closes the PowerPoint chart-data placeholder gap for the current semantic HTML lane.

Added:

- `OfficeIMO.PowerPoint.Html` now emits an accessible `officeimo-chart-data` table under each semantic chart inventory item when `PowerPointChartSnapshot` is available.
- The table carries chart categories, series names, and numeric values as inspectable HTML instead of hidden JSON or a new interchange model.
- `LoadPowerPointFromHtmlWithResult` now reads that table and rebuilds the native chart with the original semantic chart data.
- The placeholder chart-data diagnostic is now only used when semantic chart data is absent or invalid.

Verified:

- `PowerPointHtml_ExportsSemanticSlidesWithExtractionProof` requires the chart data table and a real value from the source chart.
- `PowerPointHtml_LoadsSemanticRichPresentationBackToNativePresentation` asserts imported native chart categories, series name, and values through `TryGetSnapshot`.
- The focused PowerPoint adapter lane passes across `net472`, `net8.0`, and `net10.0`: 3 passed on each target framework.

This keeps ownership in the PowerPoint HTML adapter and consumes the existing PowerPoint chart snapshot. It avoids adding a new `.XXX` package, avoids a second chart brain, and moves the chart path from screenshot/inventory proof to editable semantic roundtrip proof for supported chart snapshots.

## Twenty-Second Slice Completed

This slice closes the same chart-data gap for the Excel semantic HTML lane.

Added:

- `OfficeIMO.Excel.Html` now emits an accessible `officeimo-chart-data` table under each semantic chart inventory item when `ExcelChartSnapshot` is available.
- The table carries chart categories, series names, and numeric values as inspectable HTML instead of hidden JSON or a new interchange model.
- `LoadExcelFromHtmlWithResult` now reads that table and rebuilds native Excel charts with the original semantic chart data.
- Excel gallery manifests now mark charts as `Preserved` when semantic chart data is present and emit `ExcelChartSemanticDataPreserved`.

Verified:

- `ExcelHtml_ExportsSemanticWorksheetRichContent` requires the chart data table and source chart category/series labels.
- `ExcelHtml_LoadsSemanticRichWorkbookBackToNativeWorkbook` asserts imported native chart categories, series name, and values through `TryGetSnapshot`.
- `ExcelHtml_CapabilityGalleryWritesSharedManifestForRichWorkbook` asserts the preserved chart expectation and manifest diagnostic.
- The focused rich adapter lane passes across `net472`, `net8.0`, and `net10.0`: 5 passed on each target framework.

This keeps ownership in the Excel HTML adapter and consumes the existing Excel chart snapshot/data APIs. It avoids adding a new `.XXX` package, avoids a second chart brain, and moves Excel chart proof from inventory/range reconstruction toward editable semantic roundtrip proof for supported chart snapshots.

## Twenty-Third Slice Completed

This slice removes the last stale Word proof contradiction found during the rich-content audit.

Fixed:

- `WordTableCell` now normalizes `w:tcPr` child element order after table-cell property mutations such as width, merge, borders, shading, margins, text direction, fit text, and vertical alignment.
- The horizontal merge-to-gridSpan conversion path now routes the restart cell through the same ordering guard instead of inserting `w:gridSpan` at the front of `w:tcPr`.
- The Word HTML gallery resource manifest now uses the trusted OfficeIMO URL policy by default, matching the gallery's trusted document import profile and allowing embedded data-URI image proof.
- `HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml` now asserts that the generated DOCX has `WordOpenXmlPackageValid`, has no `WordOpenXmlValidationError`, has no `ImageResourceRejectedByPolicy`, and reports 1 allowed / 0 blocked resources for the embedded image fixture.

Verified:

- `HtmlArtifactGallery_GeneratesValidDocxAndRoundTripHtml` passes across `net472`, `net8.0`, and `net10.0`.
- The focused 40-test HTML/adapter/theme/PDF routing evidence lane passes across `net472`, `net8.0`, and `net10.0`.
- The focused 7-test PDF-to-HTML / HTML-to-PDF profile lane passes across `net472`, `net8.0`, and `net10.0`.
- `dotnet run --project output\html-visual-proof\Generator\Generator.csproj` regenerated the rich gallery artifacts.
- Playwright refreshed the six rich screenshots from the regenerated HTML: Word source, Word roundtrip, Excel semantic, Excel visual review, PowerPoint semantic, and PowerPoint visual review.
- Pixel sampling confirmed the refreshed screenshots are non-blank 1280x900 PNGs.

Current Word manifest proof:

- `word-roundtrip.manifest.md` reports `WordOpenXmlPackageValid`.
- `word-roundtrip.manifest.md` reports `Allowed: 1` and `Blocked: 0` for the embedded data-URI image.
- `word-roundtrip.manifest.md` no longer reports `WordOpenXmlValidationError` or `ImageResourceRejectedByPolicy`.

This keeps the Word lane honest: a preserved DOCX package expectation now corresponds to real OpenXML validation evidence, and image preservation no longer conflicts with the resource manifest.

## Guardrail

The clean target shape is:

```text
source format
  -> source-owned model and semantics
  -> shared conversion/proof contract
  -> shared assets, diagnostics, theme, and visual primitives
  -> target-owned adapter
  -> artifact plus proof report
```

Avoid:

```text
source format
  -> private converter-specific model
  -> private diagnostics
  -> private theme rules
  -> private renderer
```

The second path is how the repo gets two brains again.

## Bottom Line

OfficeIMO is missing less engine than policy. The engines are mostly in the right owners already. To make the roundtrips feel premium and consistent, add one shared conversion/proof contract, one neutral visual theme layer, and explicit semantic vs visual profiles for HTML/PDF/Office conversions. Then keep all visual and layout reuse flowing into `OfficeIMO.Drawing` and `OfficeIMO.Pdf`, while Word, Excel, PowerPoint, Markdown, HTML, Reader, and Markup stay as thin source/target adapters with honest diagnostics.
