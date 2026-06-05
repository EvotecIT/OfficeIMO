# OfficeIMO PDF Implementation Review - 2026-06-05

## Scope

Reviewed the first-party PDF implementation on `origin/master` at `fd15131e` from the isolated worktree `C:\Support\GitHub\OfficeIMO-pdf-implementation-review-20260605`.

The review focused on current PDF creation/conversion surfaces, especially the new PowerPoint-to-PDF adapter, and on what blocks OfficeIMO from becoming a premium, visually consistent PDF library that can compete with QuestPDF/iText-style layout and conversion scenarios.

## Current Strengths

- The core `OfficeIMO.Pdf` package remains dependency-light and is now the shared engine for Word, Excel, Markdown, and PowerPoint PDF export.
- The support matrix shows broad first-party progress: flow layout, tables, headers/footers, forms groundwork, outlines/bookmarks, attachments/e-invoice groundwork, native Word export, Excel export, Markdown export, and PowerPoint export.
- Visual/raster baselines already exist for core PDF scenarios plus native Word, Excel, and Markdown conversion fixtures.
- The adapter direction is right: source converters target reusable PDF primitives instead of each adapter hand-rolling drawing, table, image, and chart behavior.

## Findings

### P1 - PowerPoint PDF needs raster visual gates

`OfficeIMO.PowerPoint.Pdf` is now a fixed-layout slide exporter, but the existing PowerPoint tests are mostly structural/content checks. The Poppler raster baseline lane covers core, Word, Excel, and Markdown scenarios, but not PowerPoint. For a slide-to-PDF exporter, text placement, z-order, clipping, table layout, chart placement, shape rendering, and background composition are the product contract.

Add a PowerPoint raster baseline scenario that renders a representative slide deck with:

- slide background color/image/gradient,
- text boxes with mixed alignment, rich runs, bullets/numbering, hyperlinks, and vertical anchoring,
- pictures with crop/fit behavior,
- fixed-position tables with merged cells and borders,
- charts through the shared vector renderer,
- rectangle/rounded rectangle/ellipse/line shapes,
- grouped/inherited layout content and off-slide clipping.

### P1 - PowerPoint text can be silently clipped

The mixed-paragraph/list path in `RenderParagraphTextBox` estimates paragraph height with a fixed `fontSize * 0.52` average character width, then clamps each paragraph height to the remaining text box height and stops when the cursor reaches the text box bottom. There is no warning when paragraphs are clipped or skipped.

This is risky for premium conversions: a slide can look "successfully exported" while text disappears. The exporter should use the shared PDF text measurement/wrapping path, and it should emit a `PowerPointPdfSaveOptions.Warnings` entry when text is truncated, clipped, or unsupported layout data forces an approximation.

### P2 - PowerPoint PDF is missing from the dependency-light guardrail

`OfficeIMO.Tests/PackageDependencyGuardrails.cs` protects `OfficeIMO.Pdf`, `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, and `OfficeIMO.Markdown.Pdf`, but not `OfficeIMO.PowerPoint.Pdf`. The PowerPoint adapter currently has no package references, so this is cheap to lock down before accidental dependencies creep in.

### P2 - PowerPoint table cell rich runs are flattened

PowerPoint table cell text walks OpenXML runs, but each run is converted with cell-level font/style/color settings. Run-level bold/italic/font/color/size inside table cells is not preserved. This is visible in real slides with partially emphasized table cells and should either be implemented or warned as simplified export.

### P2 - Conversion parity blockers remain large

The support matrix correctly marks several important areas as partial or planned:

- font shaping, glyph fallback, OpenType/CFF embedding, and subsetting,
- formal PDF/A/PDF/UA/Factur-X/ZUGFeRD validation,
- encryption/signature/redaction and safe complex rewrite flows,
- Excel fit-to-height, automatic pagination/scaling, richer image placement, chart fidelity, and locale-specific formats,
- Markdown fully paginated panel containers and deeper theme/layout controls,
- PowerPoint SmartArt, media, theme color resolution, richer shape geometry, richer text layout, richer chart/theme/table fidelity, and full slide fidelity.

These are the right roadmap items. The premium requirement is to turn them into staged contract surfaces with visual and validator proof, not to claim broad support before the evidence exists.

### P3 - Several PDF files are above the structure threshold

Current line counts indicate split pressure before the next large behavior wave:

- `OfficeIMO.PowerPoint.Pdf/PowerPointPdfConverterExtensions.cs` is over 1200 lines.
- `OfficeIMO.Pdf/Rendering/PdfWriter.cs` is over 1100 lines.
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Layout.Context.Process.Row.cs` is over 1000 lines.
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Text.cs` is over 1000 lines.
- `OfficeIMO.Pdf/Rendering/Writer/PdfWriter.Drawing.cs` is over 800 lines.

Before adding more premium layout features, split by semantic ownership: PowerPoint backgrounds, text, tables, charts, shapes/groups, bounds/warnings, and core writer text/table/drawing pipelines.

## Recommended Next Steps

### Phase 0 - Immediate hardening

1. Add `OfficeIMO.PowerPoint.Pdf/OfficeIMO.PowerPoint.Pdf.csproj` to `DependencyLightProjects_HaveNoPackageReferences`.
2. Add the first PowerPoint Poppler raster baseline fixture.
3. Add warnings for clipped/truncated PowerPoint text boxes and unsupported table-cell run styling.
4. Replace PowerPoint paragraph height estimation with shared text measurement/wrapping where possible.
5. Add focused tests for PowerPoint table run formatting, list indentation, overflow warnings, and no silent drop behavior.

### Phase 1 - Visual consistency engine

1. Introduce a shared layout diagnostics model for overflow, clipping, unsupported transforms, skipped content, and bounds corrections.
2. Centralize text measurement, wrapping, line-breaking, baseline placement, vertical alignment, and clipping decisions across flow paragraphs, canvas text boxes, tables, and converters.
3. Make every converter expose a compact visual quality report that can be used by tests and callers.
4. Expand raster baselines from "representative fixture" to scenario families: reports, invoices, dashboards, forms, slides, spreadsheets, technical docs, and mixed-media docs.

### Phase 2 - Conversion fidelity

1. Word: improve floats, anchored shapes, complex tables, equations, SmartArt fallbacks, field evaluation, and tracked revision handling.
2. Excel: implement fit-to-height, automatic page scaling, real page-break pagination, richer cell formats, chart fidelity, worksheet drawing placement, and locale-aware number/date rendering.
3. Markdown/HTML: add robust pagination for panels/callouts/code blocks, CSS-like spacing controls, repeated table headers, and better page-break policy.
4. PowerPoint: implement theme resolution, slide master/layout inheritance fidelity, group transforms, table theme fidelity, SmartArt/media fallbacks, and richer text layout.

### Phase 3 - Premium PDF authoring parity

Add high-value authoring primitives that users expect from a premium PDF engine:

- sections, columns, keep-together/keep-with-next, widows/orphans, running elements, page templates,
- floating/anchored boxes, absolute layers, z-index, clipping, transforms, transparency, blend modes where practical,
- footnotes/endnotes, outlines/bookmarks, named destinations, annotations, links, attachments,
- reusable styles/themes, table styles, chart styles, document templates,
- tagged structure, reading order, alt text, artifact handling, and accessibility diagnostics.

### Phase 4 - iText/PSWritePDF-style manipulation

Continue separating creation from parser/rewrite promises. For manipulation parity, target:

- object streams, xref streams, incremental updates, encryption/decryption, signatures inspection,
- safe redaction, page insertion/extraction/reordering, metadata/catalog preservation,
- forms appearance regeneration and flattening for broader field types,
- conservative rewrite blockers when fidelity cannot be preserved.

### Phase 5 - Compliance proof

Do not claim PDF/A, PDF/UA, Factur-X, or ZUGFeRD support until the generated files pass external validator lanes. Keep using readiness diagnostics, but gate profile claims on veraPDF/Mustang or equivalent CI evidence.

## Validation Performed

- `dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~PackageDependencyGuardrailTests" --no-restore --verbosity minimal`
  - Passed: 21, Failed: 0.
- `dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.Pdf.PowerPointSaveAsPdfTests" --no-restore --verbosity minimal`
  - Passed: 55, Failed: 0.

Both commands build the test project and emit existing nullable warnings, mostly in Excel tests. No test failures were observed in these focused slices.

## Phase 3/4 Local Progress

- Added frame-aware tab leader clamping in the shared PDF text layout path so right/center/decimal tab stops cannot escape columns, table cells, or other text frames.
- Embedded Word document default fonts when available and embedded Arial in the core visual gates that previously relied on viewer-substituted Helvetica.
- Improved PowerPoint image handling: slide background images use proportional cover fitting, uncropped pictures default to proportional contain fitting, and explicit stretch records `picture-aspect-distortion` diagnostics when the frame would visibly distort the source image.
- Added `Docs/reviews/officeimo.pdf-visual-review-gallery-2026-06-05.md` as the manual review index for the generated PDF pack.
