# OfficeIMO Markup Prototype Plan

This checklist is the working plan for the unified OfficeIMO Markup layer and editor prototype. Keep every entry as a checkbox and update it when the matching implementation or verification is actually done.

## Locked Design

- [x] Use YAML front matter for document-level metadata such as `profile`, `title`, `author`, `theme`, and slide size.
- [x] Support `profile: presentation`, `profile: document`, `profile: workbook`, and `profile: common`.
- [x] Keep regular Markdown as the shared authoring core for headings, paragraphs, lists, code blocks, images, and tables.
- [x] Use `@...` directives for major Office containers such as `@slide`, `@section`, and `@sheet`.
- [x] Use `::...` directives for Office-aware blocks such as `::notes`, `::textbox`, `::image`, `::chart`, `::mermaid`, `::range`, and `::formula`.
- [x] Treat `---` as a slide separator only after front matter in the presentation profile.
- [x] Keep the AST semantic and independent from C#, PowerShell, or direct OfficeIMO API calls.
- [x] Keep C# and PowerShell generation as separate emitters.

## Core Library

- [x] Create `OfficeIMO.Markup` with a shared semantic AST.
- [x] Create `OfficeIMO.Markup.PowerPoint` for real PowerPoint file export.
- [x] Create `OfficeIMO.Markup.Excel` for real Excel workbook export.
- [x] Create `OfficeIMO.Markup.Word` for real Word document export.
- [x] Add parser support for fenced `officeimo` compatibility blocks.
- [x] Add parser support for YAML front matter metadata.
- [x] Add parser support for `@slide` and presentation slide separators.
- [x] Add parser support for initial `::` block directives.
- [x] Add validation for profile-specific constructs.
- [x] Add initial C# emitter.
- [x] Add initial PowerShell emitter.
- [x] Add initial real `.pptx` exporter for the presentation profile.
- [x] Add initial real `.docx` exporter for the document profile.
- [x] Add native PowerPoint chart export for inline CSV chart data.
- [x] Add JSON preview/export DTOs as a stable tooling contract.
- [x] Add more precise AST nodes for layout blocks such as textbox, columns, column, and card.
- [x] Add a reusable placement model shared by positioned blocks.
- [x] Add normalized theme/style resolution for semantic layout nodes.
- [x] Add Mermaid render/export integration.
- [x] Add bundled/editor-friendly Mermaid renderer acquisition.
- [x] Add CLI fallback discovery for the editor-installed Mermaid renderer.
- [x] Integrate PowerPoint designer composition helpers for suitable markup slides.
- [x] Add workbook/range-backed chart export integration.
- [x] Add sheet-qualified Excel table chart sources so dashboards can chart data tables from other sheets.
- [x] Support sheet-qualified workbook targets for ranges, formulas, formatting, tables, and chart placement.
- [x] Add safe workbook export polish for auto-fit columns, hidden gridlines, frozen table headers, styled Markdown headings, and chart series colors.
- [x] Fix Excel autofit so formula-only sheets do not emit invalid empty `<cols />` records.
- [x] Fix Excel chart series shape-property ordering so styled series still validate as OpenXML.
- [x] Map workbook chart semantic attributes for axis titles, number formats, legends, data labels, and gridlines into native Excel charts.
- [x] Preserve workbook chart semantic attributes in emitted C# and PowerShell escape-hatch output.
- [x] Preserve sheet-qualified workbook targets in emitted C# and PowerShell escape-hatch output.
- [x] Preserve combined Excel cell style state when markup/export applies number format, fill, font color, bold, alignment, and wrap settings to the same cell.
- [x] Support workbook `::format` font color attributes in real `.xlsx` export.
- [x] Support workbook `::format` alignment and wrap attributes in real `.xlsx` export.
- [x] Support workbook `::format` italic and underline attributes in real `.xlsx` export.
- [x] Support workbook `::format` vertical alignment and simple all-sides border attributes in real `.xlsx` export.
- [x] Emit concrete workbook `::format` worksheet calls in C# and PowerShell escape-hatch output.
- [x] Add Word document export for headings, paragraphs, lists, tables, sections, headers, footers, TOC, page breaks, images, and inline charts.
- [x] Preserve semantic layout, placement, notes, transitions, formulas, ranges, and chart data in C# emitter output.
- [x] Resolve presentation transition directives with directional variants into native PowerPoint transitions where the core library supports them.
- [x] Expose structured transition metadata to tooling/preview contracts and emitted-code comments.
- [x] Apply presentation transition timing, speed, and advance metadata to native PowerPoint export plus emitted C# and PowerShell escape-hatch code.
- [x] Preserve semantic layout, placement, notes, ranges, formulas, and concrete inline chart data in PowerShell emitter output.
- [x] Add default PowerPoint chart panels, chart typography, legend, gridline, and series color styling.
- [x] Map presentation chart semantic attributes for axis titles, number formats, legends, data labels, and gridlines into native PowerPoint charts.
- [x] Add PowerPoint text autofit for styled markup text boxes to reduce clipping in generated decks.
- [x] Add explicit `fit=contain` and `fit=stretch` behavior for PowerPoint image and Mermaid diagram exports.
- [x] Add a branded PowerPoint fallback canvas for explicit-position slides and summary cards for simple Markdown bullet slides.
- [x] Add PowerPoint visual panels for rendered Mermaid diagrams, plus opt-in image panels.
- [x] Support presentation background image and overlay directives in real `.pptx` export.
- [x] Support native PowerPoint linear-gradient slide backgrounds from semantic `background: gradient(...)` directives.
- [x] Support semantic gradient angle directives such as `background: gradient(primary, accent1) angle=45` in PowerPoint export and VS Code preview.
- [x] Skip branded PowerPoint fallback canvas chrome when a slide declares its own explicit background.
- [x] Compose semantic `layout: two-columns` presentation slides into themed designer-style column panels when the slide stays in semantic Markdown content.
- [x] Preserve presentation slide `section:` metadata through the semantic AST, emitters, preview contract, and native PowerPoint section export.

## CLI

- [x] Create `OfficeIMO.Markup.Cli`.
- [x] Add `parse` command with JSON output.
- [x] Add `validate` command with JSON diagnostics.
- [x] Add `preview` command with a preview-friendly JSON contract.
- [x] Add `emit --target csharp`.
- [x] Add `emit --target powershell`.
- [x] Add `export --target pptx --output <file.pptx>`.
- [x] Add `export --target xlsx --output <file.xlsx>`.
- [x] Add `export --target docx --output <file.docx>`.
- [x] Support stdin input for editor integrations.
- [x] Resolve relative export assets such as presentation background images against the markup source directory.
- [x] Add the CLI project to `OfficeIMO.sln`.

## VS Code Extension

- [x] Create `OfficeIMO.Markup.VSCode`.
- [x] Add language registration for `.office.md` and `.omd`.
- [x] Add TextMate grammar for Markdown plus OfficeIMO directives.
- [x] Add snippets for front matter, slides, notes, charts, Mermaid, document sections, and workbook sheets.
- [x] Add diagnostics by calling the CLI.
- [x] Add preview webview for presentation slides.
- [x] Add auto-refresh for open preview panels.
- [x] Add lightweight chart rendering in the live preview.
- [x] Show chart semantic metadata in the live preview, including source, axis titles, formats, legend, labels, and gridlines.
- [x] Add multi-series grouped chart rendering in the live preview for inline CSV charts.
- [x] Show source-backed workbook chart placeholders with source kind, anchor cell, and intended size in the live preview.
- [x] Add Mermaid rendering in the live preview with raw-source fallback.
- [x] Keep VS Code presentation preview faithful to authored slide content instead of injecting generated summary/debug text.
- [x] Align VS Code blank-slide preview title behavior with PowerPoint export to avoid duplicate authored titles.
- [x] Add document-profile live preview for pages, sections, TOC, headers/footers, tables, and page breaks.
- [x] Add workbook-profile live preview for sheets, ranges, named tables, formulas, and charts.
- [x] Apply workbook `::format` cell styling to the live preview grid so alignment, emphasis, fill, text color, wrap, and number-format hints are visible while authoring.
- [x] Apply workbook `::format` vertical alignment and simple border styling to the live preview grid.
- [x] Add commands to generate C# and PowerShell.
- [x] Add command to export the current presentation markup to PowerPoint.
- [x] Add command to export the current workbook markup to Excel.
- [x] Add command to export the current document markup to Word.
- [x] Add profile-aware `Export Office Document` command that chooses PowerPoint, Word, or Excel from the markup profile.
- [x] Add profile-aware `Export and Open Office Document` command that exports the right Office file and launches it.
- [x] Add `Generate C# File` and `Generate PowerShell File` commands that save emitted code beside the markup source.
- [x] Add `Generate Artifacts` command that emits C#, PowerShell, and the profile-selected Office file in one pass.
- [x] Add configurable output-directory defaults so generated code and Office files can land in a `generated` subfolder instead of beside the markup source.
- [x] Add preview-surface actions for refresh, validate, generate artifacts, and export/open so the webview becomes an authoring control surface.
- [x] Add output-location visibility plus an `Open Output Folder` action so the preview and command surface make generated file placement explicit.
- [x] Add `Open Generated C#` and `Open Generated PowerShell` commands plus smarter artifact notifications for faster code round-tripping.
- [x] Add command to install an extension-local Mermaid renderer.
- [x] Add configuration for CLI path and debounce delay.
- [x] Add workspace-first CLI discovery for packaged VS Code extension installs.
- [x] Bundle the CLI with packaged VS Code extension installs so previews work outside the source repository.
- [x] Add extension project scripts patterned after ForgeFlow.
- [x] Map validation diagnostics to the offending markup source line in VS Code.
- [x] Add editor title, editor context, and explorer context menus for `.md`, `.office.md`, and `.omd` preview/export actions.
- [x] Limit OfficeIMO context menus to `.omd` and `.office.md` so plain `.md` keeps the built-in Markdown preview menu.
- [x] Bundle a local browser Mermaid renderer into the VSIX for live preview without CDN/runtime network dependencies.
- [x] Fix VS Code Mermaid preview so the render target remains measurable and falls back to a simple SVG flowchart when Mermaid cannot render.
- [x] Force VS Code presentation preview to render one slide per row so wide editor panes do not squeeze slides side by side.
- [x] Add export success actions to open or reveal generated Office files directly from VS Code.
- [x] Honor explicit slide backgrounds in the VS Code presentation preview, including local image assets and overlays.
- [x] Resolve presentation background theme aliases such as `primary`, `accent1`, `accent2`, `background`, and `surface` in the VS Code preview.

## Developer Scripts

- [x] Add extension `build.ps1`.
- [x] Add `scripts/dev-install.ps1`.
- [x] Add `scripts/install-insiders.ps1`.
- [x] Add `scripts/auto-install-insiders.ps1`.
- [x] Add shell variants for non-Windows contributors.

## Verification

- [x] Build `OfficeIMO.Markup` for `net8.0`.
- [x] Test markup parser unit tests.
- [x] Build `OfficeIMO.Markup.Cli`.
- [x] Smoke test CLI `parse`.
- [x] Smoke test CLI `emit --target csharp`.
- [x] Smoke test CLI `export --target pptx`.
- [x] Verify generated sample `.pptx` contains native chart parts and embedded workbook data.
- [x] Verify generated sample `.xlsx` contains native worksheet, table, drawing, and chart parts.
- [x] Verify generated sample `.docx` contains native document, header/footer, TOC, table, page break, and chart parts.
- [x] Add unit coverage that exports and reopens a real `.pptx`.
- [x] Add unit coverage that exports and validates a real `.xlsx`.
- [x] Add unit coverage that exports, validates, and reopens a real `.docx`.
- [x] Compile VS Code extension TypeScript.
- [x] Package VS Code extension VSIX.
- [x] Smoke test VS Code extension install scripts in VS Code Insiders.
- [x] Merge current `master` PowerPoint designer improvements into the markup branch.
- [x] Verify generated sample `.pptx` contains rendered Mermaid image media instead of raw diagram source.
- [x] Verify generated sample `.pptx` opens through PowerPoint COM and exports slide PNG previews.
- [x] Add regression coverage for PowerPoint summary cards and branded fallback canvas shapes.
- [x] Add regression coverage that configured Mermaid rendering inserts PNG media into `.pptx` output instead of raw diagram text.
- [x] Add regression coverage for relative PowerPoint background-image export with overlay and fallback-canvas suppression.
- [x] Add regression coverage for native PowerPoint gradient background export in both the core slide API and markup exporter.
- [x] Add regression coverage for semantic gradient angle export and directional PowerPoint transition mapping.
- [x] Add regression coverage for native PowerPoint transition timing/speed round-trip and markup-driven transition timing export.
- [x] Add regression coverage for semantic two-column PowerPoint slide composition without fallback canvas chrome.
- [x] Add regression coverage for presentation section metadata through parser, emitters, and native PowerPoint section output.
- [x] Verify generated sample `.xlsx` opens through Excel COM without repair prompts.
- [x] Add regression coverage for valid Excel chart series styling and sheet-qualified dashboard charts.
- [x] Add regression coverage for composed Excel cell styles and workbook `::format` font-color export.
- [x] Add regression coverage for workbook `::format` italic and underline export/emitter support.
- [x] Add regression coverage for workbook `::format` vertical alignment and border export/emitter support.
- [x] Route workbook markup `.xlsx` export through the existing Excel save-time preflight, defined-name repair, and Open XML validation options.
- [x] Add a Windows-only Excel COM smoke regression for markup-exported workbooks when local Excel automation is available.
- [x] Expose workbook export hardening switches through the markup CLI while keeping the strict path enabled by default.
- [x] Repackage and reinstall the VS Code Insiders extension after preview and CLI-discovery improvements.
- [x] Package VS Code extension with bundled license and no VSIX license warning.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.1` after the context menu and preview fidelity fixes.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.2` after the Mermaid preview collapse fix.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.3` after menu routing and profile-aware export fixes.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.4` after preview menu placement and keybinding fixes.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.5` after export workflow polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.6` after saved code generation polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.7` after unified artifact generation polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.8` after configurable output-directory polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.9` after preview action-bar polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.10` after output-folder visibility polish.
- [x] Repackage and reinstall the VS Code Insiders extension as version `0.1.11` after generated-code navigation polish.
