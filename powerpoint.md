# PowerPoint Parity Plan

Last updated: 2026-06-18. Goal: make OfficeIMO.PowerPoint generate PPTX files that are structurally close to the PowerPoint-generated templates in `Assets/PowerPointTemplates` (blank, title, tables+charts), preserve advanced package metadata during normal Open XML round trips, and ship a no-repair experience across net472/net8/net10 while keeping code split into focused partial/typed units that mirror the OpenXml SDK object model.

## Current Findings
- PowerPoint templates include creation IDs on shapes/placeholders, full layout sets, SaveSubsetFonts, and language tags (pl-PL/en-US). Shape/text creation still needs deeper creation ID and placeholder parity, but table cell text now carries language-aware run and end-paragraph defaults.
- Table generation now emits tableStyleId, first-row/banded-row flags, a16:colId/rowId extLst entries, proportional widths/heights, and a public styling surface for fills, borders, alignment, padding, banding flags, and typed object binding.
- Template tableStyles.xml is bundled as an embedded resource and loaded into generated packages instead of writing a bare TableStyleList.
- Chart packaging now creates chart parts under `ppt/charts`, companion style/color parts, and embedded workbooks named `Microsoft_Excel_Worksheet*.xlsx` through deterministic indexed part URIs. Chart authoring still needs deeper template-schema parity and broader series/category editing coverage.
- PowerPoint now has `InspectFeatures()` parity with Word/Excel feature reports, so callers can distinguish editable, partially editable, preserved, and unsupported deck features before round trips.

## Checklist (to be ticked while implementing)
- [x] Unblock tests on all TFMs: guard the Examples namespace import for net472 and ensure TemplateParity tests run where Examples are available.
- [x] Baseline diff harness: manifest parity + required-entry checks now enforced in tests (blank, title, tables/charts); extend to XML node diffs if more deltas appear.
- [ ] Template-aligned scaffolding: adjust `PowerPointUtils.CreatePresentationParts` to emit the same master/layout/theme/view/tableStyles/core/app/thumbnail structure as `PowerPointBlank.cs`, including creation IDs and language defaults where required.
- [x] Structure and typing: refactor oversized files (>~600 lines) into partials or focused classes; favor strongly-typed OpenXml elements/enums over manual XML, matching SDK patterns for maintainability.
- [x] Table generation parity: update `AddTable`/`PowerPointTable` to emit tableStyleId, bandRow/firstRow flags, column and row IDs (a16 ext), default widths/heights matching the templates, and seed header/body paragraphs with language-aware run/paragraph properties.
- [x] Table styling surface: expose API for borders, fills, alignment, padding, autofit, style presets, style names, and banding toggles; ensure emitted XML carries table style metadata and tableStyles.xml content.
- [x] Chart packaging parity foundation: create chart parts, style/color parts, and embedded workbooks at deterministic `ppt/charts/chart{n}.xml` and `ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx` package locations.
- [ ] Chart authoring API: allow editing series/categories while preserving the template schema (style/color links, external workbook references), with axis ID generation that matches PowerPoint conventions; add parity tests for both charts in the template deck.
- [ ] Text and bullets parity: add creation IDs/placeholders when creating shapes, carry default paragraph/run properties (language, endParaRPr), and expand bullet support (levels, numbering, indent) to match template slides; verify title/subtitle/textbox slides against templates.
- [ ] Properties, notes, and thumbnails: align core/app properties, notes master, view properties, SaveSubsetFonts, and thumbnail generation to the template values; validate with OpenXmlValidator to ensure zero errors across TFMs.
- [x] Feature inspection and round-trip preflight: expose `PowerPointPresentation.InspectFeatures()` for editable, partially editable, preserved, and unsupported features including tables, charts, images, media, SmartArt, notes, transitions, external hyperlinks, animations/timing, custom XML, embedded packages, VBA macros, web extensions/task panes, comments, and digital signatures.
- [ ] Regression gates: extend TemplateParity tests to check manifests and selected XML nodes for all three templates; run on net472/net8/net10; add a smoke test that opens/saves and confirms no repair dialog by structure (Open XML validation + part placement).

## Exit Criteria
- All checklist items checked.
- Parity tests green on net472/net8/net10; no repair prompt when opening generated decks in PowerPoint.
- Features in templates (layouts, text, bullets, tables, charts, styles) are available through the public API with documented usage.
