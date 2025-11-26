# PowerPoint Parity Plan

Last updated: 2025-11-26. Goal: make OfficeIMO.PowerPoint generate PPTX files that are structurally identical to the PowerPoint-generated templates in `Assets/PowerPointTemplates` (blank, title, tables+charts) and ship a no-repair experience across net472/net8/net9, while keeping code split into small (~600 lines) partial/typed units that mirror the OpenXml SDK object model (no ad-hoc zip surgery after save).

## Current Findings
- PowerPoint templates include creation IDs on shapes/placeholders, full layout set, SaveSubsetFonts, and language tags (pl-PL/en-US). Our shape/text creation omits creation IDs and most language defaults.
- Table slides in the templates carry tableStyleId, banding flags, a16:colId/rowId extLst entries, and richer cell properties; our `AddTable` creates a minimal table (no IDs, banding, borders, or style metadata).
- Template tableStyles.xml in the tables+charts deck contains full style definitions; our generator writes only a bare TableStyleList with a default GUID.
- Chart packaging in templates lives under `ppt/charts` with paired style/color parts and Excel embeddings named `Microsoft_Excel_Worksheet*.xlsx`; our `AddChart` builds a minimal chart and relies on `RebaseChartParts`, with hard-coded part names that may not scale past two charts and performs post-save rewrites we want to eliminate.
- Manifest parity tests compile only for net8/net9; net472 build fails because `PowerPoint.TemplateParityTests` references `OfficeIMO.Examples.PowerPoint` even though Examples are not referenced on NETFRAMEWORK.

## Checklist (to be ticked while implementing)
- [x] Unblock tests on all TFMs: guard the Examples namespace import for net472 and ensure TemplateParity tests run where Examples are available.
- [x] Baseline diff harness: manifest parity + required-entry checks now enforced in tests (blank, title, tables/charts); extend to XML node diffs if more deltas appear.
- [ ] Template-aligned scaffolding: adjust `PowerPointUtils.CreatePresentationParts` to emit the same master/layout/theme/view/tableStyles/core/app/thumbnail structure as `PowerPointBlank.cs`, including creation IDs and language defaults where required.
- [x] Structure and typing: refactor oversized files (>~600 lines) into partials or focused classes; favor strongly-typed OpenXml elements/enums over manual XML, matching SDK patterns for maintainability.
- [ ] Table generation parity: update `AddTable`/`PowerPointTable` to emit tableStyleId, bandRow/firstRow flags, column and row IDs (a16 ext), default widths/heights matching the templates, and seed header/body paragraphs with the same run/paragraph properties.
- [ ] Table styling surface: expose API for borders, fills, alignment, and banding toggles; ensure emitted XML mirrors template table cells and tableStyles.xml content; add parity tests for slides 2 and 5 from the tables+charts template.
- [ ] Chart packaging parity: base chart part creation on the template structure (chartSpace, axes, externalData, style/color parts, embeddings, media) and remove the need for post-save rebasing by writing parts to their final locations with deterministic naming aligned to `ppt/charts/chart{n}.xml` and `ppt/embeddings/Microsoft_Excel_Worksheet{n}.xlsx` (currently still using one-time rebase after save because OpenXml SDK lacks a target-URI overload for chart parts).
- [ ] Chart authoring API: allow editing series/categories while preserving the template schema (style/color links, external workbook references), with axis ID generation that matches PowerPoint conventions; add parity tests for both charts in the template deck.
- [ ] Text and bullets parity: add creation IDs/placeholders when creating shapes, carry default paragraph/run properties (language, endParaRPr), and expand bullet support (levels, numbering, indent) to match template slides; verify title/subtitle/textbox slides against templates.
- [ ] Properties, notes, and thumbnails: align core/app properties, notes master, view properties, SaveSubsetFonts, and thumbnail generation to the template values; validate with OpenXmlValidator to ensure zero errors across TFMs.
- [ ] Regression gates: extend TemplateParity tests to check manifests and selected XML nodes for all three templates; run on net472/net8/net9; add a smoke test that opens/saves and confirms no repair dialog by structure (Open XML validation + part placement).

## Exit Criteria
- All checklist items checked.
- Parity tests green on net472/net8/net9; no repair prompt when opening generated decks in PowerPoint.
- Features in templates (layouts, text, bullets, tables, charts, styles) are available through the public API with documented usage.
