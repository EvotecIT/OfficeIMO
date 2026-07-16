# OfficeIMO.PowerPoint Competitive Gap Review

Date: 2026-07-10

Branch/worktree: `codex/powerpoint-competitive-gap-20260710` at `C:\Support\GitHub\_worktrees\OfficeIMO-powerpoint-competitive-gap`

Baseline: `origin/master` at `28384ffa8ad626cf086475a96d88f0ba11d45e5e`

## Binary PowerPoint update (2026-07-16)

The original review below is a dated PPTX-focused snapshot. The current `feature/binary-powerpoint` work adds
first-party `.ppt`, `.pot`, and `.pps` lifecycle support through the normal editable model: dependency-free
compound-file and record parsing, fresh native writing, preservation-aware incremental edits, explicit
PPTX-to-binary conversion gates, RC4 CryptoAPI password compatibility, and signature mutation policy.

`LegacyPptCapabilityCatalog` is the current source of truth. Every feature reports import, new-binary,
binary-round-trip, and PPTX-to-binary states as native, preserved, converted, or blocked; no row remains
provisional. This does not imply that the PowerPoint 97-2003 format can represent every PPTX feature. Unsafe
or non-representable conversion remains blocked unless the catalog defines and preflight reports an explicit
conversion.

## Executive Verdict

OfficeIMO.PowerPoint is already a serious editable-PPTX engine. It is well beyond a basic Open XML wrapper: the current package has a broad slide object model, layout and arrangement helpers, data-bound tables and charts, media, sections, transitions, notes, theme editing, slide import and duplication, feature preflight, designer composition, semantic deck plans, first-party image/HTML/PDF paths, and markup-to-PowerPoint generation.

The product is not yet end-to-end competitive with mature commercial presentation engines. The biggest gap is not another low-level shape API. It is the reliable path from content and a brand to a beautiful, accessible, overflow-free deck whose PPTX, preview images, HTML, and PDF agree closely enough to trust.

The strongest market position is:

> An MIT-licensed, server-safe .NET presentation engine where developers describe real business content, apply a brand or template, receive an editable and visually coherent deck, and can prove the result without installing PowerPoint.

OfficeIMO should not try to become a free Aspose clone feature by feature. It should keep the current Open XML core conservative, close the most visible authoring and fidelity gaps, and invest heavily in the semantic design, content-fit, template, proof, and multi-format workflow that commercial object models do not make especially pleasant.

## Proof Run

The current dedicated PowerPoint lane passes on the implemented P0-P6 branch:

```powershell
dotnet test OfficeIMO.PowerPoint.Tests\OfficeIMO.PowerPoint.Tests.csproj -c Release -f net8.0 --nologo
```

Result: `621/621` tests passed, with no failures or skips.

The reviewed area contains:

- 185 C# files and roughly 36,500 source lines in `OfficeIMO.PowerPoint`;
- 70 PowerPoint test files and roughly 18,900 test lines;
- 21 runnable PowerPoint example files;
- dedicated `OfficeIMO.PowerPoint.Html`, `OfficeIMO.PowerPoint.Pdf`, and `OfficeIMO.Markup.PowerPoint` packages;
- no open PowerPoint-focused pull requests at the time of the review.

## What We Have Today

### Editable presentation core

`OfficeIMO.PowerPoint` can create, open, edit, validate, save, encrypt, and inspect `.pptx` presentations without Office automation. The public model covers slides, sizes, layouts, placeholders, text boxes, paragraphs and runs, pictures, auto-shapes, connectors, groups, tables, charts, audio, video, SmartArt, notes, sections, transitions, backgrounds, themes, properties, and slide view settings.

Important strengths include:

- slide add, remove, move, duplicate, import, export, hide, and layout selection;
- shape lookup, duplication, grouping, alignment, distribution, stacking, grids, fit, resize, rotation, crop, fills, outlines, shadows, glow, reflection, and z-order;
- rich text, hyperlinks, Markdown-to-text-box formatting, bullets and numbering, margins, spacing, anchoring, direction, and auto-fit;
- typed table binding, rows and columns, cell merges, fills, borders, padding, paragraph lists, sizing, and style metadata;
- data-bound column, line, scatter, pie, and doughnut chart authoring, plus extensive axis, label, series, marker, legend, and area formatting;
- embedded audio and video with poster frames and playback timing metadata;
- common transitions including Morph fallback markup, timing, speed, and advance controls;
- theme colors and font editing across slide masters;
- Open XML validation and password-based encrypted open/save.

This is enough for many reporting, training, status, and template-update workflows today.

### Round-trip preflight

`PowerPointPresentation.InspectFeatures()` is a meaningful differentiator. It classifies discovered features as editable, partially editable, preserved, or unsupported before an edit-heavy round trip.

The report currently treats slides, package scaffolding, sections, text boxes, tables, table style metadata, speaker notes, and common transitions as editable. It explicitly calls charts, images, audio/video, SmartArt, and external relationships partially editable. It detects preserve-only rich notes, advanced timing, comments, custom XML, embedded packages, ActiveX, VBA, task panes, web extensions, and unknown transition markup. Digital signatures are reported as unsupported.

That honesty should remain part of the product contract as deeper support is added.

### Designer and semantic authoring

The designer layer is more ambitious than the normal .NET presentation-library surface. It includes:

- `PowerPointDesignBrief`, design directions, recipes, deterministic seeds, preference scoring, and explainable recommendations;
- palette strategies, typography strategies, moods, density, visual styles, layout strategies, and creative-direction packs;
- `PowerPointDeckPlan` with section, case-study, process, card-grid, logo-wall, coverage, capability, and custom slide kinds;
- per-slide layout variants, card surfaces, process connectors, metric strips, title accents, motifs, and visual-frame treatments;
- `PowerPointPresentation.Compose(...)` for semantic decks and `PowerPointSlideCompositionContext` for custom slides inside a plan;
- editable visual placeholders, image frames, real logo assets, coverage pins, metric strips, and diagram-like surfaces.

The checked-in screenshots demonstrate strong typography, alignment, whitespace, editable geometry, and coherent palette use. They also expose the current ceiling: repeated panels and card grids, decorative placeholders instead of real visual evidence, limited slide grammar, and insufficient content-aware variation across a full story.

### Semantic input and adjacent outputs

The PowerPoint area already spans more than `.pptx` authoring:

- `OfficeIMO.Markup.PowerPoint` maps the presentation markup AST to editable slides, uses the designer theme, supports explicit placement and columns, creates native charts, resolves local images, and can render Mermaid diagrams through an optional external renderer.
- `OfficeIMO.PowerPoint.Html` exports semantic slide HTML or positioned review HTML and can reconstruct a bounded subset from OfficeIMO semantic HTML.
- `OfficeIMO.PowerPoint.Pdf` exports supported slide content through `OfficeIMO.Pdf`, reports conversion warnings, and can import logical PDF tables into editable PowerPoint table slides.
- the PowerPoint image exporter produces PNG and SVG proof without PowerPoint automation and emits diagnostics for unsupported content.
- extraction helpers produce slide-aligned Markdown chunks for search, ingestion, and summarization workflows.

This package family is a good foundation for an end-to-end product. The adapters must stay thin; layout, chart, text measurement, and visual-scene behavior should not be reimplemented separately in Markup, HTML, PDF, examples, or PSWriteOffice.

## Competitive Benchmark

The comparison uses representative products rather than pretending every competitor has the same goal.

| Benchmark | What it establishes | OfficeIMO position |
| --- | --- | --- |
| [Aspose.Slides for .NET](https://docs.aspose.com/slides/net/features-overview/) | Commercial completeness: many PowerPoint formats, masters/layouts, SmartArt, OLE, ActiveX, VBA, animations, slide shows, rendering, printing, and broad conversion | Far behind on advanced feature authoring and fixed-layout fidelity; ahead on licensing openness and semantic designer direction |
| [Syncfusion PowerPoint Library](https://help.syncfusion.com/document-processing/powerpoint/powerpoint-library/net/supported-and-unsupported-features) | Enterprise .NET baseline: create/edit, placeholders, SmartArt, merge/copy, encryption, properties, PDF/image conversion, and notes across common application models | Competitive on core creation and several automation workflows; behind on mature conversion and SmartArt breadth |
| [GemBox.Presentation](https://www.gemboxsoftware.com/presentation) | Compact commercial .NET API with broad format, rendering, printing, comments, masters/layouts, media, HTML loading, and macro-preservation workflows | Behind on format and rendering breadth; differentiated by MIT licensing, feature preflight, and designer/markup layers |
| [PptxGenJS](https://gitbrent.github.io/PptxGenJS/docs/introduction/) | Open-source authoring usability: browser support, code-defined masters, broad media/text support, HTML table conversion, and a large example surface | OfficeIMO has the stronger typed .NET editing model and designer layer; PptxGenJS is ahead on authoring ergonomics in several high-value workflows |
| [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler) | Direct MIT .NET comparison for a clean Open XML object model and common slide, shape, table, image, and chart manipulation | OfficeIMO is broader and more productized, but must keep its public API and docs as approachable as this simpler alternative |

Three PptxGenJS capabilities are especially relevant rather than merely impressive on a checklist:

- [code-defined slide masters and named placeholders](https://gitbrent.github.io/PptxGenJS/docs/masters/) make branded deck generation straightforward;
- [table and HTML-table auto-paging](https://gitbrent.github.io/PptxGenJS/docs/html-to-powerpoint/) creates additional slides when rows overflow and can repeat headers;
- [broad, combo, and 3D chart authoring](https://gitbrent.github.io/PptxGenJS/docs/api-charts.html) includes area, bar, bubble, doughnut, line, pie, radar, scatter, and mixed charts with secondary axes.

Those are closer to OfficeIMO's intended generated-document workflow than obscure compatibility features.

## Competitive Scorecard

| Capability | Current posture | Evidence | Main gap |
| --- | --- | --- | --- |
| Core PPTX create/edit | Strong | Broad typed model, 584 passing tests, no Office dependency | Continue fixture depth and template round-trip proof |
| Slide and shape composition | Strong | Layout boxes, columns, grids, alignment, groups, effects, units | No complete freeform/custom geometry authoring model; some renderer gaps |
| Designer and semantic deck planning | Differentiated but narrow | Design briefs, alternatives, 8 semantic slide kinds, composer presets | Limited narrative grammar, real-media use, and deck-level rhythm |
| Template and brand workflow | Partial | Existing master/layout selection, placeholder editing, theme color/font updates | No first-class create-from-template/brand-kit contract or code-defined master/layout model |
| Content fit and pagination | Weak | Count-based plan diagnostics and PowerPoint auto-fit metadata | No measured preflight, collision detection, table/text auto-splitting, or repeated-header pagination |
| Tables | Strong on one slide | Typed binding, formatting, merges, sizing, style metadata | No overflow-to-next-slide workflow; no reusable appendix/table-story component |
| Charts | Strong shared authoring | All 16 `OfficeChartKind` families, categorical combo charts, secondary value axes, embedded data, accessibility summaries, and shared snapshots | Bubble, stock, surface, 3D, modern, and unusual extension families remain demand-driven gaps |
| Media and visual assets | Partial | Pictures, crop, SVG input, audio/video, poster frames, logo images | No semantic asset pipeline, annotation/callout workflow, icon strategy, or content-aware crop/focal point |
| SmartArt and diagrams | Partial | Editable Basic Process, Basic Hierarchy, and Basic Cycle semantic workflows; Mermaid adapter | Broader SmartArt gallery semantics and a general deck-plan diagram model remain demand-driven |
| Transitions and motion | Partial | 19 transition values, Morph fallback, media timing, and typed timing-tree inspection | Shape/text/chart animation authoring remains preserve-only |
| Review workflows | Read-only typed | Classic and modern comments, replies, authors, status, anchors, and JSON review inspection | Mutation, reply, and resolution APIs remain intentionally deferred |
| Accessibility | Strong generated-deck contract | Title/description/decorative/language/reading-order APIs, link/table/color/contrast checks, strict/default policies, CI JSON | Imported decks can still need remediation; more nuanced non-solid contrast and RTL policy can deepen later |
| PowerPoint fidelity and repair safety | Proven shared path | Validator, feature report, generated and PowerPoint-authored fixtures, artifact hashes, compatibility lanes, opt-in Desktop reference | Keynote/Google Slides still require external authenticated evidence; LibreOffice depends on local availability |
| Image export | Shared visual owner | Dependency-free PNG/SVG with inherited content, typography, images, groups, custom geometry, effects, tables, charts, and diagnostics | Unsupported extension content remains reported rather than silently approximated |
| PDF export | Shared and layout-aware | Faithful shared-snapshot pages, legacy selective path, warnings, notes pages, and 1/2/3/4/6/9-up handouts | Application-specific print nuances and unsupported extension content remain bounded gaps |
| HTML round trip | Bounded | Semantic/positioned export and bounded semantic import | Import is not arbitrary HTML-to-slide design and may reconstruct chart placeholder data |
| Markup-to-deck | Promising | Designer-backed semantic slides, native charts, images, Mermaid, transitions, notes | No measured overflow pagination, proof report, or reusable template/brand input |
| Performance and scale evidence | Missing | No PowerPoint benchmark project found | No repeatable generation/load/save/render baselines for large decks |
| Public onboarding | Weak | Package README is mostly accurate | Product-page quick start uses non-existent APIs and several feature claims are broader than the public surface |

## Most Important Gaps

### 1. Content fit is the central missing engine

The semantic planner currently validates content counts. It knows, for example, that more than five process steps is dense or that a case-study slide has too many sections. It does not measure the actual title, body, bullet, table-cell, or card text against the selected font, theme, and region.

`OfficeIMO.Drawing` already owns `OfficeTextLayoutEngine`, text measurement, wrapping, fitting, clipping, chart data, chart kinds, and visual rendering primitives. The PowerPoint designer and markup exporter should consume those capabilities before slide XML is emitted.

The target contract should provide:

- measured title, paragraph, list, card, cell, and label fit;
- minimum readable font and line-height policies;
- collision, off-slide, clipped-content, and empty-region diagnostics;
- deterministic alternatives: resize within limits, select another layout, split content, or fail;
- table pagination with repeated headers and configurable continuation titles;
- long-list, card-grid, capability-section, and appendix splitting;
- a preflight report that can be inspected before writing the deck;
- the same decisions for direct designer plans and OfficeIMO Markup.

This is the highest-value competitive slice because it directly turns “usually looks fine” into a contract.

### 2. Brand and template ingestion must be first class

The current designer can start from a color, fonts, a seed, identity text, and manual palette overrides. The core can open a template-like presentation, select layouts, inspect placeholders, and edit theme colors/fonts. These capabilities are not yet a coherent template workflow.

The product needs a bounded template/brand contract that can:

- create a new output deck from `.pptx` or `.potx` source material without leaving template sample slides in the result;
- inventory masters, layouts, named placeholders, theme colors/fonts, logos, footers, page size, and safe-area guides;
- select layouts by semantic role and named placeholder rather than integer index;
- map a brand kit into both native theme parts and designer tokens;
- preserve unedited master/layout content and unsupported package parts;
- explain missing or ambiguous placeholders before generation;
- optionally define new layouts in code after template consumption is reliable.

Template consumption should come before a large master-authoring framework. Real customers already have branded templates; making those work predictably closes more business value with less speculative API design.

### 3. The semantic slide grammar is too small for a complete story

The existing eight deck-plan kinds are a solid start, but most current showcase slides resolve to cards, rails, panels, placeholder visuals, or metric strips. Beautiful decks need coherent variation across the whole narrative.

The next reusable slide families should be driven by real deliverables:

- executive summary and decision slide;
- agenda and section navigator;
- hero image or screenshot story;
- KPI/dashboard and chart-story slide;
- comparison, before/after, and option-decision slide;
- architecture/system diagram and annotated process;
- quote, customer proof, and case-study result;
- editable table/appendix continuation;
- team/profile and operating-model slide;
- closing recommendation, next actions, and Q&A.

Each family needs multiple layout variants, real image or chart inputs, content-fit rules, accessibility defaults, and proof fixtures. Examples should use these shared components instead of becoming a second design engine.

### 4. Chart authoring is behind both competitors and OfficeIMO's own shared model

The public PowerPoint authoring surface provides clustered column/bar, line, scatter, pie, and doughnut helpers. At the same time:

- `OfficeIMO.Drawing.OfficeChartKind` already defines 16 chart families, including stacked/100% column and bar, stacked/100% line, area variants, and radar;
- the PowerPoint chart snapshot reader already understands stacked bar/column/line, area, radar, scatter, pie, doughnut, and some mixed chart content;
- `OfficeIMO.Excel` exposes a much broader chart family and combo/secondary-axis concepts.

The right fix is to converge on shared chart semantics in `OfficeIMO.Drawing` while keeping PPTX serialization in `OfficeIMO.PowerPoint`. Do not add another PowerPoint-only enum and another parallel data model.

The first parity slice should author every chart family already represented by `OfficeChartKind`, then add combo/secondary-axis support. Bubble, stock, surface, modern charts, 3D, and unusual PowerPoint chart extensions can follow demand and fixture proof.

### 5. Accessibility needs a real product contract

`AltText` is useful but not sufficient. A generated deck should be able to prove:

- informative visuals have meaningful alternative text;
- decorative geometry is marked decorative where the file format supports it;
- every slide has a title or an explicit exception;
- reading order is intentional and inspectable;
- text and background contrast meet the selected policy;
- links have useful labels;
- language and right-to-left settings propagate consistently;
- tables identify header structure;
- no important meaning depends only on color or position.

Add an `InspectAccessibility()`-style report before adding a large set of ad hoc setters. Designer components should emit accessible defaults, and the report should remain useful for imported decks.

### 6. Export paths need one fidelity story

PPTX, PNG/SVG, HTML, and PDF currently share some `OfficeIMO.Drawing` primitives, but they are not yet equivalent views of one proven slide scene. The differences are visible in image codec support, custom geometry, inherited content, text layout, fonts, effects, backgrounds, and unsupported-shape behavior.

The shared owner should be a dependency-light slide visual snapshot/scene projected from PowerPoint content. Image, HTML positioned review, and PDF adapters should consume that snapshot and add format-specific policy rather than rediscovering slide semantics.

The fidelity roadmap should cover:

- master/layout inherited shapes and background resolution;
- text shaping, fallback, wrapping, paragraph spacing, bullets, and RTL;
- all advertised image input formats in preview paths or explicit normalization during authoring;
- gradient, pattern, transparency, custom geometry, group transforms, connectors, crop, and effects;
- chart and table parity through shared snapshots;
- SmartArt/media poster fallbacks with diagnostics;
- notes-page and handout PDF layouts after ordinary slide fidelity is stable.

Do not hide unsupported content. Keep warnings machine-readable and make strict profiles fail when fidelity requirements are not met.

### 7. Public docs currently weaken trust

The package README uses current APIs, but the website product-page quick start uses `PowerPointDocument`, `SlideLayoutType`, `ChartType.ColumnClustered`, `chart.Title.Text`, and `chart.AddSeries`, none of which form the current OfficeIMO.PowerPoint public workflow.

The product page also claims full chart support, area chart authoring, freeform shapes, and broad custom-theme behavior more strongly than the inspected public API supports.

Before marketing more capability:

- compile every public C# snippet in CI;
- generate the product-page feature list from a maintained capability manifest or contract tests;
- label features as author, edit, preserve, render, import, or inspect rather than using one ambiguous “supported” flag;
- publish the designer screenshots and the source command that generated them;
- add one end-to-end guide from data and a brand to PPTX, PNG, PDF, and diagnostics;
- keep PowerShell examples thin over the same OfficeIMO.PowerPoint engine.

### 8. Advanced PowerPoint completeness remains a later gap

Commercial engines are still ahead on:

- shape, text, table, and chart animations with timelines and triggers;
- rich notes-page and notes-master editing;
- classic and modern comments, replies, authors, and resolution state;
- broad SmartArt layout creation and editing;
- OLE object authoring and embedded package management;
- VBA project editing, ActiveX, task panes, and web extensions;
- digital signature inspection, validation, signing policy, and save safety;
- custom slide shows, handouts, printing, and broader presentation formats such as PPT, PPSX, ODP, and POTX/POTM.

These matter for object-model completeness, but most do less for beautiful generated business decks than content fit, templates, visual grammar, charts, accessibility, and fidelity. Preserve and report them safely first. Implement them only through real user workflows and sanitized fixture proof.

## Recommended Roadmap

### P0: Trustworthy end-to-end generation

- [x] Fix the PowerPoint product-page quick start and narrow unsupported marketing claims.
- [x] Add compile-tested public snippets and a maintained author/edit/preserve/render capability matrix.
- [x] Introduce a measured deck preflight contract using `OfficeIMO.Drawing` text layout and chart primitives.
- [x] Detect overflow, collisions, off-slide shapes, clipped text, unreadable font reduction, and missing visual assets.
- [x] Add automatic continuation slides for tables and long semantic content, including repeated table headers.
- [x] Add one machine-readable generation report shared by designer and markup workflows.
- [x] Publish one real end-to-end example that emits PPTX, PNG/SVG previews, PDF, and the proof report.

Exit criteria:

- a generated deck can fail before save when content cannot fit under the selected policy;
- normal overflow can split deterministically without losing content;
- all public examples compile and every claim maps to a tested support level;
- proof artifacts are easy to inspect without PowerPoint.

### P1: Template and brand-kit workflow

- [x] Define the smallest template inventory model for masters, layouts, placeholders, theme, logos, footer content, slide size, and safe areas.
- [x] Add a create-from-template workflow with deterministic removal or retention of source slides.
- [x] Select layout roles and placeholders by semantic name with clear ambiguity diagnostics.
- [x] Map imported brand tokens into `PowerPointDesignBrief` and native PowerPoint theme parts.
- [x] Prove preservation with sanitized PowerPoint-authored `.pptx` and `.potx` fixtures.
- [x] Add code-defined layouts only after the template-consumption contract is stable.

Exit criteria:

- a caller can point at a real corporate template, inspect what OfficeIMO understands, render a semantic plan into named layouts, and save without repair or unintended master loss.

### P2: Broader visual grammar and real assets

- [x] Add executive-summary, chart-story, comparison, screenshot-story, appendix-table, architecture, and closing slide families first.
- [x] Give each family at least two materially different layouts rather than color-only variations.
- [x] Accept real images, charts, tables, diagrams, and callouts through semantic content models.
- [x] Add crop/focal-point, annotation, caption, and provenance/alt-text metadata for images.
- [x] Add deck-level rhythm checks for consecutive repeated variants, density, dark/light balance, and section pacing.
- [x] Expand the market-facing showcase with complete multi-slide deliverables rather than isolated component demos.

Exit criteria:

- the same semantic story can render in several recognizably different but coherent brand directions;
- generated decks use real evidence and data, not decorative placeholders, when assets are supplied;
- a full 12-20 slide business deck does not look like one card-grid template repeated.

### P3: Shared chart authoring parity

- [x] Route PowerPoint chart family and data semantics through `OfficeIMO.Drawing.OfficeChartKind` and shared chart contracts.
- [x] Author all currently shared column, bar, line, area, scatter, radar, pie, and doughnut variants.
- [x] Add combo charts and primary/secondary axis assignment.
- [x] Keep embedded workbook data, cached values, chart XML, snapshots, image export, HTML, and PDF consistent.
- [x] Add chart accessibility metadata and data-summary output.
- [x] Add broader families only with real fixture and renderer demand.

Exit criteria:

- PowerPoint authoring no longer supports fewer chart families than its own snapshot/export model;
- the same shared chart data can drive Excel, PowerPoint, image, HTML, and PDF surfaces without duplicated semantics.

### P4: Accessibility and validation

- [x] Add an accessibility inspection model and strict policy profile.
- [x] Support shape title, description, decorative state, reading order, language, and meaningful link labels where OOXML allows it.
- [x] Add contrast, missing-title, missing-alt-text, table-header, and color-only-meaning checks.
- [x] Make designer components emit accessible defaults.
- [x] Add fixture-based reports for generated and imported decks.

Exit criteria:

- accessibility can be enforced in CI with structured findings;
- designer-generated decks pass the default accessibility profile without caller cleanup.

### P5: Fixed-layout fidelity and proof corpus

- [x] Define one shared slide visual snapshot used by dependency-free image, positioned HTML, and PDF paths.
- [x] Close inherited master/layout, typography, image-codec, group-transform, custom-geometry, effect, table, and chart fidelity gaps in that shared owner.
- [x] Build a sanitized golden-deck corpus covering generated decks and PowerPoint-authored real-world decks.
- [x] Record structural, extraction, accessibility, conversion-warning, artifact-hash, and perceptual visual evidence per deck.
- [x] Add opt-in PowerPoint desktop reference rendering without making Office automation the default engine.
- [x] Add explicit LibreOffice, Keynote, and Google Slides compatibility lanes. LibreOffice runs when installed; Keynote and Google Slides accept externally recorded evidence because those environments are not available in ordinary Windows CI.

Exit criteria:

- every promoted showcase deck has PPTX, PNG/SVG, PDF, HTML review, structural, accessibility, and diagnostic proof;
- visual differences from the accepted reference are reviewable and intentional.

### P6: Advanced presentation workflows

- [x] Add typed classic/modern comment and reply read models before comment mutation.
- [x] Add animation/timing-tree inspection before shape/text/chart animation authoring.
- [x] Deepen SmartArt through Basic Process, Basic Hierarchy, and Basic Cycle semantic workflows.
- [x] Add speaker-notes pages and 1/2/3/4/6/9-up handout PDF export after ordinary slide fidelity is stable.
- [x] Add a safe block/remove/preserve signature mutation policy before any signing API.
- [x] Keep OLE, macros, custom shows, and additional file-format authoring behind concrete consumer demand; existing feature inspection and package preservation remain the contract.

## PR-Sized First Implementation Train

The roadmap should not land as one giant redesign. A practical first train is:

1. **Docs truth** — replace the broken product quick start, correct chart/theme claims, and compile public snippets.
2. **Preflight model** — add read-only measured content and collision diagnostics without changing rendering.
3. **Continuation engine** — split typed tables and long card/list content into deterministic continuation slides.
4. **Markup adoption** — make `OfficeIMO.Markup.PowerPoint` consume the same preflight and continuation decisions.
5. **Template inventory** — inspect a real template and expose semantic layout/placeholder selection.
6. **Chart convergence** — author the existing shared `OfficeChartKind` families before introducing new chart families.
7. **Accessibility report** — inspect generated/imported decks and make designer output pass the default policy.

The first three slices produce the fastest visible gain: public trust, no silent overflow, and better long-report generation.

## Ownership Boundaries

Keep one brain for each capability:

| Owner | Responsibility |
| --- | --- |
| `OfficeIMO.PowerPoint` | PPTX package model, editable slide content, template/master/layout mapping, deck planning, semantic components, PowerPoint-specific validation and preservation |
| `OfficeIMO.Drawing` | Text measurement and layout, shared chart kinds/data/style, scene primitives, geometry, image normalization, rendering diagnostics, visual snapshots |
| `OfficeIMO.PowerPoint.Pdf` | Thin mapping from the shared slide visual snapshot to `OfficeIMO.Pdf` options and diagnostics |
| `OfficeIMO.PowerPoint.Html` | Thin semantic/positioned HTML adapter and bounded HTML round-trip policy |
| `OfficeIMO.Markup.PowerPoint` | Maps markup AST and attributes into PowerPoint deck-plan/component inputs; it must not own another layout or chart engine |
| `OfficeIMO.Examples`, Website, PSWriteOffice | Public examples and thin consumer surfaces over the same core contracts |

Do not add compatibility bridges in PSWriteOffice or examples while waiting for an OfficeIMO package release. Validate unreleased work through local source or built packages, publish the owner, then repin consumers.

## Maintainability Guardrails

Several files are at or above the repository's structural review threshold:

| File | Approximate lines | Split when touched |
| --- | ---: | --- |
| `Imaging/PowerPointSlideImageRenderer.cs` | 1,210 | Shape dispatch, snapshot projection, background/inheritance, and diagnostics |
| `PowerPointFeatureReport.cs` | 1,087 | Feature inspectors by structure, media, review, timing, and compatibility |
| `PowerPointDesignBrief.cs` | 918 | Alternative generation, preference scoring, content-fit scoring, and overrides |
| `PowerPointSlide.Charts.cs` | 883 | Chart-family creation and typed data overloads |
| `PowerPointDesignModels.cs` | 859 | Design enums, slide options, and semantic content models |
| `PowerPointDesignSpecializedExtensions.cs` | 853 | Logo wall, coverage, and capability components |
| `PowerPointUtils.Charts.cs` | 763 | Package creation, chart XML families, and workbook data |
| `PowerPointTableCell.cs` | 728 | Text, geometry, style, and merge responsibilities |
| `PowerPointTextBox.cs` | 725 | Body properties, paragraphs, text formatting, and layout behavior |

Do not start a broad file-shuffling project. Split a file only when the next feature touches a clear responsibility, move behavior without changing the public contract where possible, and validate after each extraction.

## Market-Readiness Proof Set

Promote a small corpus of complete, realistic decks rather than hundreds of disconnected feature slides:

1. Executive weekly status and decision deck.
2. Customer QBR with charts, tables, screenshots, and notes.
3. Technical architecture and rollout proposal.
4. Product launch or release story.
5. Training/onboarding deck with diagrams and callouts.
6. Financial/KPI dashboard with appendix tables.
7. Case study with real visual proof and customer outcomes.
8. Template-driven corporate deck generated from a `.potx` or branded `.pptx`.

Every promoted deck should provide:

- source data/markup and the normal public API command that generated it;
- editable `.pptx` output;
- Open XML validation and package-integrity result;
- feature and accessibility reports;
- content-fit/overflow report;
- PNG and SVG slide previews;
- PDF and positioned HTML review output;
- conversion diagnostics and artifact hashes;
- optional PowerPoint desktop reference rendering;
- a browsable gallery that pairs each slide with its proof.

Performance should be measured for small, normal, and large decks across create/save, open/edit/save, image export, and PDF export. Establish baselines before setting budgets; do not optimize against invented targets.

## End-to-End Definition Of Done

“Beautiful documents end to end” should mean all of the following, not only an attractive screenshot:

- **Intentional** — the deck has a clear hierarchy, narrative rhythm, coherent brand, and appropriate visual variety.
- **Complete** — no input content is silently dropped, clipped, hidden, or collapsed into unreadable text.
- **Editable** — text, tables, charts, diagrams, and shapes remain native where the selected profile promises editability.
- **Accessible** — reading order, titles, alternative text, language, contrast, and table semantics pass a defined policy.
- **Compatible** — the file validates, opens without repair, and preserves advanced content outside the edited surface.
- **Consistent** — PPTX, preview, HTML, and PDF differences are bounded and reported.
- **Repeatable** — the same content, seed, brand, and options produce stable output suitable for CI.
- **Observable** — callers receive structured layout, compatibility, accessibility, and conversion reports.
- **Maintainable** — shared layout, chart, and rendering logic lives in the owning core, with thin adapters and examples.

That is the competitive bar OfficeIMO.PowerPoint should optimize for.
