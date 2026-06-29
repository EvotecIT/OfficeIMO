# OfficeIMO.Markdown Markdig Parity Gap Plan

This is the working board for getting `OfficeIMO.Markdown` to Markdig-class behavior without looping through disconnected fixture additions.

Parity is not "more tests." Tests are the measuring system. Parity means the reusable engine, AST, source model, renderer, writer, extension seams, security profiles, docs, and benchmarks all agree on a contract.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 316 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 652 of 652 official CommonMark `0.31.2` examples currently match; 0 are failing in the generated CommonMark inventory |
| GFM corpus | 52 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| GFM inventory | 52 tracked GFM fixtures in the generated GFM inventory: 48 upstream cmark-gfm fixtures, 4 OfficeIMO supplements, 52 passing, 0 failing |
| Markdig extension inventory | 33 Markdig extension-family rows in `Docs/officeimo.markdown.markdig-extension-inventory.md`: 10 covered, 9 partial, 3 intentional, 11 gap |
| Markdig extension compatibility matrix | Generated control board in `Docs/officeimo.markdown.markdig-compatibility-matrix.md` splits every Markdig row into Decision, Engine parser, AST/source, Writer/render, Proof, and next-action lanes |
| Remaining architecture gaps | broader GFM breadth, Markdig extension breadth, canonical AST ownership, full lossless trivia/source mapping, source-aware renderer/writer extension seams, security/profile separation, and release-mode benchmark evidence |

## No-Loop Missing Parity Checklist

Use this section to choose work. If a row below is not the active row, do not add nearby fixtures just because they are convenient.

- [ ] **P0 - Finish the active `UseDefinitionLists` promotion.**
  - [ ] Engine: close the remaining lazy-continuation cases, especially paragraphs after blank lines, multiple lazy lines, lazy lines after nested blocks, and remaining interruption cases from list/table-shaped starts.
  - [ ] Engine: close nested-body breadth for blockquotes, fenced code, HTML, setext headings, setext-looking lines inside `<dd>`, and list tails after nested list bodies.
  - [ ] AST/source: finish exact spans for marker lines, continuation indentation, blank separators, generated paragraph wrappers, and native `definitionBody` values.
  - [ ] Writer: prove every fixed shape writes Markdown that reparses to the same Markdig-compatible HTML and keeps source-backed edits valid.
  - [ ] Proof: only then move `UseDefinitionLists` from `Partial` to `Covered` in the generated inventory.

- [ ] **P1 - Close the remaining high-value Markdig partial rows.**
  - [ ] `UseGenericAttributes`: extend from the covered standalone/block/inline shapes to arbitrary supported block families and remaining inline families, with source tokens and writer behavior.
  - [ ] `UseAlertBlocks`: decide whether Markdig alert callbacks become an OfficeIMO renderer contract or remain an intentional callout difference; then align AST/source/writer around that decision.
  - [ ] `UseCjkFriendlyEmphasis`: add a real Markdig-compatible delimiter option with source-token proof, or document it as deferred/intentional.
  - [ ] `UsePreciseSourceLocation`: keep partial until lossless trivia, original mapping, generated-node diagnostics, and broader source edits are complete.

- [ ] **P2 - Decide the real gaps before implementing optional syntax.**
  - [ ] `UseCustomContainers`: decide optional extension shape, then implement colon-fenced containers with child source ownership, renderer seams, and writer support.
  - [ ] `UseGridTables`: decide whether grid tables belong in core or optional extension; if yes, add parser, semantic table model, malformed fallback, source spans, renderer, and writer.
  - [ ] `UseListExtras`: inventory Markdig list-extra syntax first; do not start from tests until the syntax contract is known.
  - [ ] Optional transforms: keep `UseEmojiAndSmiley`, `UseJiraLinks`, and `UseSmartyPants` out of the core path unless a product need makes them active.
  - [ ] Deferred rows: keep `UseCitations`, `UseFooters`, `UseGlobalization`, and `UsePragmaLines` deferred until a real consumer needs them.

- [ ] **P3 - Make the app editor-grade, not only renderer-compatible.**
  - [ ] Canonicalize duplicated semantic/syntax ownership for lists, tables, definition lists, callouts, footnotes, front matter, and extension nodes.
  - [ ] Associate syntax nodes with semantic subobjects such as callout titles, list item paragraphs, definition groups/definitions, and sequence-style inline wrappers.
  - [ ] Capture lossless trivia for whitespace, blank lines, tabs, delimiters, raw slices, normalized text, and generated-node diagnostics.
  - [ ] Complete delimiter-token coverage for emphasis extras, links/images, escapes/entities, breaks, HTML, footnotes, front matter, tables, and extension nodes.
  - [ ] Establish one original-to-normalized mapping story for line endings, tab expansion, nested containers, generated nodes, and normalized paragraph text.
  - [ ] Broaden `MarkdownRoundtripWriter` beyond unchanged documents and explicit native edits, with precise fallback diagnostics.

- [ ] **P4 - Keep renderer, writer, and security policy explicit.**
  - [ ] Build source-aware extension seams for custom blocks, inlines, transforms, renderers, and writers so downstream code does not rescan strings.
  - [ ] Separate raw HTML grammar from security policy: CommonMark raw HTML, GFM tag filtering, allow/strip/escape/sanitize modes, URL policy, source metadata, and Markdown writing.
  - [ ] Bound renderer/host rows such as `UseDiagrams`, `UseFigures`, `UseMathematics`, `UseMediaLinks`, and `UseReferralLinks` before parser work starts.

- [ ] **P5 - Use tests as proof, not as the product.**
  - [ ] Regenerate inventories and compatibility matrix after each promoted row.
  - [ ] Broaden GFM fixtures only after the covered grammar/source behavior is stable.
  - [ ] Run release-mode benchmarks after correctness, source mapping, and writer behavior stop moving.

## No-Loop Parity Exit Plan

Parity is closed only when the boxes below are closed. A test can prove a box, but it cannot replace the implementation behind it.

- [ ] **Close required grammar behavior.**
  - [ ] Finish `UseDefinitionLists` promotion or document the exact remaining writer/source limits.
  - [ ] Finish `UseGenericAttributes` for the remaining Markdig-supported block and inline targets, not just the already-probed standalone-attribute cases.
  - [ ] Decide whether `UseCustomContainers`, `UseGridTables`, `UseListExtras`, and `UseCjkFriendlyEmphasis` are core engine work, optional extensions, or intentional differences.
  - [ ] Implement only the rows that are classified as engine or optional-extension parity.

- [ ] **Close AST and source parity.**
  - [ ] Every promoted row has a semantic AST shape and a syntax AST shape.
  - [ ] Every promoted row has source spans for user-addressable pieces, not only the outer block.
  - [ ] Native/source-edit APIs can find and update the important fields without string rescanning.
  - [ ] Parser-created wrapper nodes are marked or mapped clearly when they do not correspond to original source.

- [ ] **Close writer and renderer parity.**
  - [ ] HTML output matches Markdig for the active profile, or the difference is intentional and documented.
  - [ ] Markdown writing preserves behavior on reparse.
  - [ ] Exact source preservation is implemented where the app/editor contract needs it, and normalized writer output is explicitly documented where exact preservation is not promised.
  - [ ] Renderer/host-policy rows such as diagrams, math, media links, figures, and referral links have explicit ownership before parser features are added.

- [ ] **Close security/profile policy.**
  - [ ] Raw HTML grammar is separate from allow/strip/escape/sanitize policy.
  - [ ] GFM, Markdig-compatible, and OfficeIMO-safe profiles do not silently share incompatible defaults.
  - [ ] URL and raw HTML policies carry enough source metadata for diagnostics and writing.

- [ ] **Close proof and release confidence last.**
  - [ ] Inventories and compatibility matrix are regenerated after each implementation slice.
  - [ ] Broader GFM and Markdig fixture sweeps run after engine behavior is stable.
  - [ ] Release-mode benchmarks run after parser/source/writer behavior stops moving.

## What Is Missing

- [x] **CommonMark parser correctness is closed.** The official CommonMark `0.31.2` inventory is green: 652 of 652 examples match and 0 are failing.
- [x] **Core GFM behavior is real engine behavior.** Pipe tables, task lists, footnotes, strikethrough, auto identifiers, extended autolinks, soft-line-as-hard-line, YAML front matter, abbreviations, and tracked GFM fixtures have parser/render/write/source proof.
- [x] **The scoreboards exist.** CommonMark inventory, generated GFM inventory, Markdig extension inventory, Markdig extension compatibility matrix, broad compatibility matrix, benchmark hooks, and this gap plan are checked in.
- [ ] **Markdig extension parity is not closed.** The current inventory has 10 covered, 9 partial, 3 intentional, and 11 gap rows.
- [ ] **AST/source/lossless parity is not closed.** Full trivia capture, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, broader source edits, and extension-node roundtrip still need work.
- [ ] **Performance parity is not known.** Release-mode benchmark comparisons should run after parser/source/writer behavior stops moving.

## Missing Work By Kind

- [ ] **Engine work: finish real Markdown grammar gaps.**
  - [ ] `UseGenericAttributes`: arbitrary block-family attachment, every Markdig-supported inline-family target, and writer/source preservation across those shapes. Known probed standalone-target gaps are now closed for fenced code, root lists, blockquotes, HTML blocks, dash setext/thematic forms, indented code, and definition-list-looking text when only `UseGenericAttributes` is enabled.
  - [ ] `UseCustomContainers`: colon-fenced container parsing, nested child-block source ownership, renderer/writer seams, and Markdig comparison fixtures.
  - [ ] `UseGridTables`: grid-table parser, semantic table model, malformed fallback, source spans, HTML rendering, and Markdown writer behavior.
  - [ ] `UseListExtras`: inventory Markdig list-extra syntax first, then decide whether it is core list behavior or an optional parser extension.
  - [ ] `UseCjkFriendlyEmphasis`: Markdig-compatible delimiter option with CJK comparison and delimiter source-token proof.

- [ ] **Engine-plus-proof work: promote partial rows that already mostly exist.**
  - [ ] `UseDefinitionLists`: close remaining parser/source-map/writer breadth for lazy continuation variants, nested bodies, multiline bodies, and reparse stability.
  - [ ] `UseAlertBlocks`: align callout/alert AST fields, source spans, renderer customization, writer output, and Markdig/GFM comparison behavior.
  - [ ] `UseFigures`: separate HTML-import figure recovery from Markdown figure syntax, then prove parser, renderer, writer, and source behavior.
  - [ ] `UseMathematics`: decide inline/block math delimiter ownership, then add AST/source/native/writer and renderer handoff contracts.
  - [ ] `UseMediaLinks`: decide provider model and safe renderer policy for shortcut media links before adding parser behavior.

- [ ] **AST/source architecture work: make the app editor-grade, not just renderer-grade.**
  - [ ] Canonicalize duplicated semantic/syntax shapes for lists, tables, definition lists, callouts, footnotes, front matter, and extension nodes.
  - [ ] Finish syntax association for semantic subobjects such as callout titles, list-item paragraph blocks, definition-list groups, and sequence-style inline wrappers.
  - [ ] Capture lossless trivia: whitespace, blank lines, tabs, delimiters, raw slices, normalized text, and generated-node diagnostics.
  - [ ] Complete delimiter-token coverage for emphasis extras, links/images, escapes/entities, breaks, HTML, footnotes, front matter, tables, and extension nodes.
  - [ ] Establish one original-to-normalized mapping story for CRLF/LF/CR, tab expansion, nested containers, generated nodes, and normalized paragraph text.
  - [ ] Broaden `MarkdownRoundtripWriter` beyond unchanged documents and explicit native edits, with precise fallback diagnostics.

- [ ] **Renderer/writer/security work: keep policies explicit.**
  - [ ] Finish source-aware extension seams for custom blocks, inlines, transforms, renderers, and writers without downstream string rescanning.
  - [ ] Separate raw HTML grammar from security policy: CommonMark HTML, GFM tag filtering, allow/strip/escape/sanitize modes, URL policy, source metadata, and Markdown writing.
  - [ ] Bound renderer/host rows (`UseDiagrams`, `UseFigures`, `UseMathematics`, `UseMediaLinks`, `UseReferralLinks`) before implementing them as parser features.

- [ ] **Optional/deferred rows: do not let them masquerade as current parity blockers.**
  - [ ] Optional extension candidates: `UseEmojiAndSmiley`, `UseJiraLinks`, `UseSmartyPants`.
  - [ ] Deferred until real consumer need: `UseCitations`, `UseFooters`, `UseGlobalization`, `UsePragmaLines`.
  - [ ] Intentional differences to keep documented, not implemented as Markdig clones: `UseAdvancedExtensions`, `UseBootstrap`, `UseSelfPipeline`.

- [ ] **Proof-only work: useful after behavior exists.**
  - [ ] Broaden GFM fixture breadth beyond the current 52 tracked fixtures.
  - [ ] Refresh Markdig inventory rows after each engine slice.
  - [ ] Run release-mode benchmarks only after correctness and source behavior settle.

## Non-Looping Execution Rules

- [ ] Pick exactly one primary row before starting a slice: one Markdig extension family, GFM breadth, AST/source/lossless, renderer/writer seams, security/profile policy, or performance.
- [ ] If behavior is missing, improve the reusable engine first: parser, semantic AST, syntax AST, native/source projection, renderer, writer, or extension APIs.
- [ ] If behavior already exists but is unproven, add focused proof only: Markdig comparison cases, generated inventories, source/native snapshots, writer checks, renderer checks, or benchmarks.
- [ ] Do not promote a row to `Covered` until parser behavior, semantic/syntax/native/source behavior, HTML rendering, Markdown writing or explicit writer limits, docs, and proof all agree.
- [ ] Make scope decisions before new optional features. Grid tables, custom containers, math, diagrams, attributes, SmartyPants, citations, media links, and similar rows must be classified as core engine, optional extension, renderer/host policy, deferred, or intentional difference before implementation.
- [ ] Benchmark last, after correctness and source behavior are stable enough for the numbers to mean something.

## P0 - Active Slice

- [ ] **Promote or explicitly bound `UseDefinitionLists`.**
  Covered now: structured definition-list AST, Markdig-style colon-marker term grouping, multiple definitions in one group, marker syntax tokens, native source-backed marker fields/source edits, loose-definition writer preservation, blank-separated marker-group writer preservation, blank-separated pre-marker term boundary proof, table-shaped continuation profile proof with literal paragraphs when tables are off and nested tables when pipe tables are on, tight nested-list writer preservation, setext-continuation writer reparse proof, setext-following and thematic-break-following lazy-continuation boundary behavior, nested-list and nested-blockquote heading/thematic interruption boundaries, paragraph-plus-thematic-break writer reparse preservation, empty-marker first-continuation handling, empty-marker blank-separated body source/writer preservation, and multiline definition-body edits that keep continuation indentation valid for simple and marker forms.
  Missing before promotion:
  - [ ] Broaden lazy-continuation variants beyond the setext boundary: paragraphs after blank lines, multiple lazy lines, remaining lazy lines after nested blocks, and interruption by list/table-like starts, each with semantic AST, syntax AST, native source field, HTML, writer, and reparse checks.
  - [ ] Broaden nested-body cases for blockquote source breadth, fenced code, HTML, setext headings, setext-looking lines inside `<dd>`, and list tails beyond the immediate nested-list item boundary.
  - [ ] Finish multiline-body source mapping, including exact spans for marker lines, blank separators, continuation indentation, generated paragraph wrappers, and normalized native `definitionBody` values versus original source spans.
  - [ ] Promote to `Covered` only after parser behavior, source/native projection, HTML rendering, Markdown writing, reparse stability, generated inventory docs, and the compatibility matrix all agree.

## P1 - Markdig Extension Rows After Active Slice

- [ ] **Decide and close `UseAlertBlocks`.**
  Missing: an explicit decision whether Markdig alert rendering callbacks become an OfficeIMO renderer contract or remain an intentional OfficeIMO callout difference.
- [ ] **Decide and close `UseCjkFriendlyEmphasis`.**
  Missing: either a real delimiter option with CJK comparison/source-token proof, or a documented deferred/intentional decision.
- [ ] **Keep `UsePreciseSourceLocation` as a cross-cutting partial row until lossless work closes.**
  Missing: full trivia/original mapping, generated-node diagnostics, and source-edit coverage.

## P2 - AST, Source, And Lossless

- [ ] **Canonicalize duplicated AST shapes.**
  Current hotspots: `ListItem`, `TableBlock`, `DefinitionListBlock`, `CalloutBlock`, `FootnoteDefinitionBlock`, front matter, and extension-owned nodes.
- [ ] **Finish syntax association for semantic subobjects.**
  Known gaps include callout title inlines, list-item paragraph blocks, definition-list groups/definitions, and sequence-style inline wrappers.
- [ ] **Complete lossless trivia capture.**
  Missing: whitespace, blank lines, tabs, delimiter trivia, raw slices, normalized text, and generated-node diagnostics owned by parser data.
- [ ] **Complete delimiter-token coverage.**
  Missing: every editor-addressable spelling for emphasis extras, links/images, escapes/entities, hard/soft breaks, HTML tags, footnotes, front matter, tables, and extension nodes.
- [ ] **Complete original-to-normalized mapping.**
  Missing: one reliable mapping story for CRLF/LF/CR inputs, tab expansion, nested containers, transformed/generated nodes, and normalized paragraph text.
- [ ] **Broaden `MarkdownRoundtripWriter`.**
  Missing: source-preserving edits beyond unchanged documents and explicit native edits, precise fallback diagnostics, and extension-node roundtrip.

## P3 - Extension, Renderer, Writer, And Security

- [ ] **Finish source-aware extension seams.**
  Missing: custom block, inline, transform, renderer, and writer APIs that carry source slices and token metadata without downstream string rescanning.
- [ ] **Separate raw HTML grammar from security policy.**
  Missing: independent contracts for CommonMark raw HTML, cmark-gfm tag filtering, OfficeIMO allow/strip/escape/sanitize modes, URL policy, source metadata, and Markdown writing.
- [ ] **Close renderer/host rows only with explicit ownership.**
  Rows such as `UseDiagrams`, `UseFigures`, `UseMathematics`, and `UseMediaLinks` need parser/AST/source/renderer/writer promotion bars before implementation.
- [ ] **Keep optional transform/parser rows optional unless product need changes.**
  Rows such as `UseGridTables`, `UseCustomContainers`, `UseListExtras`, `UseEmojiAndSmiley`, `UseJiraLinks`, and `UseSmartyPants` need separate optional contracts.
- [ ] **Keep deferred rows deferred until a consumer needs them.**
  Rows such as `UseCitations`, `UseFooters`, `UseGlobalization`, and `UsePragmaLines` need real requirements before implementation.

## P4 - Proof-Only Work

- [ ] **Broaden GFM fixture breadth.**
  This is proof-only unless mismatches expose real engine gaps. The current tracked GFM inventory is green but small at 52 tracked GFM fixtures.
- [ ] **Refresh Markdig inventory rows after each engine slice.**
  Update Route, Scope decision, promotion bar, current state, next action, and status when behavior changes.
- [ ] **Run release-mode benchmarks last.**
  Compare parse, parse-with-syntax, HTML render, Markdown write, transforms, source edits, allocations, and representative README/docs/chat corpora against Markdig.

## Next Ordered Work

- [ ] **1. Promote or bound `UseDefinitionLists`.**
  Current probes closed empty-marker blank-separated body source/writer preservation, the blank-separated pre-marker term boundary, table-shaped continuation profile behavior, the setext-following lazy-continuation boundary, the thematic-break-following lazy-continuation boundary with writer reparse preservation, and nested-list/nested-blockquote heading/thematic interruption boundaries. Next probes should target remaining lazy continuation variants plus nested/multiline body source spans.
- [ ] **2. Continue `UseGenericAttributes`, but only after probing actual missing Markdig behavior.**
  The most recent probes closed standalone attributes before fenced code, the literal standalone-attribute-before-blockquote boundary, HTML blocks, dash setext/thematic forms, indented code, no-space paragraph attributes, and definition-list-looking text without `UseDefinitionLists`. Next probes should move beyond this standalone-target sweep into less-covered block families and remaining inline families.
- [ ] **3. Decide `UseAlertBlocks` and `UseCjkFriendlyEmphasis`.**
  These need scope decisions before more fixtures.
- [ ] **4. Return to AST/source/lossless architecture.**
  Canonical node ownership, trivia, delimiter tokens, source mapping, and broader roundtrip edits are the next big body of work.
- [ ] **5. Expand optional Markdig rows and benchmarks only after the engine/source contracts settle.**

## Recently Closed

- [x] `UseAutoLinks` moved to covered.
- [x] `UseAbbreviations` moved to covered.
- [x] `UseNonAsciiNoEscape` moved to covered.
- [x] `UsePipeTables` moved to covered.
- [x] `UseTaskLists` moved to covered.
- [x] `UseFootnotes` moved to covered.
- [x] CommonMark `0.31.2` full inventory moved to green.
- [x] `UseGenericAttributes` moved through fenced-code info attributes, standalone attribute blocks before fenced code/headings/paragraphs/root ordered and unordered lists/HTML blocks/dash setext forms/indented code, standalone-before-blockquote literal behavior, common inline elements, reference links/images, angle autolinks, pipe tables, blockquote behavior bounds, list items, task-list interaction, footnotes, standalone-before-footnote-definition consumption, paragraph separator whitespace, no-space bare-URL paragraph attributes, no-space abbreviation-ending paragraph attributes, ordinary no-space plain-text paragraph attributes, raw inline HTML/image edges, emphasis-extra parity bounds, and definition-list interactions.
- [x] `UseDefinitionLists` gained empty-marker blank-separated body source/writer preservation, so `Term`, `:   `, blank, indented body reparses to Markdig-compatible loose definition HTML instead of collapsing the body to a tight definition.
- [x] `UseDefinitionLists` gained blank-separated pre-marker term boundary proof: `Term 1`, blank, `Term 2`, `:   Definition` keeps `Term 1` as a paragraph and starts the Markdig-compatible definition list at `Term 2`.
- [x] `UseDefinitionLists` gained table-shaped continuation profile proof: with pipe tables off, table-looking continuation lines stay literal and write as escaped literal text for stable reparse; with pipe tables on, the same continuation parses as a nested table inside `<dd>`.
- [x] `UseDefinitionLists` gained setext-following lazy-continuation boundary behavior: after a setext heading inside `<dd>`, the following unindented paragraph now closes the definition list like Markdig instead of being swallowed into the definition body.
- [x] `UseDefinitionLists` gained thematic-break-following lazy-continuation boundary behavior and writer preservation: after a thematic break inside `<dd>`, the following unindented paragraph now closes the definition list like Markdig, and paragraph-plus-rule bodies write with a blank separator so reparse does not turn the paragraph into a setext heading.
- [x] `UseDefinitionLists` gained nested-list interruption boundaries: after a nested list inside `<dd>`, unindented ATX headings and thematic breaks now close the definition list like Markdig while preserving definition-body source/native fields and Markdown writer reparse behavior.
- [x] `UseDefinitionLists` gained nested-blockquote interruption boundaries: after a nested blockquote inside `<dd>`, unindented ATX headings and thematic breaks now close the definition list like Markdig while preserving definition-body source/native fields and Markdown writer reparse behavior.
