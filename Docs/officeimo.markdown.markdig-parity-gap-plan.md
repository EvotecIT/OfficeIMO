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
| Remaining architecture gaps | broader GFM breadth, Markdig extension breadth, canonical AST ownership, full lossless trivia/source mapping, source-aware renderer/writer extension seams, security/profile separation, and release-mode benchmark evidence |

## What Is Missing

- [x] **CommonMark parser correctness is closed.** The official CommonMark `0.31.2` inventory is green: 652 of 652 examples match and 0 are failing.
- [x] **Core GFM behavior is real engine behavior.** Pipe tables, task lists, footnotes, strikethrough, auto identifiers, extended autolinks, soft-line-as-hard-line, YAML front matter, abbreviations, and tracked GFM fixtures have parser/render/write/source proof.
- [x] **The scoreboards exist.** CommonMark inventory, generated GFM inventory, Markdig extension inventory, compatibility matrix, benchmark hooks, and this gap plan are checked in.
- [ ] **Markdig extension parity is not closed.** The current inventory has 10 covered, 9 partial, 3 intentional, and 11 gap rows.
- [ ] **AST/source/lossless parity is not closed.** Full trivia capture, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, broader source edits, and extension-node roundtrip still need work.
- [ ] **Performance parity is not known.** Release-mode benchmark comparisons should run after parser/source/writer behavior stops moving.

## Missing Work By Kind

- [ ] **Engine work: finish real Markdown grammar gaps.**
  - [ ] `UseGenericAttributes`: arbitrary block-family attachment, every Markdig-supported inline-family target, no-space bare-URL behavior, and writer/source preservation across those shapes.
  - [ ] `UseCustomContainers`: colon-fenced container parsing, nested child-block source ownership, renderer/writer seams, and Markdig comparison fixtures.
  - [ ] `UseGridTables`: grid-table parser, semantic table model, malformed fallback, source spans, HTML rendering, and Markdown writer behavior.
  - [ ] `UseListExtras`: inventory Markdig list-extra syntax first, then decide whether it is core list behavior or an optional parser extension.
  - [ ] `UseCjkFriendlyEmphasis`: Markdig-compatible delimiter option with CJK comparison and delimiter source-token proof.

- [ ] **Engine-plus-proof work: promote partial rows that already mostly exist.**
  - [ ] `UseDefinitionLists`: close remaining source-map and writer edge breadth for marker groups, lazy continuation, loose definitions, nested bodies, multiline bodies, and reparse stability.
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

- [ ] **Finish `UseGenericAttributes`.**
  Covered now: fenced code, headings, paragraphs, consumed paragraph separator whitespace, thematic-break-like attributed paragraphs including dash/asterisk/underscore spellings, root/nested/blockquote list items, task-list interaction, pipe-table promotion, definition-list term projection, definition-list definition-value consumption, inline links/reference links/images/reference images/linked images/common emphasis/code spans/angle autolinks/superscript/subscript, raw inline HTML consumption, footnote-definition paragraph attributes, footnote-reference attribute consumption, and Markdig-compatible literal handling for strikethrough/highlight/inserted emphasis-extra attribute blocks.
  Missing before promotion:
  - [ ] Probe the remaining Markdig-supported block and inline shapes before writing code.
  - [ ] Implement only shapes Markdig actually consumes or projects as generic attributes.
  - [ ] Use shared attribute parser/writer helpers instead of per-block string rescans.
  - [ ] Project each newly supported shape through semantic AST, syntax AST, native source fields/metadata, HTML rendering, Markdown writing, and preserved-trivia source edits.
  - [ ] Document unsupported or intentional differences instead of leaving silent gaps, including no-space bare-URL paragraph attributes.

## P1 - Markdig Extension Rows

- [ ] **Promote or explicitly bound `UseDefinitionLists`.**
  Missing: remaining marker-group, lazy-continuation, loose-definition, nested-definition, multiline-body, source-map, writer, and reparse-stability edge breadth.
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

- [ ] **1. Continue `UseGenericAttributes`, but only after probing actual missing Markdig behavior.**
  The most recent simple HTML-block probe passed, so it is not an engine gap by itself.
- [ ] **2. Promote or bound `UseDefinitionLists`.**
  This is the next concrete parser/source/writer row after generic attributes.
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
- [x] `UseGenericAttributes` moved through fenced code, headings, paragraphs, common inline elements, reference links/images, angle autolinks, pipe tables, blockquote behavior bounds, list items, task-list interaction, footnotes, paragraph separator whitespace, raw inline HTML/image edges, emphasis-extra parity bounds, and definition-list interactions.
