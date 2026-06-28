# OfficeIMO.Markdown Markdig Parity Gap Plan

This is the working plan for getting `OfficeIMO.Markdown` to Markdig-class parity without looping through disconnected fixture additions.

The important distinction: parity is not "more tests." Parity means the parser, AST, renderer, writer, extension model, source mapping, lossless behavior, and performance story all line up behind a documented contract. Tests and inventories are the proof gates for those contracts.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 316 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 652 of 652 official CommonMark `0.31.2` examples currently match; 0 are failing in `Docs/officeimo.markdown.commonmark-inventory.md` |
| GFM corpus | 44 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| GFM tracked inventory | 44 tracked GFM fixtures in `Docs/officeimo.markdown.gfm-inventory.md`: 40 upstream cmark-gfm fixtures, 4 OfficeIMO supplements, 44 passing, 0 failing |
| Markdig extension inventory | 33 Markdig extension-family rows in `Docs/officeimo.markdown.markdig-extension-inventory.md`: 7 covered, 11 partial, 4 intentional, 11 gap |
| Covered CommonMark sections | ATX headings, Setext headings, thematic breaks, indented code blocks, fenced code blocks, HTML blocks, block quotes, list items, lists, paragraphs, hard breaks, soft breaks, links, images, autolinks, raw HTML, backslash escapes, entity and numeric character references, link reference definitions, tabs |
| Remaining CommonMark parser clusters | None in the official CommonMark `0.31.2` inventory |
| Remaining Markdig-class architecture gaps | broader GFM corpus coverage, full lossless trivia capture, full parser pipeline parity, renderer/writer plugin parity, extension-family implementation breadth, release-mode benchmark review |

## Current Answer

We are not doing only tests. Tests are the measuring system. Parity work means improving the reusable `OfficeIMO.Markdown` engine and then using inventories, fixtures, native snapshots, renderer checks, writer checks, and benchmarks to prove the contract moved.

Current truth:

- [x] CommonMark, tracked GFM fixtures, and reflected Markdig extension families are measurable from checked-in reports.
- [x] GFM smoke behavior is green for the fixture corpus we track today.
- [x] CommonMark official inventory is closed: 0 official examples are failing.
- [ ] Markdig extension parity is not closed: 7 extension families meet the full `Covered` bar; the remaining extension families still need implementation, proof, or intentional-out-of-scope decisions.
- [ ] AST/source/lossless parity is not closed: full trivia, source edits, generated-node diagnostics, and source-aware extension paths are still partial.
- [ ] Performance parity is not known: release-mode Markdig comparisons still need a stable benchmark pass after correctness stops moving.

## Parity Closure Checklist

This is the short board to keep us out of loops. A row is not done because it has tests; a row is done when the reusable engine behavior and its proof are both complete.

- [x] **CommonMark correctness baseline:** all 652 official CommonMark `0.31.2` examples currently match.
- [x] **Core GFM features already promoted:** pipe tables, task lists, footnotes, strikethrough, auto identifiers, soft-line-as-hard-line option, and current tracked GFM smoke fixtures are green.
- [x] **Scoreboards exist:** CommonMark inventory, GFM inventory, Markdig extension inventory, compatibility matrix, and this plan are checked in.
- [x] **Finish `UseEmphasisExtras`:** strikethrough, inserted text, highlight/mark, superscript, and subscript now have parser, AST/source/native, HTML render, Markdown write, Markdig comparison, and explicit GFM single-tilde profile evidence.
- [ ] **Close `UseAutoLinks`:** finish extended bare URL, `www`, email, scheme, boundary, punctuation, Unicode, table-cell, writer, and source/native evidence before promoting from `Partial`.
- [ ] **Close raw HTML and GFM tag-filter separation:** keep CommonMark raw HTML parsing, GFM tag filtering, sanitizer/escape/strip/allow behavior, URL policy, source metadata, and writer behavior as separate contracts.
- [ ] **Close `UseDefinitionLists`:** finish source-map and writer edge breadth for marker groups, lazy continuation, nested blocks, loose definitions, empty markers, and reparsing.
- [ ] **Close high-value partial Markdig rows:** work through `UseAlertBlocks`, `UseGenericAttributes`, `UsePreciseSourceLocation`, and parser/render extension rows with the same engine-plus-proof bar.
- [ ] **Make scope decisions for gap rows before coding them:** `UseCustomContainers`, `UseGridTables`, `UseSmartyPants`, `UseCitations`, `UseMathematics`, `UseMediaLinks`, `UseDiagrams`, `UseFigures`, `UseListExtras`, and similar rows must be assigned to core, optional extension, renderer/host policy, or intentional out-of-scope.
- [ ] **Finish canonical AST cleanup:** remove duplicated/adapted node shapes so semantic blocks, syntax nodes, native snapshots, renderer contexts, writer contexts, and source edits use one ownership model.
- [ ] **Finish lossless/source architecture:** complete trivia capture, delimiter-token capture, original-to-normalized mapping, generated-node diagnostics, caret/source-edit coverage, and source-preserving roundtrip fallbacks.
- [ ] **Finish renderer/writer extension parity:** custom nodes must render and write through source-aware extension contracts without downstream string rescanning.
- [ ] **Run performance proof last:** after correctness and source behavior settle, capture release-mode Markdig comparisons for parse, parse-with-syntax, render, write, transforms, source edits, and allocations.

Immediate execution queue:

- [x] **1. Commit the superscript implementation slice.** This was engine work, not just tests.
- [x] **2. Decide the subscript/profile rule.** Markdig-style subscript uses `~sub~`; GFM keeps single-tilde as strikethrough by disabling subscript in that profile.
- [x] **3. Finish and promote `UseEmphasisExtras`.**
- [ ] **4. Pick the next partial row from the Markdig inventory, not from nearby tests.** Recommended order: autolinks, raw HTML/tag-filter/security split, definition lists, generic attributes/source-location.
- [ ] **5. Do one scope pass over the gap rows before large feature work.** The output should be decisions first, implementation second.

## Missing Work Plan

This is the non-looping backlog. Parity slices are grouped by what they actually move. Test-only work is allowed only when it creates a missing scoreboard, proves a newly fixed contract, or documents an intentional Markdig difference.

### A. Engine And Parser Behavior

- [ ] **GFM breadth is still thin.** The current GFM inventory is green, but only 44 tracked fixtures are imported. Missing work: broaden autolinks, strikethrough delimiter edges, HTML tag filtering, and extension-interaction fixtures against upstream-compatible behavior.
- [ ] **Autolinks are still partial.** CommonMark angle autolinks are green, Markdig-style previous-character/domain-without-period/query-fragment options exist, bare `ftp://` and `tel:` scheme autolinks now have Markdig/source/writer evidence, and source-backed Markdown writing preserves parsed bare and angle autolink spelling, but `UseAutoLinks` remains partial because extended bare URL, `www`, plain-email, scheme, boundary, punctuation, and Unicode edge breadth still need broader Markdig/GFM evidence before promotion.
- [ ] **Raw HTML and GFM tag filtering are still partial.** CommonMark raw HTML is green and cmark-gfm HTML output now has a first-class `HtmlOptions` profile, but broader GFM tag-filter corpus coverage, sanitizer/escape/strip/allow mode evidence, source/writer behavior, and URL policy still need to stay separated so parser parity is not confused with security policy.
- [ ] **Definition-list syntax breadth is partial.** OfficeIMO now parses the pinned Markdig colon-marker form, including multiple terms, multiple definitions, grouped AST/source/native/html proof, Markdig lazy paragraph and nested block continuation, loose-definition HTML, edge-continuation comparison, empty-marker first-continuation source mapping, grouped Markdown writing that preserves the marker form for reparsing, loose-definition writer preservation, and blank-separated marker-group writer preservation. Remaining source-map and writer edge breadth still need focused comparison before `UseDefinitionLists` can move to `Covered`.
- [x] **Emphasis extras are covered.** Strikethrough, inserted text, mark/highlight, superscript, and subscript have first-class parser/source/native/render/write coverage, with GFM single-tilde strikethrough kept explicit through profile settings.

### B. Markdig Extension Scope Decisions

- [ ] **Markdig extension-family coverage is far from closed.** The current inventory is 7 `Covered`, 11 `Partial`, 4 `Intentional`, and 11 `Gap`. Every non-covered row needs one decision: implement in core, implement as optional extension, route to renderer/host policy, or mark intentional out of scope.
- [ ] **High-priority partial rows need closure.** Work through `UseAutoLinks`, `UseDefinitionLists`, `UseAlertBlocks`, `UseGenericAttributes`, `UsePreciseSourceLocation`, and parser/render extensions with parser, AST/source, renderer, writer, and fixture evidence.
- [ ] **High-priority gap rows need scope decisions before implementation.** Decide whether `UseCustomContainers`, `UseGridTables`, `UseSmartyPants`, `UseCitations`, `UseMathematics`, `UseMediaLinks`, `UseDiagrams`, `UseFigures`, `UseListExtras`, and similar rows belong in core, optional packages, renderer policy, or intentional differences.
- [ ] **Abbreviation parity is partial, not closed.** `UseAbbreviations` now has opt-in parser, semantic AST, syntax/native metadata, HTML rendering, source-edit, and selected Markdig comparison evidence, but still needs broader edge cases and a decision on definition-preserving Markdown writer reconstruction before promotion.

### C. AST, Source, And Lossless Architecture

- [ ] **Canonical AST cleanup is incomplete.** Continue removing duplicated or adapter-heavy node shapes, especially around `ListItem`, `TableBlock`, `CalloutBlock`, `DefinitionListBlock`, front matter, and extension-owned nodes, so semantic blocks, syntax nodes, native snapshots, renderer contexts, writer contexts, and source edits agree on one boundary model.
- [ ] **Lossless trivia capture is incomplete.** Parser-owned data still needs full whitespace, blank-line, tab, delimiter, marker, raw slice, normalized text, and generated-node diagnostic coverage before editor-grade roundtrip can be claimed.
- [ ] **Delimiter-token coverage is incomplete.** Inline and block token spans are much better now, but parity still requires complete source tokens for every editor-addressable spelling: emphasis extras, links/images, escapes/entities, hard/soft breaks, HTML tags, footnotes, front matter, tables, and extension inlines/blocks.
- [ ] **Original-to-normalized source mapping is incomplete.** CRLF/LF/CR inputs, tab expansion, nested containers, transformed/generated nodes, and normalized paragraph text need one reliable mapping story with diagnostics when exact mapping is impossible.
- [ ] **Roundtrip editing is not yet Markdig-class lossless mode.** `MarkdownRoundtripWriter` handles unchanged documents and explicit native source edits, but broader source-preserving edits, writer fallback diagnostics, and extension-node roundtrip need to be finished.

### D. Renderer, Writer, Security, And Performance

- [ ] **Renderer/writer extension APIs are still partial.** Custom containers, alerts, diagrams, math, attributes, media links, and other in-scope extension nodes must render and write from source-aware contracts without downstream string rescanning.
- [ ] **Security profiles are not fully separated from parser parity.** Raw HTML allow/strip/escape/sanitize/GFM-tag-filter behavior and URL policy need independent tests and docs so security choices do not masquerade as Markdown grammar.
- [ ] **Performance parity is unproven.** Run release-mode benchmarks only after correctness stabilizes enough to compare parse, parse-with-syntax, HTML render, Markdown write, transforms, source edits, allocations, and representative README/docs/chat corpora against Markdig.

## Execution Order

Use this order to avoid looping:

- [ ] **1. Pick one scoreboard row.** The next slice must name exactly one primary row: GFM breadth, a Markdig extension family, AST/source/lossless, renderer/writer, security, or performance.
- [ ] **2. If behavior is missing, improve the engine first.** Parser, AST, source mapping, renderer, writer, or extension APIs move before fixtures are promoted.
- [ ] **3. If behavior exists but is unproven, add focused proof.** This is the test-only lane: Markdig comparison cases, inventory rows, native snapshots, writer checks, or renderer checks.
- [ ] **4. Promote only when the whole row is covered.** A row moves to `Covered` only with parser behavior, semantic AST/source/native projection where applicable, HTML rendering, Markdown writing or explicit writer limits, fixture/inventory evidence, and profile documentation.
- [ ] **5. Make scope decisions before large new features.** Grid tables, custom containers, math, diagrams, attributes, SmartyPants, citations, media links, and similar rows should not be half-added without deciding core versus optional extension versus renderer policy. Abbreviations already have an in-core partial implementation and should now be completed as a writer/edge-breadth slice.
- [ ] **6. Benchmark last.** Do not optimize or claim performance parity until correctness, source mapping, and writer behavior are stable enough for the numbers to mean something.

## Parity Work Board

Use this as the non-looping execution board. Each item must either move engine behavior, mark a deliberate out-of-scope choice, or improve a scoreboard that is missing.

### Now: Close CommonMark Correctness

- [x] **HTML block/container boundary:** official CommonMark example #174 is pinned and passing after blockquote-contained raw HTML block rendering started preserving the boundary line break before the quote closes.
- [x] **CommonMark character references:** official CommonMark examples #25 and #26 are pinned and passing after reusable named/numeric decoding started handling the remaining HTML5 named references and invalid numeric replacement behavior.
- [x] **Hard-line-break grammar:** official CommonMark examples #642, #643, and #644 are pinned and passing after paragraph joining stopped treating markers inside raw inline HTML tags and final-line trailing backslashes as hard breaks.
- [x] **Emphasis delimiter slice:** official CommonMark examples #408, #438, #441, #450, #453, and #470 are pinned and passing after delimiter-run handling learned root dual italic runs, empty opposite-marker spans, and mixed-marker literal precedence.
- [x] **CommonMark emphasis inventory:** official CommonMark examples #418 and #432 are pinned and passing after same-marker emphasis handling started preferring immediate double closers and nested single-marker spans before splitting double runs into dual italic.
- [x] **Container indentation model:** tabs, blockquote continuation, list continuation, and indented-code boundaries now share corrected visual-column behavior for the official CommonMark cases.
  - [x] **#9 tabs and nested list columns:** tab residual columns are preserved when stripping container indentation, so the third item nests under `bar` instead of becoming its sibling.
  - [x] **#111 indented-code blank-line preservation:** whitespace-only blank lines stay inside one indented code block when a later indented line continues the block.
  - [x] **#231 indented blockquote precedence:** four-space-indented `>` lines are indented code, not block quotes.
  - [x] **#242 adjacent blockquote separation:** `> foo`, blank line, `> bar` becomes two blockquote blocks.
  - [x] **#252 blockquote inner indentation threshold:** quote content strips only the marker/optional following space plus paragraph indentation, keeping `>     code` as code and `>    not code` as paragraph text in a separate quote.
  - [x] **#264 list-contained indented-code continuation:** blank lines stay inside a list item's nested indented code block when a later line still meets the continuation indent.
  - [x] **Source-map guard:** the six official examples are pinned with top-level syntax-tree assertions and the existing invariant checks.
- [x] **CommonMark inventory closure:** the generated full-corpus inventory is refreshed and the newly understood official examples are pinned as smoke fixtures.

### Next: Broaden GFM And Markdig Extension Coverage

- [ ] **GFM fixture breadth:** expand beyond the current 44 tracked fixtures for autolinks, strikethrough, tag filtering, and extension interactions.
- [x] **Pipe tables:** moved from partial support to covered support by proving malformed delimiters, alignment, containers, source spans, renderer output, and writer behavior.
- [x] **Task lists:** moved from partial support to covered support by proving nested markers, exact marker source spans, native snapshots/source edits, renderer output, and ordered/unordered writer behavior.
- [x] **Footnotes:** moved from partial support to covered support by proving Markdig/GFM breadth, label/body source mapping, renderer output, backlink behavior, and writer behavior.
- [x] **Soft line break as hard line break:** moved from missing support to covered support by adding an explicit reader option with HTML output, Markdown writing, nested paragraph propagation, and native metadata proof that synthetic hard breaks do not claim a fake marker.
- [x] **Auto identifiers:** moved from missing support to covered support by proving automatic heading ids, disable behavior, Markdig default and GitHub slug styles, duplicate tracking, GFM profile wiring, and existing heading source/native metadata.
- [x] **YAML front matter:** moved from partial support to covered support by preserving raw YAML as the AST payload, keeping structured helpers for simple entries, exposing body/fence/key/value source spans through syntax/native snapshots, omitting front matter from HTML, and preserving the raw body through Markdown writing.
- [ ] **Autolinks and tag filter:** separate CommonMark autolinks, GFM extended autolinks, and GFM tag-filter behavior into explicit parser/render/security contracts.
- [ ] **Extension-family decisions:** for every `Partial` or `Gap` row in `Docs/officeimo.markdown.markdig-extension-inventory.md`, choose one outcome: implement in core, implement as optional extension, route to renderer/host policy, or mark intentional out of scope.

### Next: Finish AST, Source, And Lossless Claims

- [ ] **Canonical node model:** finish cleanup for duplicated or adapter-heavy nodes so semantic blocks, syntax nodes, native snapshots, renderer contexts, writer contexts, and source-edit helpers agree on boundaries.
- [ ] **Trivia capture:** capture whitespace, blank lines, tabs, delimiter trivia, and raw source slices in parser-owned data rather than reconstructing them downstream.
- [ ] **Delimiter token capture:** cover all inline delimiter spellings for emphasis, links, images, code spans, escapes, entities, hard breaks, HTML tags, footnotes, and extension inlines.
- [ ] **Original-to-normalized mapping:** make source spans reliable across CRLF/LF/CR inputs, tab expansion, nested containers, generated nodes, and normalized paragraph text.
- [ ] **Lossless edits:** broaden `MarkdownRoundtripWriter` from unchanged documents and explicit native edits toward general source-preserving edits with precise fallback diagnostics.
- [ ] **Native/document snapshots:** keep every editor-addressable token and canonical node visible through native snapshots, source fields, caret lookup, and source-edit helpers.

### Then: Renderer, Writer, Security, And Performance

- [ ] **Renderer profile parity:** keep CommonMark and GFM HTML output spec-compatible while preserving explicit OfficeIMO profile behavior for raw HTML, images, tables, and document-specific rendering.
- [ ] **Writer parity:** ensure Markdown writing can roundtrip parser-owned syntax and extension nodes without downstream string rescanning.
- [ ] **Extension rendering:** make parser, transform, renderer, and writer extension APIs source-slice aware for custom containers, alerts, diagrams, math, attributes, media links, and other in-scope extension nodes.
- [ ] **Security profiles:** independently test raw HTML allow/strip/escape/sanitize/GFM-tag-filter behavior and URL policy behavior so security choices are not confused with parser parity.
- [ ] **Benchmarks:** capture release-mode Markdig comparisons for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative README/docs/chat corpora.

## Per-Slice Acceptance Checklist

Every parity slice should use this checklist before it is called done.

- [ ] Name the scoreboard row being moved: CommonMark, GFM, Markdig extension family, AST/source/lossless, renderer/writer, security profile, or benchmark proof.
- [ ] Fix the reusable engine/core behavior first when behavior is missing.
- [ ] Add or update focused contract tests only after the behavior contract is understood.
- [ ] Refresh generated inventories when the scoreboard should move.
- [ ] Promote official fixtures only when they prove a stable contract, not merely because they happen to pass.
- [ ] Update the compatibility matrix and this plan when counts, priority, or scope changed.
- [ ] Validate the relevant Markdown lanes across the framework targets affected by the slice.

References:

- CommonMark `0.31.2` official JSON: https://spec.commonmark.org/0.31.2/spec.json
- Markdig README and feature baseline: https://github.com/xoofx/markdig
- Markdig roundtrip/trivia target: https://github.com/xoofx/markdig/blob/master/src/Markdig/Roundtrip.md
- Markdig extension specs index: https://github.com/xoofx/markdig/blob/master/src/Markdig.Tests/Specs/readme.md

## What We Already Have

- [x] A repo-owned CommonMark inventory that runs all 652 official examples without forcing every known failure to fail CI.
- [x] A checked-in failure-cluster report so we can pick work by root cause instead of nearby example numbers.
- [x] A pinned CommonMark smoke lane with 316 official examples.
- [x] CommonMark code spans are green in the generated inventory after preserving NBSP as non-collapsible text in the HTML comparison harness and pinning official examples 333 and 334.
- [x] CommonMark links are green in the generated inventory after fixing link-label inline-span precedence.
- [x] Strong current coverage for headings, thematic breaks, fenced code, lists, paragraphs, soft breaks, backslash escapes, autolinks, images, and link reference definitions.
- [x] A generated GFM inventory tracks the current extension fixture corpus by section and source, separating upstream cmark-gfm fixtures from OfficeIMO supplements.
- [x] A generated Markdig extension inventory tracks reflected Markdig extension-family entry points and classifies current OfficeIMO support as covered, partial, intentional, or gap.
- [x] Syntax/source/native tests for many existing nodes, including source slices and source-edit helpers.
- [x] Markdig package baseline guarded across tests, benchmarks, and docs.

## What We Are Missing

These are the actual parity gaps. The test work is listed only where it creates a missing scoreboard.

- [x] **CommonMark HTML block boundary parity.** HTML blocks, raw HTML, and the blockquote-contained raw HTML boundary case are green in the full inventory.
- [x] **CommonMark entity decoder parity.** Entity and numeric character references are green in the full inventory after shared CommonMark decoding covered the remaining official named/numeric examples.
- [x] **CommonMark hard-break and inline precedence parity.** Hard line breaks are green in the full inventory after preserving raw inline HTML line endings and final-line literal backslashes.
- [x] **Markdig soft-line-as-hard-line option parity.** `UseSoftlineBreakAsHardlineBreak` is covered by an explicit reader option, renderer/writer behavior, nested paragraph propagation, and native metadata evidence.
- [x] **Markdig auto-identifier parity.** `UseAutoIdentifiers` is covered by explicit HTML enablement, Markdig-compatible slug styles, duplicate handling, and GFM profile evidence.
- [x] **CommonMark emphasis inventory parity.** The official CommonMark emphasis section is green in the full inventory; broader Markdig emphasis-extra, source-token, writer, and lossless claims still need their own evidence.
- [x] **CommonMark container indentation parity.** The remaining 6 failures around tabs, blockquote/list continuation, indented-code boundaries, and source-map-safe column handling are closed in the official inventory.
  - [x] Normalize indentation decisions around visual columns for root blocks, nested blocks, and list marker metadata.
  - [x] Make blank-line continuation rules shared between root indented code and nested indented code.
  - [x] Make blockquote parsing respect CommonMark precedence: indented code wins at four columns, quote marker stripping removes at most one following space, and a blank line between quoted paragraphs closes the current quote.
  - [x] Promote #9, #111, #231, #242, #252, and #264 after the engine behavior was fixed and pinned with syntax-tree invariant coverage.
- [ ] **GFM corpus expansion.** Continue extending the generated GFM inventory beyond the current tracked fixture corpus so autolink, strikethrough, tag-filter, and interop behavior are measured against broader upstream-compatible coverage.
- [ ] **Markdig extension implementation breadth.** Move selected partial/gap rows from the generated Markdig extension inventory into real parser, AST/source, renderer, writer, fixture, or intentional-deviation work.
- [ ] **AST/source/lossless completeness.** Finish canonical node cleanup, full trivia capture, delimiter-token capture, original-to-normalized mapping, generated-node diagnostics, and broader byte-preserving source edits.
- [ ] **Renderer/writer extension parity.** Make parser, transform, renderer, and writer extension APIs source-slice aware where custom nodes need to render or roundtrip without string rescanning.
- [ ] **Renderer/security profile parity.** Keep CommonMark/GFM HTML output spec-compatible while making OfficeIMO-specific raw HTML, sanitizing, escaping, URL policy, and GFM tag-filter behavior explicit and independently tested.
- [ ] **Performance proof.** Capture release-mode benchmarks against Markdig for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative README/docs/chat corpora.

## Immediate Work Order

- [x] Finish the remaining HTML block/container boundary case (#174), pin it as a smoke fixture, and refresh the inventory.
- [x] Tackle the entity decoder as a reusable parser service and route the required decode contexts through it.
- [x] Tackle hard breaks as a parser/source-map slice and clear the remaining raw-inline-HTML/final-backslash examples.
- [x] Close the CommonMark emphasis failure cluster while keeping broader Markdig/source/lossless emphasis work explicit.
- [x] Tackle container indentation as a separate source-map column-model slice.
- [x] Refresh the CommonMark inventory and smoke fixture set once container indentation is green.
- [x] Promote `UseSoftlineBreakAsHardlineBreak` only after parser option, renderer, writer, nested paragraph, and native metadata evidence landed.
- [x] Promote `UseAutoIdentifiers` only after explicit renderer options, Markdig/GitHub slug-style evidence, duplicate handling, and GFM profile wiring landed.
- [ ] Use `Docs/officeimo.markdown.markdig-extension-inventory.md` to pick the next extension-family slices by row; do not promote `Partial` rows to `Covered` without parser, AST/source, renderer, writer, and fixture evidence.
- [ ] Return to AST/source/lossless completion once the major parser clusters stop moving node boundaries every slice.
- [ ] Run release-mode benchmarks only after correctness stabilizes enough for the numbers to mean something.

## CommonMark Failure Clusters

| Cluster | Failing | Sections | First examples | Work type |
| --- | ---: | --- | --- | --- |
| None | 0 | None | None | The official CommonMark `0.31.2` inventory is green; remaining parity work is GFM breadth, Markdig extension breadth, AST/source/lossless, renderer/writer, security, and performance |

## Pinned CommonMark Coverage By Section

This is fixture coverage, not a claim that unpinned examples fail. The generated inventory is the pass/fail source of truth.

| Section | Pinned | Total | Missing | Primary missing work |
| --- | ---: | ---: | ---: | --- |
| Emphasis and strong emphasis | 19 | 132 | 113 | CommonMark inventory is green; remaining work is breadth and Markdig/GFM comparison |
| Links | 27 | 90 | 63 | CommonMark inventory is green; remaining work is breadth and Markdig/GFM comparison |
| HTML blocks | 20 | 44 | 24 | Current CommonMark inventory is green; keep broader source-span/writer coverage aligned |
| Raw HTML | 8 | 20 | 12 | Current CommonMark inventory is green; keep broader source-span/writer coverage aligned |
| Images | 2 | 22 | 20 | Breadth and source metadata, not current CommonMark failures |
| Code spans | 8 | 22 | 14 | Current CommonMark inventory is green; keep delimiter/source-token coverage aligned |
| Link reference definitions | 10 | 27 | 17 | Breadth and source metadata, not current CommonMark failures |
| Block quotes | 13 | 25 | 12 | Current CommonMark inventory is green; keep source-span/writer coverage aligned |
| Autolinks | 12 | 19 | 7 | Official CommonMark is green; keep GFM/profile extensions separate |
| Indented code blocks | 2 | 12 | 10 | Current CommonMark inventory is green; keep source-span/writer coverage aligned |
| Tabs | 1 | 11 | 10 | Current CommonMark inventory is green; keep source-span/writer coverage aligned |
| Entity and numeric character references | 9 | 17 | 8 | Current CommonMark inventory is green; keep broader entity/source-span/writer coverage aligned |
| Hard line breaks | 9 | 15 | 6 | Current CommonMark inventory is green; keep marker/source-span/writer coverage aligned |
| List items | 39 | 48 | 9 | Current CommonMark inventory is green; keep source-span/writer coverage aligned |
| Backslash escapes | 8 | 13 | 5 | Breadth, not current CommonMark failures |
| Textual content | 0 | 3 | 3 | Baseline text breadth |
| Blank lines | 0 | 1 | 1 | Baseline block breadth |
| Inlines | 0 | 1 | 1 | Inline precedence breadth |
| Precedence | 0 | 1 | 1 | Parser precedence breadth |
| ATX headings | 18 | 18 | 0 | Covered; keep invariant coverage |
| Fenced code blocks | 29 | 29 | 0 | Covered; keep invariant coverage |
| Lists | 26 | 26 | 0 | Covered; keep invariant coverage |
| Paragraphs | 8 | 8 | 0 | Covered; keep invariant coverage |
| Setext headings | 27 | 27 | 0 | Covered; keep invariant coverage |
| Soft line breaks | 2 | 2 | 0 | Covered; keep invariant coverage |
| Thematic breaks | 19 | 19 | 0 | Covered; keep invariant coverage |

## Detailed Phase Plan

### Phase 0: Inventories And Scoreboards

- [x] CommonMark full-corpus inventory exists.
- [x] CommonMark failures are grouped by root parser cause.
- [x] Add the same generated inventory style for enabled cmark-gfm extension fixtures.
- [x] Add a Markdig comparison inventory that separates OfficeIMO profile differences from portable/CommonMark profile differences.
- [x] Add an extension-family table with `Covered`, `Partial`, `Intentional`, or `Gap` for every reflected Markdig extension family.

Done means:

- [x] We can answer "what is missing?" for CommonMark, tracked GFM fixtures, and Markdig extension parity from checked-in reports. Broader GFM corpus coverage and implementation work behind partial/gap Markdig rows remain open.
- [ ] Every future engine slice names which scoreboard row it moves.

### Phase 1: CommonMark Parser Closure

- [x] HTML/raw HTML: implement the remaining CommonMark HTML block/profile behavior covered by the current inventory.
- [x] Emphasis: close the official CommonMark emphasis inventory; keep Markdig emphasis extras and lossless delimiter-token breadth separate.
- [x] Containers: implement tab expansion and continuation indentation as parser/source-map primitives for the official CommonMark cases.
  - [x] Centralize "visual indentation columns" so root parsers, list parsers, nested block parsers, and source slices agree on spaces versus tabs.
  - [x] Centralize "blank line belongs to code block when a later indented line continues it" for root and nested indented code.
  - [x] Centralize blockquote marker handling: at most three leading indentation columns, one `>` marker, one optional following space, and explicit block termination across blank lines.
  - [x] Re-check list continuation indentation after the quote/code fixes so list items do not absorb sibling containers or split nested code blocks.
- [x] Hard breaks: clear the raw-inline-HTML and final-backslash exclusions without losing marker source spans.
- [x] Code spans: clear the NBSP/equivalence bucket and pin the official cases.
- [x] Entities: clear the remaining official named/numeric character-reference cases with a shared decoder.

Done means:

- [x] The official CommonMark `0.31.2` inventory has no unexplained failures.
- [x] Any intentional OfficeIMO profile differences are documented in the compatibility matrix instead of hidden as failing examples.

### Phase 2: GFM And Markdig Extension Breadth

- [x] Expand GFM tables beyond smoke fixtures, including malformed delimiters, container interactions, source spans, renderer output, and writer behavior.
- [x] Expand GFM task-list coverage, including marker source spans, nested list behavior, native source edits, renderer output, and writer behavior.
- [ ] Expand GFM autolinks and tag-filter coverage.
- [x] Expand GFM footnote coverage, including source spans, nested block bodies, repeated backrefs, renderer output, and writer behavior.
- [ ] Expand GFM strikethrough coverage.
- [ ] Decide which remaining Markdig extensions are in scope: grid tables, emoji, math, diagrams, SmartyPants, citations, custom containers, generic attributes, media links, alerts, advanced links, and list/emphasis extras. Abbreviations are now an in-core partial row that still needs writer and edge-breadth closure.
- [ ] Route in-scope extension work to the right owner: core `OfficeIMO.Markdown`, renderer layer, or separate extension package.
- [ ] Document out-of-scope Markdig extensions as intentional differences.

Done means:

- [x] Enabled tracked GFM fixture behavior is measured by inventory.
- [x] Markdig extension parity is an explicit support matrix, not an implied promise.

### Phase 3: AST, Source Mapping, And Lossless Roundtrip

- [ ] Finish canonical AST cleanup for `ListItem`, `TableBlock`, `CalloutBlock`, `DefinitionListBlock`, and any remaining duplicated mutable node shapes.
- [ ] Complete parser-owned trivia capture for whitespace, blank lines, tabs, delimiter trivia, and generated-node diagnostics.
- [ ] Complete original-to-normalized offset mapping for all parser paths.
- [ ] Expand `MarkdownRoundtripWriter` from unchanged documents and explicit native source edits toward general lossless edits.
- [ ] Keep semantic nodes, syntax nodes, native snapshots, renderer contexts, writer contexts, and source-edit helpers aligned on the same boundaries.
- [ ] Keep public semantic APIs stable or document intentional breaking cleanup before merge.

Done means:

- [ ] Editor-grade claims are backed by native snapshots, source edits, syntax invariants, and roundtrip diagnostics across real mixed documents.

### Phase 4: Renderer, Security, And Performance

- [ ] Keep HTML rendering spec-compatible for CommonMark/GFM profiles.
- [ ] Keep OfficeIMO profile differences explicit for raw HTML, images, tables, and document-specific behavior.
- [ ] Independently test raw HTML allow/strip/escape/sanitize/GFM-tag-filter behavior.
- [ ] Ensure renderer and writer extensions can handle custom nodes without string rescanning.
- [ ] Capture release-mode benchmarks against Markdig for parse, syntax-tree parse, HTML render, Markdown write, source edits, transforms, and allocations.
- [ ] Compare stable README/docs/chat/transcript corpora against the Markdig baseline.

Done means:

- [ ] Compatibility, security, extension, and performance claims can be repeated from documented commands.

## Rules To Stop The Loop

- [x] Do not add an official fixture just because it currently passes; first classify which engine contract it proves.
- [x] Do not add a fixture for a failing example until the engine behavior is fixed or the deviation is documented as intentional.
- [x] Every parity slice must say which bucket it reduces: parser grammar, AST/source mapping, renderer/writer behavior, extension seam, lossless roundtrip, GFM extension parity, or benchmark evidence.
- [x] Each completed slice must update the compatibility matrix count, the gap plan if it changes priority, and the focused test lane.
- [x] If a slice touches parsing, validate at least net8 plus the broad Markdown lane before commit.
- [x] If a slice touches public AST/source/native APIs, add source-span/native/syntax invariant proof, not only HTML output.
- [ ] If a slice only changes tests or comparison tooling, state that plainly and do not call it an engine fix.
- [ ] If a slice changes engine behavior, promote fixtures only after proving the changed behavior against official examples and focused regression tests.

## Final Parity Gates

Do not claim "Markdig-class parity" until all of these are true:

- [ ] Full CommonMark `0.31.2` corpus is imported with pass/fail/intentional-deviation inventory.
- [ ] All parser failures are either fixed or explicitly documented as intentional profile differences.
- [ ] Enabled GFM behavior is covered against cmark-gfm or Markdig-compatible specs.
- [ ] Markdig extension families are inventoried, with in-scope work implemented or planned and out-of-scope work documented.
- [ ] AST/source/native projection coverage exists for every canonical parser node and every editor-addressable token.
- [ ] Lossless/trivia behavior meets the documented roundtrip design or the remaining limits are explicit.
- [ ] Renderer and writer extension APIs can handle custom nodes without string rescanning.
- [ ] Release-mode benchmarks have been captured and reviewed against the local Markdig baseline.
- [ ] Compatibility docs, README-level claims, and package behavior all describe the same current truth.
