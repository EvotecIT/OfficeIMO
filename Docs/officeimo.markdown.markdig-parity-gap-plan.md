# OfficeIMO.Markdown Markdig Parity Gap Plan

This is the working plan for getting `OfficeIMO.Markdown` to Markdig-class parity without looping through disconnected fixture additions.

The important distinction: parity is not "more tests." Parity means the parser, AST, renderer, writer, extension model, source mapping, lossless behavior, and performance story all line up behind a documented contract. Tests and inventories are the proof gates for those contracts.

## Operator Plan

This is the short checklist to use before starting the next slice. If a task cannot name one of these rows, it is probably scope drift.

- [x] **CommonMark parser correctness.** The official CommonMark `0.31.2` inventory is green: 652 of 652 examples match.
- [x] **Core scoreboards.** CommonMark, tracked GFM, Markdig extension-family inventory, compatibility docs, and benchmark hooks exist.
- [x] **Covered GFM/core features.** Pipe tables, task lists, footnotes, strikethrough/emphasis extras, auto identifiers, extended autolinks, soft-line-as-hard-line, YAML front matter, abbreviations, and current tracked GFM fixtures have engine behavior plus proof.
- [ ] **Engine slice 1: continue `UseGenericAttributes` breadth.** List-item trailing attributes are now closed for root, nested, ordered, unordered, blockquote-contained, and task-list interaction cases; paragraph separator whitespace, thematic-break-like attributed paragraphs, footnote-definition paragraphs, inline images, linked images, and raw-inline-HTML consumption proof are also in place. Next probe the remaining block/container shapes and fix only supported shapes in parser, AST, syntax/native source fields, HTML render, Markdown writer, and reparse/source-edit proof.
- [ ] **Engine slice 2: finish or explicitly bound `UseDefinitionLists`.** Close remaining source-map and writer edges for marker groups, lazy continuation, loose definitions, nested definitions, multiline body edits, and reparse stability.
- [ ] **Decision slice 3: close `UseAlertBlocks` and `UseCjkFriendlyEmphasis`.** Decide whether each is an OfficeIMO compatibility contract, renderer/profile policy, deferred row, or intentional difference before adding more fixtures.
- [ ] **Architecture slice 4: finish source/lossless parity.** Complete canonical node shapes, trivia capture, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, and broader source-preserving edits.
- [ ] **Extension slice 5: implement only chosen optional Markdig rows.** Custom containers, grid tables, math, diagrams, figures, media links, emoji/smiley, Jira links, list extras, and SmartyPants need explicit ownership and promotion bars before implementation.
- [ ] **Policy slice 6: separate parser parity from security/render policy.** Raw HTML, GFM tag filtering, sanitizing/escaping/stripping, URL policy, referral links, and theme/bootstrap behavior stay independent contracts.
- [ ] **Proof-only slice 7: broaden fixtures and benchmarks after engine contracts stabilize.** Expand GFM fixture breadth and run release-mode Markdig benchmarks only after parser/source/writer behavior stops moving.

Rule of thumb:

- [ ] If behavior is missing, change the reusable engine first.
- [ ] If behavior exists but is not proven, add focused comparison/source/writer proof.
- [ ] If the row is optional, deferred, or policy-only, document that decision before writing parser code.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 316 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 652 of 652 official CommonMark `0.31.2` examples currently match; 0 are failing in `Docs/officeimo.markdown.commonmark-inventory.md` |
| GFM corpus | 52 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| GFM tracked inventory | 52 tracked GFM fixtures in `Docs/officeimo.markdown.gfm-inventory.md`: 48 upstream cmark-gfm fixtures, 4 OfficeIMO supplements, 52 passing, 0 failing |
| Markdig extension inventory | 33 Markdig extension-family rows in `Docs/officeimo.markdown.markdig-extension-inventory.md`: 10 covered, 9 partial, 3 intentional, 11 gap |
| Covered CommonMark sections | ATX headings, Setext headings, thematic breaks, indented code blocks, fenced code blocks, HTML blocks, block quotes, list items, lists, paragraphs, hard breaks, soft breaks, links, images, autolinks, raw HTML, backslash escapes, entity and numeric character references, link reference definitions, tabs |
| Remaining CommonMark parser clusters | None in the official CommonMark `0.31.2` inventory |
| Remaining Markdig-class architecture gaps | broader GFM corpus coverage, full lossless trivia capture, full parser pipeline parity, renderer/writer plugin parity, extension-family implementation breadth, release-mode benchmark review |

## Current Answer

We are not doing only tests. Tests are the measuring system. Parity work means improving the reusable `OfficeIMO.Markdown` engine and then using inventories, fixtures, native snapshots, renderer checks, writer checks, and benchmarks to prove the contract moved.

Current truth:

- [x] CommonMark is closed for the official `0.31.2` inventory: 652 of 652 examples currently match.
- [x] Core GFM features are real engine behavior now, not just fixtures: pipe tables, task lists, footnotes, strikethrough, auto identifiers, soft-line-as-hard-line, front matter, and current tracked GFM smoke fixtures are green.
- [x] The comparison system exists: CommonMark inventory, GFM inventory, Markdig extension inventory, compatibility matrix, benchmarks, and this plan are checked in.
- [ ] Markdig extension parity is not closed: 10 extension families are `Covered`, 9 are `Partial`, 3 are `Intentional`, and 11 are `Gap`. The 11 `Gap` rows now have explicit scope decisions in the generated Markdig extension inventory, but their behavior is still not implemented unless they are deferred or intentional.
- [ ] AST/source/lossless parity is not closed: full trivia, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, source edits, and source-aware extension paths are still partial.
- [ ] Performance parity is not known: release-mode Markdig comparisons still need a stable pass after correctness and source behavior stop moving.

## Missing Parity Plan

This is the non-looping backlog. A row can move to done only when reusable engine behavior, public contract, docs, and proof are all present. Test-only work is valid only when the row says `Proof only` or the engine behavior already exists and lacks evidence.

The short answer to "are we only doing tests?" is no:

- Engine work means parser, semantic AST, syntax AST, source/native projection, renderer, writer, or extension APIs change.
- Proof work means Markdig comparison cases, generated inventories, snapshots, writer checks, renderer checks, and benchmarks prove the engine contract.
- Parity needs both, but proof does not replace missing engine behavior.

### P0 - Current Source Of Truth

- [x] **CommonMark correctness is closed.** The official CommonMark `0.31.2` inventory is green: 652 of 652 examples currently match.
- [x] **The parity measuring system exists.** CommonMark inventory, GFM inventory, Markdig extension inventory, compatibility docs, and benchmark hooks are checked in.
- [x] **Gap rows are classified before implementation.** Every reflected Markdig extension row now has a scope decision in `Docs/officeimo.markdown.markdig-extension-inventory.md`.
- [ ] **Markdig extension parity is not closed.** Current count: 10 covered, 9 partial, 3 intentional, 11 gap.
- [ ] **AST/source/lossless parity is not closed.** Full trivia, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, source edits, and source-aware extension paths are still partial.
- [ ] **Performance parity is not known.** Release-mode Markdig comparisons should run after parser/source/writer behavior stops moving.

### P1 - Engine Work We Are Missing Next

- [ ] **Finish `UseGenericAttributes` breadth.** This is the active engine slice. Already covered: fenced code, ATX headings, Setext headings, paragraphs with Markdig-compatible consumed separator whitespace, thematic-break-like attributed paragraph lines, root/nested/blockquote/task-list list-item trailing attributes, footnote-definition body paragraphs, Markdig-style pipe-table cell attributes that promote to the owning table, linked-image and image inline attributes, raw-inline-HTML attribute-block consumption, and common no-space inline attributes. Missing before promotion:
  - [ ] Probe and close remaining block families: thematic breaks, table-edge forms, HTML blocks, footnotes, definition lists, callouts, standalone media/image shapes, and any list-extra forms Markdig treats differently from ordinary list items.
  - [ ] Implement only the shapes Markdig actually treats as generic attributes, using shared parser helpers instead of per-block string rescans.
  - [ ] Project every newly supported attribute through semantic AST, syntax AST, native fields/metadata, HTML rendering, Markdown writing, and preserved-trivia source edits.
  - [ ] Keep unsupported or intentionally different shapes documented as profile differences, not silent gaps.
- [ ] **Finish `UseDefinitionLists`.** Engine plus proof. Remaining work is source-map and writer edge breadth for marker groups, lazy continuation, loose definitions, nested definitions, multiline body edits, and reparse stability.
- [ ] **Decide and close `UseAlertBlocks`.** Choose whether Markdig alert rendering callbacks become an OfficeIMO renderer contract or remain an intentional OfficeIMO callout difference, then implement or document the chosen path.
- [ ] **Decide and close `UseCjkFriendlyEmphasis`.** Either add a real delimiter option with CJK comparison fixtures and source-token proof, or document it as deferred/intentional.
- [ ] **Finish raw HTML, GFM tag filtering, and security policy as separate contracts.** Keep CommonMark raw HTML grammar, cmark-gfm tag filtering, OfficeIMO allow/strip/escape/sanitize modes, URL policy, source metadata, and Markdown writing independently testable.

### P2 - AST, Source, And Lossless Work We Are Missing

- [ ] **Finish `UsePreciseSourceLocation`.** Cross-cutting source architecture. Complete lossless trivia/original mapping, generated-node diagnostics, and source-edit coverage before claiming parity.
- [ ] **Canonicalize remaining AST shapes.** Remove duplicated or adapter-heavy ownership around `ListItem`, `TableBlock`, `CalloutBlock`, `DefinitionListBlock`, front matter, and extension nodes.
- [ ] **Complete lossless trivia capture.** Capture whitespace, blank lines, tabs, delimiter trivia, raw slices, normalized text, and generated-node diagnostics in parser-owned data.
- [ ] **Complete delimiter-token coverage.** Every editor-addressable spelling needs source tokens: emphasis extras, links/images, escapes/entities, hard/soft breaks, HTML tags, footnotes, front matter, tables, and extension inlines/blocks.
- [ ] **Complete original-to-normalized mapping.** CRLF/LF/CR inputs, tab expansion, nested containers, transformed/generated nodes, and normalized paragraph text need one reliable mapping story with diagnostics when exact mapping is impossible.
- [ ] **Expand roundtrip editing beyond unchanged documents.** `MarkdownRoundtripWriter` handles unchanged documents and explicit native edits; broader source-preserving edits, writer fallback diagnostics, and extension-node roundtrip remain open.

### P3 - Extension And Optional Feature Work We Are Missing

- [ ] **Finish parser/render extension seams.** Custom block, inline, transform, renderer, and writer extension APIs need source slices and token metadata so extension nodes can render and roundtrip without downstream string rescanning.
- [ ] **Implement `UseCustomContainers` if kept in core.** Needs block parser contract, nested child source mapping, renderer/writer source slices, and Markdig fixtures.
- [ ] **Close high-value renderer/host rows only with explicit ownership.** `UseDiagrams`, `UseFigures`, `UseMathematics`, and `UseMediaLinks` need parser/AST/source/renderer/writer promotion bars before implementation.
- [ ] **Keep optional transform rows optional.** `UseGridTables`, `UseListExtras`, `UseEmojiAndSmiley`, `UseJiraLinks`, and `UseSmartyPants` need separate optional parser/transform contracts.
- [ ] **Keep deferred rows deferred until a consumer needs them.** `UseCitations`, `UseFooters`, `UseGlobalization`, and `UsePragmaLines` need real product requirements before implementation.
- [ ] **Keep renderer-policy rows out of parser parity.** `UseReferralLinks` and similar rows should remain opt-in render policies with safe defaults.

### P4 - Proof-Only Work We Are Missing

- [ ] **Broaden the GFM fixture inventory.** Proof only unless mismatches appear. Current tracked GFM inventory is green but small at 52 fixtures; broaden autolinks, strikethrough delimiter edges, tag filtering, tables, task lists, footnotes, and extension interactions.
- [ ] **Refresh Markdig extension rows after each engine slice.** Proof and docs. Update status, scope decision, route, promotion bar, OfficeIMO state, and next action when behavior changes.
- [ ] **Run release-mode benchmarks last.** Proof only. Compare parse, parse-with-syntax, HTML render, Markdown write, transforms, source edits, allocations, and representative README/docs/chat corpora against the pinned Markdig baseline.

### Next Ordered Slices

- [ ] **1. Continue the active `UseGenericAttributes` slice.** List-item trailing attributes are closed; next probe Markdig's remaining supported shapes, then implement the missing reusable parser/source/writer behavior for those block/inline families.
- [ ] **2. Promote or bound `UseDefinitionLists`.** Close the remaining source-map/writer edge cases or document precise writer/source limits.
- [ ] **3. Decide `UseAlertBlocks` and `UseCjkFriendlyEmphasis`.** Make the scope decision explicit before adding more fixture cases.
- [ ] **4. Return to source/lossless architecture.** Canonical node shapes, trivia, delimiter tokens, original mapping, and broader roundtrip edits are the next big parity body of work.
- [ ] **5. Only then broaden optional extension families and benchmarks.** Optional Markdig rows and release-mode performance numbers should follow stable engine/source contracts.

### Done Recently

- [x] **Promoted `UseAutoLinks`.** Parser, render, write, source/native, profile, and Markdig comparison lanes are covered for the Markdig extension row.
- [x] **Promoted `UseAbbreviations`.** Parser, semantic AST, syntax/native metadata, source-edit, renderer, writer, and selected Markdig comparison coverage are in place.
- [x] **Promoted `UseNonAsciiNoEscape`.** Renderer output paths now route text-bearing HTML text/attribute output through `HtmlOptions.EscapeNonAsciiText`; URL-bearing attributes stay routed through URL encoding.
- [x] **Moved `UseGenericAttributes` through fenced-code rendering.** Shared `MarkdownAttributeSet` storage now exists on semantic `MarkdownObject` nodes and `MarkdownSyntaxNode` nodes, and parsed fenced-code attributes render on default fenced HTML; arbitrary parser/write/source coverage remains open.
- [x] **Moved `UseGenericAttributes` through opt-in heading/paragraph trailing blocks.** `MarkdownReaderOptions.GenericAttributes` now parses Markdig-style trailing attribute blocks on ATX headings, Setext headings, and paragraphs into the shared semantic/syntax attribute storage, renders them to HTML, and writes normalized trailing attribute blocks that reparse stably.
- [x] **Moved `UseGenericAttributes` through common inline elements.** No-space generic attribute blocks now attach to links, images, emphasis, strong, code spans, strikethrough, highlight, inserted, superscript, and subscript nodes; the covered shapes flow through semantic AST, syntax AST, default HTML rendering, Markdown writing, Markdig comparison cases, and reparse proof.
- [x] **Moved `UseGenericAttributes` through native source edits for covered shapes.** Covered generic attribute blocks are captured with original source text and source spans, projected as native block `attributes` source fields and inline `attributes` metadata, included in snapshots, and editable losslessly in preserved-trivia roundtrips.
- [x] **Moved `UseGenericAttributes` through syntax-token coverage for covered shapes.** Covered block and inline attribute blocks now appear as source-addressable `GenericAttributeBlock` syntax tokens, caret/navigation can land on them, and native inline metadata remains single-projected from the token.
- [x] **Moved `UseGenericAttributes` through Markdig-style pipe-table cells.** Trailing generic attribute blocks in header or body cells are consumed from the cell text, promoted to the owning `TableBlock`, rendered on `<table>`, written back into a stable reparsable table cell form, and exposed through syntax/native/source-edit proof.
- [x] **Bound `UseGenericAttributes` inside blockquotes to Markdig behavior.** Blockquote paragraph and heading trailing attribute blocks now stay literal, including nested blockquotes, while pipe-table attributes inside blockquotes still promote to the owning table.
- [x] **Moved `UseGenericAttributes` through list-item trailing blocks.** Root ordered/unordered, nested, and blockquote-contained list items now consume Markdig-style trailing attribute blocks, preserve Markdig's consumed separator whitespace in HTML without projecting attributes onto `<li>`, write normalized trailing attribute blocks, and expose semantic attributes through syntax/native/source-edit proof.
- [x] **Proved `UseGenericAttributes` extension interactions for task lists and footnotes.** Markdig comparison coverage now protects task-list item attribute consumption when `UseTaskLists` is enabled and footnote-definition paragraph attribute projection when `UseFootnotes` is enabled, without promoting the still-partial arbitrary block-family row.
- [x] **Moved `UseGenericAttributes` through paragraph separator and inline HTML/image edges.** Paragraph attribute blocks now preserve Markdig's consumed separator whitespace in HTML and Markdown writing, thematic-break-like attributed lines stay attributed paragraphs, inline and linked images have focused Markdig comparison coverage, and raw inline HTML consumes following attribute blocks without projecting them into rendered HTML.

## Execution Rules

- [ ] Pick exactly one primary row before each slice starts: GFM breadth, one Markdig extension family, AST/source/lossless, renderer/writer, security, or performance.
- [ ] If behavior is missing, improve the engine first: parser, AST, source mapping, renderer, writer, or extension APIs.
- [ ] If behavior exists but is unproven, add focused proof: Markdig comparison cases, inventory rows, native snapshots, writer checks, or renderer checks.
- [ ] Promote a row to `Covered` only with parser behavior, semantic AST/source/native projection where applicable, HTML rendering, Markdown writing or explicit writer limits, fixture/inventory evidence, and profile documentation.
- [ ] Make scope decisions before large new features. Grid tables, custom containers, math, diagrams, attributes, SmartyPants, citations, media links, and similar rows should not be half-added without deciding core versus optional extension versus renderer policy.
- [ ] Benchmark last. Do not optimize or claim performance parity until correctness, source mapping, and writer behavior are stable enough for the numbers to mean something.

## Detailed Backlog

The sections below preserve the evidence trail behind the checklist. Test-only work is allowed only when it creates a missing scoreboard, proves a newly fixed contract, or documents an intentional Markdig difference.

### A. Engine And Parser Behavior

- [ ] **GFM breadth is still thin.** The current GFM inventory is green, but only 52 tracked fixtures are imported. Missing work: broaden strikethrough delimiter edges, HTML tag filtering, extension-interaction fixtures, and any remaining autolink edge cases against upstream-compatible behavior.
- [x] **Markdig `UseAutoLinks` is covered.** CommonMark angle autolinks are green, Markdig-style previous-character/domain-without-period/query-fragment/balanced-parenthesis punctuation, punctuation-before-closing-parenthesis preservation, single trailing punctuation/underscore trimming, trailing semicolon retention, trailing quote retention with paired single-quote literal fallback, lowercase `www.` prefix, `www` and URL authority host-underscore rejection, http/www/ftp user-info authority rejection, optional closing-bracket URL consumption, apostrophe-started bare scheme literal fallback, lowercase bare scheme prefix options, and profile-selectable bare scheme prefixes are implemented. Bare `ftp://`, `tel:`, and `mailto:` scheme autolinks, including `mailto:` path/query targets and Markdig-style `mailto:` semicolon plus address-only colon/dash behavior, have Markdig/source/writer evidence; GFM/CommonMark inventories cover the standards profiles; Markdig-compatible comparison can percent-encode href `~` and render literal non-ASCII display text through explicit HTML options. Broader GFM autolink corpus expansion remains under the GFM breadth row, not as a `UseAutoLinks` blocker.
- [ ] **Raw HTML and GFM tag filtering are still partial.** CommonMark raw HTML is green and cmark-gfm HTML output now has a first-class `HtmlOptions` profile. The GFM inventory now includes focused tag-filter coverage for dangerous inline tags and case-insensitive raw block body filtering, but broader GFM tag-filter corpus coverage, sanitizer/escape/strip/allow mode evidence, source/writer behavior, and URL policy still need to stay separated so parser parity is not confused with security policy.
- [ ] **Definition-list syntax breadth is partial.** OfficeIMO now parses the pinned Markdig colon-marker form, including multiple terms, multiple definitions, grouped AST/source/native/html proof, parsed and generated/rebuilt definition marker tokens, native source-backed marker edits, Markdig lazy paragraph and nested block continuation, loose-definition HTML, edge-continuation comparison, setext-continuation source mapping, empty-marker first-continuation source mapping, grouped Markdown writing that preserves the marker form for reparsing, loose-definition writer preservation, blank-separated marker-group writer preservation, tight nested-list writer preservation, setext-continuation writer reparse proof, and typed plus source-field multiline definition-body edits that preserve valid continuation indentation. Remaining source-map and writer edge breadth still need focused comparison before `UseDefinitionLists` can move to `Covered`.
- [x] **Emphasis extras are covered.** Strikethrough, inserted text, mark/highlight, superscript, and subscript have first-class parser/source/native/render/write coverage, with GFM single-tilde strikethrough kept explicit through profile settings.

### B. Markdig Extension Scope Decisions

- [x] **Extension-family route matrix exists.** `Docs/officeimo.markdown.markdig-extension-inventory.md` now gives every reflected Markdig row a `Scope decision`, `Route`, and `Promotion bar`, so future slices start from the owning layer and done criteria instead of re-deciding scope from scratch.
- [x] **Gap rows are classified before implementation.** Every current `Gap` row now lands in one of the execution buckets below: core engine, optional extension, renderer/host policy, deferred, or intentional difference.
- [ ] **Markdig extension-family coverage is far from closed.** The current inventory is 10 `Covered`, 9 `Partial`, 3 `Intentional`, and 11 `Gap`. Scope is classified, but every non-covered in-scope row still needs implementation, source behavior, renderer/writer behavior, docs, and proof before promotion.
- [ ] **High-priority partial rows need closure.** Work through `UseDefinitionLists`, `UseAlertBlocks`, `UseGenericAttributes`, `UsePreciseSourceLocation`, and parser/render extensions with parser, AST/source, renderer, writer, and fixture evidence.
- [x] **High-priority gap rows need scope decisions before implementation.** `UseCustomContainers`, `UseGridTables`, `UseSmartyPants`, `UseCitations`, `UseMathematics`, `UseMediaLinks`, `UseDiagrams`, `UseFigures`, `UseListExtras`, and similar rows now have generated scope-decision coverage; implementation and promotion remain open where status is `Partial` or `Gap`.
- [x] **Abbreviation parity is covered.** `UseAbbreviations` has opt-in parser, semantic AST, syntax/native metadata, HTML rendering, source-edit, selected Markdig comparison evidence across nested inline/container/table-cell contexts, Markdig-compatible trailing-dash, unresolved-bracket-text, list-item-definition, opening-punctuation behavior, Unicode visible-text rendering, empty-title definition handling, AST/native propagation proof, parse-owned definition-preserving Markdown writing, and list-contained source-token/native edit/writer proof.

#### Gap Row Scope Plan

| Markdig row | Scope decision | Missing before parity |
| --- | --- | --- |
| `UseCustomContainers` | Core engine | Block parser extension contract, nested child source mapping, renderer/writer source slices, and Markdig fixtures. |
| `UseGridTables` | Optional extension | Grid-table AST/source model, malformed-table fallback, HTML rendering, Markdown writing, and Markdig/Pandoc-style fixtures. |
| `UseListExtras` | Optional extension | Exact Markdig list-extra syntax inventory, canonical list-item mapping, source spans, writer behavior, and fixtures. |
| `UseEmojiAndSmiley` | Optional extension | Shortcode/smiley tables, opt-in transform contract, source metadata, writer rules, and normalization boundaries. |
| `UseJiraLinks` | Optional extension | Configurable issue-key resolver, link source metadata, renderer policy, and writer preservation. |
| `UseSmartyPants` | Optional extension | Smart punctuation transform, escaping rules, source/edit behavior, writer policy, and delimiter interaction proof. |
| `UseReferralLinks` | Renderer/host policy | Opt-in link rendering policy, safe defaults, rel/referrer output fixtures, and writer-neutral behavior. |
| `UseCitations` | Deferred | Real consumer requirement, citation AST contract, renderer/writer contract, and fixtures after core/GFM closure. |
| `UseFooters` | Deferred | Document footer semantics requirement, footer block parser, semantic node, renderer/writer behavior, and fixtures. |
| `UseGlobalization` | Deferred | Concrete culture-sensitive Markdown contract and fixtures. |
| `UsePragmaLines` | Deferred | Concrete metadata workflow, source-preserving pragma parser, semantic contract, writer behavior, and fixtures. |

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
- [ ] **5. Make scope decisions before large new features.** Grid tables, custom containers, math, diagrams, attributes, SmartyPants, citations, media links, and similar rows should not be half-added without deciding core versus optional extension versus renderer policy. Abbreviations already have an in-core partial implementation and should now be completed as an edge-breadth slice.
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

- [ ] **GFM fixture breadth:** expand beyond the current 52 tracked fixtures for strikethrough, tag filtering, extension interactions, and any remaining autolink edge cases.
- [x] **Pipe tables:** moved from partial support to covered support by proving malformed delimiters, alignment, containers, source spans, renderer output, and writer behavior.
- [x] **Task lists:** moved from partial support to covered support by proving nested markers, exact marker source spans, native snapshots/source edits, renderer output, and ordered/unordered writer behavior.
- [x] **Footnotes:** moved from partial support to covered support by proving Markdig/GFM breadth, label/body source mapping, renderer output, backlink behavior, and writer behavior.
- [x] **Soft line break as hard line break:** moved from missing support to covered support by adding an explicit reader option with HTML output, Markdown writing, nested paragraph propagation, and native metadata proof that synthetic hard breaks do not claim a fake marker.
- [x] **Auto identifiers:** moved from missing support to covered support by proving automatic heading ids, disable behavior, Markdig default and GitHub slug styles, duplicate tracking, GFM profile wiring, and existing heading source/native metadata.
- [x] **YAML front matter:** moved from partial support to covered support by preserving raw YAML as the AST payload, keeping structured helpers for simple entries, exposing body/fence/key/value source spans through syntax/native snapshots, omitting front matter from HTML, and preserving the raw body through Markdown writing.
- [ ] **Autolinks and tag filter:** separate CommonMark autolinks, GFM extended autolinks, and GFM tag-filter behavior into explicit parser/render/security contracts.
- [x] **Extension-family decisions:** every row in `Docs/officeimo.markdown.markdig-extension-inventory.md` now has a first-pass `Route` and `Promotion bar`; implementation slices should refine those fields when real parser/renderer evidence changes the decision.

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

These are the actual parity gaps. Engine work is listed first; tests and inventories are proof lanes, not the definition of parity.

- [x] **CommonMark parser correctness.** The official CommonMark `0.31.2` inventory is green, including HTML block boundaries, entities, hard breaks, emphasis, code spans, and container indentation.
- [x] **Core scoreboards.** Checked-in inventories exist for CommonMark, tracked GFM fixtures, and reflected Markdig extension families.
- [x] **Initial editor/source architecture.** Many canonical nodes expose syntax/source/native spans, source fields, and source-edit helpers.
- [x] **Finish the active `UseGenericAttributes` engine slice.** Covered generic attribute blocks are semantic/native/source-backed and now project as `GenericAttributeBlock` syntax tokens with caret/navigation and native metadata proof.
- [ ] **Finish `UseGenericAttributes` breadth.** Extend parsing/writing/source preservation beyond the currently covered ATX headings, Setext headings, paragraphs, fenced code, list-item trailing attributes, Markdig-style pipe-table cell attributes, blockquote literal-boundary behavior, and common no-space inline elements into the remaining block families and inline shapes that Markdig supports.
- [ ] **Finish `UseDefinitionLists` breadth.** Marker syntax and native fields exist; finish source-map and writer edge cases for marker groups, lazy continuation, loose definitions, nested definitions, multiline body edits, and reparse stability.
- [ ] **Decide `UseAlertBlocks`.** Choose whether Markdig alert rendering callbacks are an OfficeIMO engine/renderer contract or an intentional OfficeIMO callout difference, then implement or document that decision.
- [ ] **Decide `UseCjkFriendlyEmphasis`.** Either add a real delimiter option with CJK comparison fixtures and source-token proof, or document it as deferred/intentional.
- [ ] **Close GFM breadth gaps.** Expand autolink, tag-filter, and strikethrough delimiter-edge coverage only where it protects current GFM profile behavior or exposes a real parser/renderer gap.
- [ ] **Canonicalize remaining AST shapes.** Clean up duplicated or uneven node ownership for `ListItem`, `TableBlock`, `CalloutBlock`, `DefinitionListBlock`, and any other mutable node shapes that block stable source/native APIs.
- [ ] **Complete source/lossless infrastructure.** Add parser-owned trivia capture, delimiter-token coverage, generated-node diagnostics, original-to-normalized offset mapping, and broader byte-preserving `MarkdownRoundtripWriter` edits.
- [ ] **Finish renderer/writer extension parity.** Make custom parser, transform, renderer, and writer APIs source-slice aware so extension nodes do not require downstream string rescanning.
- [ ] **Make security/profile behavior explicit.** Independently prove raw HTML allow/strip/escape/sanitize, URL policy, and GFM tag-filter behavior so security policy is not mixed up with parser parity.
- [ ] **Add performance proof last.** Capture release-mode comparisons against Markdig for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative README/docs/chat corpora after correctness stabilizes.
- [ ] **Keep docs and package claims aligned.** README-level claims, compatibility docs, generated inventories, and package behavior must describe the same current truth before calling this Markdig-class parity.

## Immediate Work Order

- [ ] **Active slice: continue `UseGenericAttributes` breadth.**
  - [ ] Probe Markdig output for remaining block and inline attribute shapes before changing code; list-item trailing attributes are already closed in the current branch.
  - [ ] Implement supported missing shapes in the shared parser/source/writer path.
  - [ ] Add Markdig comparison, syntax/native/source-edit, HTML, Markdown writer, and reparse proof for each newly supported shape.
  - [ ] Refresh the Markdig extension inventory and promote `UseGenericAttributes` only if arbitrary supported block/inline breadth is truly closed.
- [ ] **Next slice: finish `UseDefinitionLists` breadth.**
  - [ ] Close marker-group, lazy-continuation, loose-definition, nested-body, source-map, writer, and reparse edge cases.
  - [ ] Promote only after source/native and writer evidence covers the remaining Markdig-compatible cases.
- [ ] **Decision slice: alert and CJK behavior.**
  - [ ] Decide whether Markdig alert rendering callbacks are in scope or an intentional OfficeIMO callout difference.
  - [ ] Decide whether `UseCjkFriendlyEmphasis` gets a real delimiter option or is documented as deferred/intentional.
- [ ] **Source/lossless slice: editor-grade parity.**
  - [ ] Canonicalize duplicated node shapes for `ListItem`, `TableBlock`, `CalloutBlock`, and `DefinitionListBlock`.
  - [ ] Implement parser-owned trivia capture, delimiter-token coverage, generated-node diagnostics, and original-to-normalized mapping.
  - [ ] Expand `MarkdownRoundtripWriter` beyond unchanged documents and explicit native edits toward broader source-preserving edit writing.
- [ ] **Proof-only slice: broaden scoreboards after engine contracts stabilize.**
  - [ ] Expand GFM autolink, tag-filter, and strikethrough delimiter-edge fixture breadth.
  - [ ] Refresh Markdig extension rows after each engine slice.
  - [ ] Run release-mode benchmarks only after correctness/source/writer behavior stops moving.

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
- [ ] Expand GFM autolinks and remaining tag-filter coverage.
- [x] Expand GFM footnote coverage, including source spans, nested block bodies, repeated backrefs, renderer output, and writer behavior.
- [ ] Expand GFM strikethrough coverage.
- [ ] Decide which remaining Markdig extensions are in scope: grid tables, emoji, math, diagrams, SmartyPants, citations, custom containers, generic attributes, media links, alerts, advanced links, and list/emphasis extras. Abbreviations are now an in-core partial row that still needs edge-breadth closure.
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
