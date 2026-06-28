# OfficeIMO.Markdown Markdig Parity Gap Plan

This is the working plan for getting `OfficeIMO.Markdown` to Markdig-class parity without looping through disconnected fixture additions.

The important distinction: parity is not "more tests." Parity means the parser, AST, renderer, writer, extension model, source mapping, lossless behavior, and performance story all line up behind a documented contract. Tests and inventories are the proof gates for those contracts.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 296 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 632 of 652 official CommonMark `0.31.2` examples currently match; 20 are failing in `Docs/officeimo.markdown.commonmark-inventory.md` |
| GFM corpus | 36 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| GFM tracked inventory | 36 tracked GFM fixtures in `Docs/officeimo.markdown.gfm-inventory.md`: 33 upstream cmark-gfm fixtures, 3 OfficeIMO supplements, 36 passing, 0 failing |
| Markdig extension inventory | 33 Markdig extension-family rows in `Docs/officeimo.markdown.markdig-extension-inventory.md`: 0 covered, 15 partial, 4 intentional, 14 gap |
| Covered CommonMark sections | ATX headings, Setext headings, thematic breaks, fenced code blocks, lists, paragraphs, soft breaks, links, images, autolinks, backslash escapes, link reference definitions |
| Remaining CommonMark parser clusters | emphasis delimiter runs, container indentation/continuation, hard-line-break edge cases, CommonMark entity decoding, one HTML block boundary case |
| Remaining Markdig-class architecture gaps | broader GFM corpus coverage, full lossless trivia capture, full parser pipeline parity, renderer/writer plugin parity, extension-family implementation breadth, release-mode benchmark review |

## Current Answer

We are not doing only tests. The inventories are the measuring system. They tell us which engine contract to improve next and stop us from calling isolated fixture additions "parity."

The current state is:

- [x] We can measure CommonMark, tracked GFM fixtures, and reflected Markdig extension families from checked-in reports.
- [x] GFM smoke behavior is green for the fixture corpus we track today.
- [ ] CommonMark is not closed: 20 official examples still fail, led by emphasis, containers, hard breaks, entities, and one HTML block boundary case.
- [ ] Markdig extension parity is not closed: 15 extension families are partial, 14 are gaps, 4 are intentional differences, and 0 meet the full `Covered` bar.
- [ ] AST/source/lossless parity is not closed: parser nodes, native snapshots, trivia, source edits, renderer contexts, and writer contexts still need to line up everywhere.
- [ ] Performance parity is not known: release-mode Markdig comparisons still need a stable benchmark pass after correctness stops moving.

So the real work now is engine-first:

1. [ ] Fix a parser/AST/renderer/writer behavior slice.
2. [ ] Refresh the inventory that proves the moved count.
3. [ ] Promote only the official examples that now represent understood contracts.
4. [ ] Update this plan and the relevant compatibility report.
5. [ ] Validate across the Markdown lanes before committing.

## Execution Checklist

Use this order unless a later discovery proves a dependency should move earlier.

- [ ] **Slice 1: Finish the last CommonMark HTML block boundary case.** The raw HTML section is green; HTML block/raw HTML is down to example #174, which is container interaction rather than inline raw HTML tokenization.
- [ ] **Slice 2: CommonMark emphasis delimiter algorithm.** Replace local emphasis heuristics with delimiter-stack behavior that handles flanking, punctuation, intraword underscores, nesting, opener/closer balancing, and precedence.
- [ ] **Slice 3: CommonMark container indentation.** Make tabs, blockquote continuation, list continuation, and indented-code boundaries share a source-map-safe column model.
- [ ] **Slice 4: CommonMark hard breaks and entities.** Finish hard-line-break marker handling and replace narrow HTML decoding with a CommonMark-complete named/numeric character reference decoder.
- [ ] **Slice 5: GFM breadth.** Keep the current green GFM smoke corpus, then broaden table/task/autolink/strikethrough/tag-filter/footnote/interoperability fixtures from upstream-compatible sources.
- [ ] **Slice 6: Markdig extension family decisions.** For each `Partial` or `Gap` row, choose one of: implement in core, implement as an optional extension, route to renderer/host policy, or document as intentional out of scope.
- [ ] **Slice 7: AST/source/lossless closure.** Finish canonical node cleanup, trivia capture, delimiter-token capture, source edits, generated-node diagnostics, and native snapshot consistency.
- [ ] **Slice 8: Renderer/writer extension parity.** Make custom parser nodes render and write from typed/source-aware contexts rather than downstream string rescanning.
- [ ] **Slice 9: Performance parity proof.** Run release-mode Markdig comparisons over parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative corpora.

References:

- CommonMark `0.31.2` official JSON: https://spec.commonmark.org/0.31.2/spec.json
- Markdig README and feature baseline: https://github.com/xoofx/markdig
- Markdig roundtrip/trivia target: https://github.com/xoofx/markdig/blob/master/src/Markdig/Roundtrip.md
- Markdig extension specs index: https://github.com/xoofx/markdig/blob/master/src/Markdig.Tests/Specs/readme.md

## What We Already Have

- [x] A repo-owned CommonMark inventory that runs all 652 official examples without forcing every known failure to fail CI.
- [x] A checked-in failure-cluster report so we can pick work by root cause instead of nearby example numbers.
- [x] A pinned CommonMark smoke lane with 296 official examples.
- [x] CommonMark code spans are green in the generated inventory after preserving NBSP as non-collapsible text in the HTML comparison harness and pinning official examples 333 and 334.
- [x] CommonMark links are green in the generated inventory after fixing link-label inline-span precedence.
- [x] Strong current coverage for headings, thematic breaks, fenced code, lists, paragraphs, soft breaks, backslash escapes, autolinks, images, and link reference definitions.
- [x] A generated GFM inventory tracks the current extension fixture corpus by section and source, separating upstream cmark-gfm fixtures from OfficeIMO supplements.
- [x] A generated Markdig extension inventory tracks reflected Markdig extension-family entry points and classifies current OfficeIMO support as covered, partial, intentional, or gap.
- [x] Syntax/source/native tests for many existing nodes, including source slices and source-edit helpers.
- [x] Markdig package baseline guarded across tests, benchmarks, and docs.

## What We Are Missing

These are the actual parity gaps. The test work is listed only where it creates a missing scoreboard.

- [ ] **CommonMark HTML block boundary parity.** Finish the remaining HTML block failure around blockquote/container boundaries; raw inline HTML and raw HTML section failures are currently green in the full inventory.
- [ ] **CommonMark emphasis engine parity.** Replace the simplified emphasis parser with the CommonMark delimiter-run algorithm covering left/right flanking, punctuation, intraword `_`, nesting, opener/closer balancing, and precedence.
- [ ] **CommonMark container indentation parity.** Finish the 6 remaining failures around tabs, blockquote/list continuation, indented-code boundaries, and source-map-safe column handling.
- [ ] **CommonMark hard-break and inline precedence parity.** Fix the 3 remaining hard-line-break failures while preserving marker source spans for two-space, backslash, and inline HTML break spellings.
- [ ] **CommonMark entity decoder parity.** Replace narrow runtime HTML decoding with a CommonMark-complete named/numeric character reference decoder, including invalid numeric replacement behavior.
- [ ] **GFM corpus expansion.** Extend the generated GFM inventory beyond the current tracked fixture corpus so table, task-list, autolink, strikethrough, tag-filter, footnote, and interop behavior are measured against broader upstream-compatible coverage.
- [ ] **Markdig extension implementation breadth.** Move selected partial/gap rows from the generated Markdig extension inventory into real parser, AST/source, renderer, writer, fixture, or intentional-deviation work.
- [ ] **AST/source/lossless completeness.** Finish canonical node cleanup, full trivia capture, delimiter-token capture, original-to-normalized mapping, generated-node diagnostics, and broader byte-preserving source edits.
- [ ] **Renderer/writer extension parity.** Make parser, transform, renderer, and writer extension APIs source-slice aware where custom nodes need to render or roundtrip without string rescanning.
- [ ] **Renderer/security profile parity.** Keep CommonMark/GFM HTML output spec-compatible while making OfficeIMO-specific raw HTML, sanitizing, escaping, URL policy, and GFM tag-filter behavior explicit and independently tested.
- [ ] **Performance proof.** Capture release-mode benchmarks against Markdig for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative README/docs/chat corpora.

## Immediate Work Order

- [ ] Finish the remaining HTML block/container boundary case (#174) if it shares work with container continuation.
- [ ] Tackle emphasis next because it is now the largest remaining CommonMark failure cluster and needs a deliberate delimiter-stack rewrite rather than local patches.
- [ ] Use `Docs/officeimo.markdown.markdig-extension-inventory.md` to pick extension-family slices by row; do not promote `Partial` rows to `Covered` without parser, AST/source, renderer, writer, and fixture evidence.
- [ ] Tackle container indentation and hard breaks as separate slices unless the source-map model forces a shared tab/column primitive.
- [ ] Tackle the entity decoder as a reusable parser service and route every required decode context through it.
- [ ] Return to AST/source/lossless completion once the major parser clusters stop moving node boundaries every slice.
- [ ] Run release-mode benchmarks only after correctness stabilizes enough for the numbers to mean something.

## CommonMark Failure Clusters

| Cluster | Failing | Sections | First examples | Work type |
| --- | ---: | --- | --- | --- |
| Emphasis delimiter algorithm | 8 | Emphasis and strong emphasis | #408, #418, #432, #438, #441, #450, #453, #470 | Inline parser rewrite |
| Container indentation and continuation | 6 | Block quotes, Indented code blocks, List items, Tabs | #9, #111, #231, #242, #252, #264 | Block parser and source-map column model |
| Inline precedence and line-break grammar | 3 | Hard line breaks | #642, #643, #644 | Inline parser and marker source spans |
| CommonMark entity decoder | 2 | Entity and numeric character references | #25, #26 | Shared character-reference decoder |
| HTML block/raw HTML grammar | 1 | HTML blocks | #174 | Blockquote/container boundary |

## Pinned CommonMark Coverage By Section

This is fixture coverage, not a claim that unpinned examples fail. The generated inventory is the pass/fail source of truth.

| Section | Pinned | Total | Missing | Primary missing work |
| --- | ---: | ---: | ---: | --- |
| Emphasis and strong emphasis | 11 | 132 | 121 | Inline delimiter algorithm |
| Links | 27 | 90 | 63 | CommonMark inventory is green; remaining work is breadth and Markdig/GFM comparison |
| HTML blocks | 19 | 44 | 25 | One remaining blockquote/container boundary failure |
| Raw HTML | 8 | 20 | 12 | Current CommonMark inventory is green; keep broader source-span/writer coverage aligned |
| Images | 2 | 22 | 20 | Breadth and source metadata, not current CommonMark failures |
| Code spans | 8 | 22 | 14 | Current CommonMark inventory is green; keep delimiter/source-token coverage aligned |
| Link reference definitions | 10 | 27 | 17 | Breadth and source metadata, not current CommonMark failures |
| Block quotes | 10 | 25 | 15 | Container continuation rules |
| Autolinks | 12 | 19 | 7 | Official CommonMark is green; keep GFM/profile extensions separate |
| Indented code blocks | 1 | 12 | 11 | Indent/tab/list interaction |
| Tabs | 0 | 11 | 11 | Source map and indentation model |
| Entity and numeric character references | 7 | 17 | 10 | CommonMark entity decoder |
| Hard line breaks | 6 | 15 | 9 | Inline break parser |
| List items | 38 | 48 | 10 | Remaining edge cases only |
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

- [ ] HTML/raw HTML: implement all seven CommonMark HTML block types and profile-aware raw HTML behavior.
- [ ] Emphasis: implement the CommonMark delimiter-run algorithm.
- [ ] Containers: implement tab expansion and continuation indentation as parser/source-map primitives.
- [ ] Hard breaks: finish CommonMark line-break grammar without losing marker source spans.
- [x] Code spans: clear the NBSP/equivalence bucket and pin the official cases.
- [ ] Entities: implement a CommonMark-complete character reference decoder.

Done means:

- [ ] The official CommonMark `0.31.2` inventory has no unexplained failures.
- [ ] Any intentional OfficeIMO profile differences are documented in the compatibility matrix instead of hidden as failing examples.

### Phase 2: GFM And Markdig Extension Breadth

- [ ] Expand GFM tables beyond smoke fixtures, including malformed delimiters and container interactions.
- [ ] Expand GFM task-list coverage, including marker source spans and nested list behavior.
- [ ] Expand GFM autolinks and tag-filter coverage.
- [ ] Expand GFM footnotes and strikethrough coverage.
- [ ] Decide which Markdig extensions are in scope: grid tables, pipe tables, auto identifiers, YAML front matter, emoji, math, diagrams, SmartyPants, abbreviations, citations, custom containers, generic attributes, media links, task lists, alerts, and advanced links.
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
