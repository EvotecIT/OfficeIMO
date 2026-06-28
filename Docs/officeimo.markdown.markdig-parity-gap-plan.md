# OfficeIMO.Markdown Markdig Parity Gap Plan

This is the working plan for getting `OfficeIMO.Markdown` to Markdig-class parity without looping through disconnected fixture additions.

The important distinction: parity is not "more tests." Parity means the parser, AST, renderer, writer, extension model, source mapping, lossless behavior, and performance story all line up behind a documented contract. Tests and inventories are the proof gates for those contracts.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 294 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 619 of 652 official CommonMark `0.31.2` examples currently match; 33 are failing in `Docs/officeimo.markdown.commonmark-inventory.md` |
| GFM corpus | 36 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| Covered CommonMark sections | ATX headings, Setext headings, thematic breaks, fenced code blocks, lists, paragraphs, soft breaks, links, images, autolinks, backslash escapes, link reference definitions |
| Remaining CommonMark parser clusters | HTML/raw HTML grammar, emphasis delimiter runs, container indentation/continuation, hard-line-break edge cases, code-span NBSP/equivalence, CommonMark entity decoding |
| Remaining Markdig-class architecture gaps | GFM/Markdig inventories, full lossless trivia capture, full parser pipeline parity, renderer/writer plugin parity, broader extension set, release-mode benchmark review |

References:

- CommonMark `0.31.2` official JSON: https://spec.commonmark.org/0.31.2/spec.json
- Markdig README and feature baseline: https://github.com/xoofx/markdig
- Markdig roundtrip/trivia target: https://github.com/xoofx/markdig/blob/master/src/Markdig/Roundtrip.md
- Markdig extension specs index: https://github.com/xoofx/markdig/blob/master/src/Markdig.Tests/Specs/readme.md

## What We Already Have

- [x] A repo-owned CommonMark inventory that runs all 652 official examples without forcing every known failure to fail CI.
- [x] A checked-in failure-cluster report so we can pick work by root cause instead of nearby example numbers.
- [x] A pinned CommonMark smoke lane with 294 official examples.
- [x] CommonMark links are green in the generated inventory after fixing link-label inline-span precedence.
- [x] Strong current coverage for headings, thematic breaks, fenced code, lists, paragraphs, soft breaks, backslash escapes, autolinks, images, and link reference definitions.
- [x] Syntax/source/native tests for many existing nodes, including source slices and source-edit helpers.
- [x] Markdig package baseline guarded across tests, benchmarks, and docs.

## What We Are Missing

These are the actual parity gaps. The test work is listed only where it creates a missing scoreboard.

- [ ] **CommonMark HTML/raw HTML engine parity.** Fix the 11 remaining failures around HTML block types, raw-text containers, multiline tags, block boundaries, and inline-vs-block raw HTML classification.
- [ ] **CommonMark emphasis engine parity.** Replace the simplified emphasis parser with the CommonMark delimiter-run algorithm covering left/right flanking, punctuation, intraword `_`, nesting, opener/closer balancing, and precedence.
- [ ] **CommonMark container indentation parity.** Finish the 6 remaining failures around tabs, blockquote/list continuation, indented-code boundaries, and source-map-safe column handling.
- [ ] **CommonMark hard-break and inline precedence parity.** Fix the 3 remaining hard-line-break failures while preserving marker source spans for two-space, backslash, and inline HTML break spellings.
- [ ] **CommonMark entity decoder parity.** Replace narrow runtime HTML decoding with a CommonMark-complete named/numeric character reference decoder, including invalid numeric replacement behavior.
- [ ] **CommonMark code-span parity.** Resolve the 2 remaining NBSP/equivalence failures and keep code-span delimiter/content source tokens stable. If the engine is already correct, fix the comparison harness and pin the official cases as validation proof.
- [ ] **GFM inventory.** Add a generated cmark-gfm inventory so table, task-list, autolink, strikethrough, tag-filter, and footnote behavior is measured instead of represented only by curated smoke fixtures.
- [ ] **Markdig extension inventory.** Add a comparison inventory for Markdig extension families and classify each as `Covered`, `Partial`, `Intentional`, or `Gap`.
- [ ] **AST/source/lossless completeness.** Finish canonical node cleanup, full trivia capture, delimiter-token capture, original-to-normalized mapping, generated-node diagnostics, and broader byte-preserving source edits.
- [ ] **Renderer/writer extension parity.** Make parser, transform, renderer, and writer extension APIs source-slice aware where custom nodes need to render or roundtrip without string rescanning.
- [ ] **Renderer/security profile parity.** Keep CommonMark/GFM HTML output spec-compatible while making OfficeIMO-specific raw HTML, sanitizing, escaping, URL policy, and GFM tag-filter behavior explicit and independently tested.
- [ ] **Performance proof.** Capture release-mode benchmarks against Markdig for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, transforms, allocations, and representative README/docs/chat corpora.

## Immediate Work Order

- [ ] Finish the current small code-span/NBSP slice if it is only a comparison-harness bug, because it clears a whole CommonMark failure bucket with low engine risk.
- [ ] Build the GFM and Markdig inventories next, before broad extension work, so extension parity has the same scoreboard CommonMark now has.
- [ ] Tackle HTML/raw HTML after the inventory work; it is the largest remaining CommonMark failure cluster and affects renderer/security/profile behavior.
- [ ] Tackle emphasis after HTML because it needs a deliberate delimiter-stack rewrite rather than local patches.
- [ ] Tackle container indentation and hard breaks as separate slices unless the source-map model forces a shared tab/column primitive.
- [ ] Tackle the entity decoder as a reusable parser service and route every required decode context through it.
- [ ] Return to AST/source/lossless completion once the major parser clusters stop moving node boundaries every slice.
- [ ] Run release-mode benchmarks only after correctness stabilizes enough for the numbers to mean something.

## CommonMark Failure Clusters

| Cluster | Failing | Sections | First examples | Work type |
| --- | ---: | --- | --- | --- |
| HTML block/raw HTML grammar | 11 | HTML blocks, Raw HTML | #148, #174, #191, #619, #621, #622, #625, #626, #627, #628, #629 | Parser and renderer/security profile |
| Emphasis delimiter algorithm | 9 | Emphasis and strong emphasis | #353, #408, #418, #432, #438, #441, #450, #453, #470 | Inline parser rewrite |
| Container indentation and continuation | 6 | Block quotes, Indented code blocks, List items, Tabs | #9, #111, #231, #242, #252, #264 | Block parser and source-map column model |
| Inline precedence and line-break grammar | 3 | Hard line breaks | #642, #643, #644 | Inline parser and marker source spans |
| Code span normalization and precedence | 2 | Code spans | #333, #334 | Parser or comparison harness, depending on NBSP proof |
| CommonMark entity decoder | 2 | Entity and numeric character references | #25, #26 | Shared character-reference decoder |

## Pinned CommonMark Coverage By Section

This is fixture coverage, not a claim that unpinned examples fail. The generated inventory is the pass/fail source of truth.

| Section | Pinned | Total | Missing | Primary missing work |
| --- | ---: | ---: | ---: | --- |
| Emphasis and strong emphasis | 11 | 132 | 121 | Inline delimiter algorithm |
| Links | 27 | 90 | 63 | CommonMark inventory is green; remaining work is breadth and Markdig/GFM comparison |
| HTML blocks | 19 | 44 | 25 | HTML block tokenizer |
| Raw HTML | 8 | 20 | 12 | Inline/raw HTML classification |
| Images | 2 | 22 | 20 | Breadth and source metadata, not current CommonMark failures |
| Code spans | 6 | 22 | 16 | NBSP/equivalence bucket plus broader fixture promotion |
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
- [ ] Add the same generated inventory style for enabled cmark-gfm extensions.
- [ ] Add a Markdig comparison inventory that separates OfficeIMO profile differences from portable/CommonMark profile differences.
- [ ] Add an extension-family table with `Covered`, `Partial`, `Intentional`, or `Gap` for every Markdig extension family we care about.

Done means:

- [ ] We can answer "what is missing?" for CommonMark, GFM, and Markdig extension parity from checked-in reports.
- [ ] Every future engine slice names which scoreboard row it moves.

### Phase 1: CommonMark Parser Closure

- [ ] HTML/raw HTML: implement all seven CommonMark HTML block types and profile-aware raw HTML behavior.
- [ ] Emphasis: implement the CommonMark delimiter-run algorithm.
- [ ] Containers: implement tab expansion and continuation indentation as parser/source-map primitives.
- [ ] Hard breaks: finish CommonMark line-break grammar without losing marker source spans.
- [ ] Code spans: clear the NBSP/equivalence bucket and pin the official cases.
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

- [ ] Enabled GFM behavior is measured by inventory.
- [ ] Markdig extension parity is an explicit support matrix, not an implied promise.

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
