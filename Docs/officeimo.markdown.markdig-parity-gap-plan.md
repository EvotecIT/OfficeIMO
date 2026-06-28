# OfficeIMO.Markdown Markdig Parity Gap Plan

This plan is the working checklist for getting `OfficeIMO.Markdown` to Markdig-class parity without looping through random fixture additions.

The short version: parity is not "more tests." Parity means the engine, AST, renderer, extension seams, lossless behavior, and compatibility corpus all agree on a documented contract. Tests are the proof gate for each engine slice.

## Current Scoreboard

| Area | Current state |
| --- | --- |
| Local Markdig comparison package | Markdig `1.3.2`, guarded across tests, benchmarks, and compatibility docs |
| CommonMark corpus | 255 of 652 official CommonMark `0.31.2` examples pinned as smoke fixtures |
| CommonMark full inventory | 575 of 652 official CommonMark `0.31.2` examples currently match; 77 are failing in `Docs/officeimo.markdown.commonmark-inventory.md` |
| GFM corpus | 36 cmark-gfm extension smoke fixtures plus focused crash/regression coverage |
| Strong areas | ATX headings, Setext headings, thematic breaks, fenced code blocks, paragraphs, lists, soft breaks |
| Biggest remaining parser gaps | Link/reference grammar, HTML/raw HTML grammar, emphasis, container indentation, autolinks, code spans, hard-break edge cases, entities |
| Biggest Markdig-class architecture gaps | Full lossless trivia capture, full parser pipeline parity, full renderer/writer plugin parity, broader extension set, release-mode benchmark review |

References:

- CommonMark `0.31.2` official JSON: https://spec.commonmark.org/0.31.2/spec.json
- Markdig README and feature baseline: https://github.com/xoofx/markdig
- Markdig roundtrip/trivia target: https://github.com/xoofx/markdig/blob/master/src/Markdig/Roundtrip.md
- Markdig extension specs index: https://github.com/xoofx/markdig/blob/master/src/Markdig.Tests/Specs/readme.md

## What We Are Missing Right Now

This is the current `[ ]` work plan. Tests are not the goal; they are the proof that each engine slice is actually compatible.

- [ ] Fix the 24 remaining link/reference failures: link reference definitions, nested link precedence, invalid destination/title fallback, NBSP separators, and shared image/link grammar.
- [ ] Fix the 22 remaining HTML/raw HTML failures: full seven-type HTML block scanning, multiline/malformed opening tags, raw-text containers, and inline-vs-block raw HTML boundaries.
- [ ] Fix the 9 remaining emphasis failures by replacing simplified delimiter handling with the CommonMark delimiter-run algorithm.
- [ ] Fix the 8 remaining container/indentation failures around tabs, blockquote/list continuation, and indented-code boundaries.
- [ ] Fix the 4 remaining CommonMark autolink failures without regressing the GFM and OfficeIMO bare-autolink profiles.
- [ ] Fix the 4 remaining code-span failures around normalization, Unicode spaces, and precedence.
- [ ] Fix the 4 remaining hard-line-break failures and keep marker source spans stable.
- [ ] Fix the 2 remaining entity decoder failures with a CommonMark-complete named/numeric character reference decoder.
- [ ] Add a cmark-gfm inventory like the CommonMark inventory so GFM work is not limited to curated smoke fixtures.
- [ ] Add a Markdig comparison inventory that separates true gaps from intentional OfficeIMO profile differences.
- [ ] Finish AST/source/lossless work: remaining canonical node cleanup, complete trivia capture, complete delimiter-token capture, and byte-preserving source edits beyond the current safe cases.
- [ ] Review release-mode benchmarks against the Markdig baseline before making any performance or parity claims.

## Pinned CommonMark Coverage By Section

This table is fixture coverage, not a claim that every unpinned example currently fails. A required early step is to generate the full pass/fail inventory for every unpinned official example.

| Section | Pinned | Total | Missing | Primary lane |
| --- | ---: | ---: | ---: | --- |
| Emphasis and strong emphasis | 11 | 132 | 121 | Inline delimiter algorithm |
| Links | 9 | 90 | 81 | Link/image/reference grammar |
| HTML blocks | 11 | 44 | 33 | HTML block tokenizer |
| Raw HTML | 6 | 20 | 14 | Inline/raw HTML classification |
| Images | 2 | 22 | 20 | Link/image/reference grammar |
| Code spans | 5 | 22 | 17 | Code span normalization and precedence |
| Link reference definitions | 5 | 27 | 22 | Reference-definition grammar |
| Block quotes | 10 | 25 | 15 | Container continuation rules |
| Autolinks | 8 | 19 | 11 | Autolink grammar |
| Indented code blocks | 1 | 12 | 11 | Indent/tab/list interaction |
| Tabs | 0 | 11 | 11 | Source map and indentation model |
| Entity and numeric character references | 7 | 17 | 10 | CommonMark entity decoder |
| Hard line breaks | 5 | 15 | 10 | Inline break parser |
| List items | 38 | 48 | 10 | Remaining list edge cases |
| Backslash escapes | 8 | 13 | 5 | Inline escape parser |
| Textual content | 0 | 3 | 3 | Baseline text handling |
| Blank lines | 0 | 1 | 1 | Baseline block handling |
| Inlines | 0 | 1 | 1 | Inline precedence |
| Precedence | 0 | 1 | 1 | Parser precedence |
| ATX headings | 18 | 18 | 0 | Covered, keep invariant coverage |
| Fenced code blocks | 29 | 29 | 0 | Covered, keep invariant coverage |
| Lists | 26 | 26 | 0 | Covered, keep invariant coverage |
| Paragraphs | 8 | 8 | 0 | Covered, keep invariant coverage |
| Setext headings | 27 | 27 | 0 | Covered, keep invariant coverage |
| Soft line breaks | 2 | 2 | 0 | Covered, keep invariant coverage |
| Thematic breaks | 19 | 19 | 0 | Covered, keep invariant coverage |

## Rules To Stop The Loop

- [x] Do not add an official fixture just because it currently passes; first classify which engine contract it proves.
- [x] Do not add a fixture for a failing example until the engine behavior is fixed or the deviation is documented as intentional.
- [x] Every parity slice must say which bucket it reduces: parser grammar, AST/source mapping, renderer/writer behavior, extension seam, lossless roundtrip, GFM extension parity, or benchmark evidence.
- [x] Each completed slice must update the compatibility matrix count, the gap plan if it changes priority, and the focused test lane.
- [x] If a slice touches parsing, validate at least net8 plus the broad Markdown lane before commit.
- [x] If a slice touches public AST/source/native APIs, add source-span/native/syntax invariant proof, not only HTML output.

## Phase 0: Build The Gap Inventory

- [x] Add a repo-owned CommonMark inventory tool or test mode that runs all 652 official examples and writes a categorized report without making every failure fail CI at once.
- [x] Split the report into `passing-unpinned`, `failing`, and `intentional-deviation` groups.
- [x] Group failures by root parser cause, not by example number.
- [x] Add a small checked-in summary artifact with counts per section and top failure clusters.
- [ ] Add the same inventory style for enabled cmark-gfm extensions.
- [ ] Add a Markdig comparison inventory that separates OfficeIMO profile differences from portable/CommonMark profile differences.

Done means:

- [x] We can answer "what is missing?" with current generated counts instead of memory or manual probes.
- [x] The next work item can be picked from the largest failure cluster without re-discovering the same facts.

## Phase 1: CommonMark Entity Decoder

Current observed failures include broad named entities such as `&Dcaron;`, `&HilbertSpace;`, `&DifferentialD;`, `&ClockwiseContourIntegral;`, `&ngE;`, and numeric replacement handling such as `&#0;`.

- [ ] Replace the narrow runtime HTML decoder dependency with a CommonMark-complete character reference decoder.
- [ ] Decode named and numeric character references consistently in text, attribute-producing inline metadata, fenced-code info strings, and every CommonMark-required non-code context.
- [ ] Preserve source text metadata for decoded entity runs.
- [ ] Replace invalid numeric references with U+FFFD where the spec requires it.
- [ ] Keep code spans and indented code blocks literal where entity decoding must not happen.

Done means:

- [ ] Official entity examples 25-33 are either passing and pinned or intentionally documented.
- [ ] Entity source-span and native metadata tests still prove the original spelling is addressable.

## Phase 2: HTML Block And Raw HTML Grammar

Current observed failures include multiline opening tags, raw table/pre/style/textarea behavior, type 6/7 block boundaries, Markdown-vs-raw handling inside block HTML, and inline raw HTML classification.

- [ ] Implement a CommonMark-oriented HTML block scanner for all seven HTML block types.
- [ ] Preserve multiline opening tags and attributes without paragraph-wrapping them incorrectly.
- [ ] Respect raw-text containers such as `pre`, `script`, `style`, and `textarea`.
- [ ] Keep inline HTML inline when CommonMark expects paragraph parsing around it.
- [ ] Align raw HTML escaping, sanitizing, and GFM tag filtering with the profile options.
- [ ] Add syntax/native source fields for any new token boundaries introduced by the scanner.

Done means:

- [ ] The HTML block section has a passing/pinned or skipped-case inventory.
- [ ] Raw HTML has a passing/pinned or skipped-case inventory.
- [ ] Security profiles still prove strip/escape/sanitize behavior.

## Phase 3: Code Spans And Inline Precedence

Current observed failures include non-breaking-space trimming, multiline code-span normalization, code-span precedence inside link-looking text, and inline HTML/code precedence.

- [ ] Implement CommonMark code-span normalization exactly, including leading/trailing space stripping rules and line-ending collapse.
- [ ] Treat Unicode spaces according to the spec cases rather than ordinary ASCII trim rules.
- [ ] Fix precedence so code spans can defeat link parsing when the source requires it.
- [ ] Keep code-span delimiter/content source tokens stable for native editing.

Done means:

- [ ] All official code-span examples are passing and pinned or intentionally documented.
- [ ] Existing code-span source-edit and syntax-token tests still pass.

## Phase 4: Link, Reference, And Image Grammar

Current observed failures include invalid angle destinations with newlines, invalid title tails, nested link label precedence, NBSP destination/title separation, and image/link nesting rules.

- [ ] Rework inline link destination parsing around the CommonMark grammar instead of ad hoc balanced scanning.
- [ ] Rework optional title parsing so invalid extra tokens fall back to literal text.
- [ ] Rework nested link precedence so links cannot contain links where CommonMark forbids it.
- [ ] Unify inline links, reference links, shortcut/collapsed references, and images through one label/destination/title grammar.
- [ ] Preserve current source-backed delimiter, target, title, alt, and reference-definition metadata.

Done means:

- [ ] Links, link reference definitions, and images have passing/pinned or skipped-case inventories.
- [ ] Existing native source-edit coverage for links/images/reference definitions stays green.

## Phase 5: Emphasis And Strong Emphasis Algorithm

The largest remaining CommonMark section is emphasis: 121 unpinned examples.

- [ ] Replace simplified delimiter handling with a CommonMark delimiter-run algorithm.
- [ ] Support proper left/right flanking, punctuation, intraword underscore, nested delimiters, opener/closer balancing, and precedence with links/code.
- [ ] Reconcile GFM single-tilde strikethrough and OfficeIMO extensions with the CommonMark delimiter stack.
- [ ] Keep wrapper syntax nodes and native opening/closing marker metadata for editor hosts.

Done means:

- [ ] Emphasis and strong-emphasis official examples are passing and pinned or intentionally documented.
- [ ] GFM strikethrough cases still pass under the GFM profile.

## Phase 6: Tabs, Indentation, Blocks, And Breaks

- [ ] Implement tab expansion as a parser/source-map primitive, not a local string replacement.
- [ ] Finish indented-code, list-item, blockquote continuation, and hard-break edge cases against official examples.
- [ ] Keep original line/column/offset mapping stable for CRLF, CR, LF, tabs, and mixed indentation.
- [ ] Ensure list and blockquote source-edit tests still address original markers and content columns.

Done means:

- [ ] Tabs, indented code blocks, block quotes, list items, hard line breaks, blank lines, textual content, inlines, and precedence have passing/pinned or skipped-case inventories.

## Phase 7: GFM And Markdig Extension Breadth

Markdig parity is broader than CommonMark. The public Markdig target includes an extensible pipeline and many built-in extensions.

- [ ] Inventory enabled OfficeIMO/GFM extensions against cmark-gfm and Markdig specs.
- [ ] Expand pipe-table coverage beyond the current smoke set, including malformed delimiter and container interactions.
- [ ] Expand footnote rendering and source-span coverage against cmark-gfm.
- [ ] Expand autolink coverage for GFM and Markdig-style advanced autolinks.
- [ ] Decide which Markdig extensions are in-scope for OfficeIMO parity: grid tables, auto identifiers, YAML front matter, emoji, math, diagrams, SmartyPants, abbreviations, citations, custom containers, generic attributes, media links, task lists, and alerts.
- [ ] For each in-scope extension, decide whether it belongs in `OfficeIMO.Markdown`, `OfficeIMO.MarkdownRenderer`, or a separate extension package.
- [ ] Document every out-of-scope extension as intentional, not missing by accident.

Done means:

- [ ] There is an extension inventory table with `Covered`, `Partial`, `Intentional`, or `Gap` status for every Markdig extension family we care about.
- [ ] The GFM smoke lane is no longer only a curated sample.

## Phase 8: AST, Source Mapping, And Lossless Roundtrip

Markdig-class parity includes precise source locations and trivia-backed roundtrip behavior.

- [ ] Finish canonical AST cleanup for `ListItem`, `TableBlock`, `CalloutBlock`, `DefinitionListBlock`, and any remaining duplicated mutable node shape.
- [ ] Complete parser-owned trivia capture for whitespace, blank lines, tabs, delimiter trivia, and generated-node diagnostics.
- [ ] Complete original-to-normalized offset mapping for all parser paths.
- [ ] Expand `MarkdownRoundtripWriter` from unchanged documents and explicit native source edits toward general lossless edits.
- [ ] Make parser, transform, renderer, and writer extension APIs all source-slice aware where it matters.
- [ ] Keep public semantic APIs stable or document intentional breaking cleanup before merging.

Done means:

- [ ] Editor-grade claims are backed by native snapshots, source edits, syntax invariants, and roundtrip diagnostics across real mixed documents.

## Phase 9: Renderer, Security, And Performance

- [ ] Keep HTML rendering spec-compatible for CommonMark/GFM profiles and OfficeIMO-specific for OfficeIMO profile only where intentional.
- [ ] Keep raw HTML policy explicit: allow, strip, escape, sanitize, and GFM tag-filter behavior must be independently tested.
- [ ] Add release-mode benchmark snapshots for parse, parse-with-syntax, HTML render, Markdown write, source-edit roundtrip, and transforms.
- [ ] Compare representative README/docs/chat/transcript corpora against the local Markdig baseline.
- [ ] Track allocation and throughput regressions for large tables, long lists, nested inline-heavy documents, and raw HTML-heavy documents.

Done means:

- [ ] Compatibility and performance claims can be repeated from documented commands.

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
