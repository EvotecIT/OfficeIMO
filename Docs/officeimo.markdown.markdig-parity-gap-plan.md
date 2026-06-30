# OfficeIMO.Markdown Markdig Parity Plan

This is the working checklist for bringing `OfficeIMO.Markdown` to Markdig-class behavior without drifting into disconnected fixture work.

Parity is not "more tests." Tests are evidence. Parity means the parser, semantic AST, syntax/source model, native projection, renderer, Markdown writer, extension seams, security/profile behavior, docs, and benchmarks agree on a clear contract.

Use the Markdig extension inventory and Markdig extension compatibility matrix as the source boards for row status. This file is the short execution plan.

## Current State

- [x] Markdig comparison baseline is pinned to Markdig `1.3.2` across tests, benchmarks, and compatibility docs.
- [x] CommonMark `0.31.2` correctness is green: 652 of 652 official CommonMark `0.31.2` examples pass in the generated inventory; 0 are failing.
- [x] GFM tracked fixtures are green: 52 tracked GFM fixtures, 52 passing, 0 failing in the generated GFM inventory.
- [x] Markdig extension inventory exists: 33 Markdig extension-family rows.
- [x] Markdig extension compatibility matrix exists with Decision, Route, Scope decision, Engine parser, AST/source, Writer/render, Proof, and Next-action lanes.
- [ ] Markdig extension parity is not closed: 10 rows are covered, 9 partial, 3 intentional, and 11 gap.
- [ ] AST/source/lossless parity is not closed: full trivia, delimiter tokens, original-to-normalized mapping, generated-node diagnostics, broader source edits, and extension-node roundtrip still need work.
- [ ] Performance parity is not known: release-mode benchmark comparison should wait until parser/source/writer behavior stops moving.

## What Is Still Missing From Parity

This is the non-looping checklist. A test-only slice is valid only when the matching engine behavior already exists and the open lane is proof.

- [ ] `UseDefinitionLists` promotion: mostly engine plus proof.
  - [x] Engine: generated/rebuilt definition body wrappers are marked through `MarkdownSyntaxNode.IsGenerated`, and native projection emits `native.generated-definition-child` diagnostics when definition children are regenerated from semantic content.
  - [ ] Engine: broaden remaining lazy-continuation and nested-body parser cases where Markdig keeps lines literal or attaches them to definition bodies.
  - [ ] Writer: prove Markdown output reparses to equivalent definition-list behavior for the remaining edge cases.
  - [ ] Proof: refresh focused Markdig comparison, native/source-map, writer, and matrix evidence before changing the row to `Covered`.
- [ ] `UseGenericAttributes` promotion: engine plus proof.
  - [ ] Engine: extend attribute ownership from the already-covered shapes to the remaining Markdig-supported block and inline families.
  - [ ] Engine: keep attributes attached to the semantic owner and syntax/source owner through containers such as blockquotes, lists, tables, footnotes, and definition lists.
  - [ ] Writer/render: preserve attributes through HTML rendering and Markdown writing where the row claims support.
  - [ ] Proof: add target-by-target Markdig comparison and source/native fixtures only after each target exists in the engine.
- [ ] `UseAlertBlocks` decision: scope first, then engine.
  - [ ] Decide whether Markdig alert rendering callbacks become an OfficeIMO renderer contract or remain an intentional callout difference.
  - [ ] If in scope, align alert/callout parser behavior, syntax fields, native source fields, renderer customization, writer output, and comparison fixtures.
- [ ] `UseCjkFriendlyEmphasis` decision: parser option or intentional gap.
  - [ ] Decide whether to add a Markdig-compatible delimiter option.
  - [ ] If in scope, implement delimiter behavior in the inline parser and prove source-token/writer behavior with CJK comparison cases.
- [ ] `UsePreciseSourceLocation` promotion: cross-cutting AST/source architecture.
  - [ ] Engine: capture lossless trivia beyond current source slices: whitespace, blank lines, tabs, delimiters, raw slices, generated nodes, and normalized text.
  - [ ] Engine: define one original-to-normalized mapping story for CRLF/LF/CR, tabs, nested containers, transforms, generated nodes, and normalized paragraph text.
  - [ ] Engine: broaden source-edit support beyond the current native field and explicit-edit coverage.
  - [ ] Diagnostics: report precise fallback reasons when exact source preservation is unavailable.
  - [ ] Proof: add source-map and roundtrip tests after the mapping rules exist.
- [ ] Optional parser gaps need product scope before implementation.
  - [ ] `UseCustomContainers`: decide core versus optional package, then implement colon-fence parsing, child source mapping, renderer seams, and writer support.
  - [ ] `UseGridTables`: decide whether grid tables belong in core, then implement malformed fallback, source spans, renderer, and writer behavior.
  - [ ] `UseListExtras`: inventory Markdig list-extra syntax first, then choose core versus optional behavior.
  - [ ] `UseMathematics`: decide inline/block math ownership and renderer handoff before adding delimiter parsing.
  - [ ] `UseMediaLinks`: decide provider and safe-renderer policy before parser shortcuts.
  - [ ] `UseFigures`: separate HTML-import recovery from Markdown figure syntax before parser work.
  - [ ] `UseEmojiAndSmiley`, `UseJiraLinks`, and `UseSmartyPants`: keep as optional transforms unless a consumer needs them in core.
- [ ] Deferred rows stay out of the current parity push until a real consumer needs them.
  - [ ] `UseCitations`
  - [ ] `UseFooters`
  - [ ] `UseGlobalization`
  - [ ] `UsePragmaLines`
- [ ] Renderer/writer/security seams are still incomplete.
  - [ ] Build source-aware extension seams for custom block/inlines, transforms, renderers, and writers.
  - [ ] Separate raw HTML grammar from policy: CommonMark parsing, GFM tag filtering, OfficeIMO allow/strip/escape/sanitize modes, URL policy, source metadata, and Markdown writing.
  - [ ] Bound renderer-owned rows before parser work starts: `UseDiagrams`, `UseFigures`, `UseMathematics`, `UseMediaLinks`, and `UseReferralLinks`.
- [ ] Performance parity is intentionally last.
  - [ ] Add representative README/docs/chat corpora.
  - [ ] Compare parse, parse-with-syntax, HTML render, Markdown write, transforms, source edits, allocations, and throughput against Markdig after the moving parser/source/writer contracts settle.

## Promotion Gate

Do not move an extension row to `Covered` until each applicable box is true.

- [ ] Parser behavior matches Markdig for the scoped syntax, or the difference is intentional and documented.
- [ ] Semantic AST owns the behavior in stable OfficeIMO types.
- [ ] Syntax/source AST has spans for user-addressable parts, not only the outer block.
- [ ] Native/source-edit APIs can read and update the important fields without rescanning raw strings.
- [ ] HTML rendering matches the selected profile, or the rendering policy is explicit.
- [ ] Markdown writing reparses to equivalent behavior, or the writer limit is explicit.
- [ ] Security/profile behavior is separated from grammar when raw HTML, URLs, media, or renderer handoff is involved.
- [ ] Focused tests prove the current contract.
- [ ] Generated inventory and compatibility docs are refreshed.
- [ ] Broader fixture sweeps or benchmarks run only after the engine/source behavior is stable enough for the result to mean something.

## P0 - Finish Active Engine Promotion

Current active row: `UseDefinitionLists`.

- [x] Close the active nested-body equals-setext literal gap.
  - [x] Nested list lazy paragraph followed by `===` stays literal text when Markdig keeps it literal.
  - [x] Nested blockquote lazy paragraph followed by `===` stays literal text when Markdig keeps it literal.
  - [x] Markdown writer escapes or preserves the shape so OfficeIMO reparse does not create a heading by accident.
  - [x] AST/source/native spans describe the nested paragraph and definition body correctly.
- [ ] Broaden remaining lazy-continuation cases.
  - [ ] Paragraph-after-blank variants not already covered.
  - [ ] Multiple lazy lines after nested blocks.
  - [ ] Remaining list-like and table-like interruption starts, with pipe tables on and off.
- [ ] Broaden nested-body cases.
  - [ ] Blockquote source breadth beyond the already-covered heading/thematic/table-shaped cases.
  - [ ] Fenced-code variants beyond the marker-line and empty-marker boundary cases already closed.
  - [ ] List-tail variants after nested body boundaries.
    - [x] Mixed unordered-to-ordered list tails stay inside the definition body as separate list children, with syntax/native source spans and writer reparse proof.
    - [x] Unindented blockquote tails after nested list bodies now close the definition list like Markdig, with syntax/native source spans and writer reparse proof.
    - [x] Unindented list tails after nested blockquote bodies now close the definition list like Markdig, with syntax/native source spans and writer reparse proof.
- [ ] Finish definition-list source mapping.
  - [x] Marker lines are source-backed through parsed `DefinitionMarker` syntax tokens and native `definitionMarker` source fields; generated marker tokens remain source-less by design.
  - [x] Continuation indentation stripped from definition body lines now surfaces as native `definitionContinuationIndent` source fields with precise caret lookup.
  - [x] Blank separators now surface as native `definitionBlankLine` source fields with precise caret lookup while broad `definitionBody` spans remain available.
  - [x] Generated paragraph wrappers are now honest in the final syntax/native model: rebuilt semantic children carry `MarkdownSyntaxNode.IsGenerated`, and definition-list native projection reports `native.generated-definition-child` instead of presenting fallback anchors as exact parsed source.
  - [x] Normalized native `definitionBody` values versus original source spans are now explicit: `definitionBody.Value` stays semantic/normalized while `MarkdownNativeDocument` can materialize normalized or original source slices for the span-backed native field.
- [ ] Promote `UseDefinitionLists` only after parser behavior, AST/source/native projection, HTML rendering, Markdown writing, reparse stability, generated inventory, and the compatibility matrix all agree.

## P1 - Close High-Value Partial Rows

- [ ] Finish `UseGenericAttributes`.
  - [ ] Extend from covered shapes to arbitrary Markdig-supported block families.
  - [ ] Extend source-backed inline attributes across the remaining supported inline families.
  - [ ] Prove container interactions such as blockquotes, lists, tables, footnotes, and definition lists by contract, not incidental HTML output.
    - [x] Standalone generic attributes before pipe tables now target the semantic table, match Markdig HTML, expose syntax/native `attributes` source fields, and support source edits.
    - [x] Standalone generic attributes before image paragraphs match Markdig in portable profiles, and OfficeIMO-default typed image blocks now carry syntax/native/source-edit-backed attributes.
  - [ ] Keep writer behavior and source edits stable across attributed shapes.
- [ ] Decide and close `UseAlertBlocks`.
  - [ ] Decide whether Markdig alert rendering callbacks become an OfficeIMO renderer contract or remain an intentional callout difference.
  - [ ] Align callout/alert AST fields, source spans, renderer customization, writer output, and comparison fixtures around that decision.
- [ ] Decide and close `UseCjkFriendlyEmphasis`.
  - [ ] Add a Markdig-compatible delimiter option with CJK comparison/source-token proof, or document it as deferred/intentional.
- [ ] Keep `UsePreciseSourceLocation` partial until lossless trivia, original mapping, generated-node diagnostics, and broader source edits are complete.

## P2 - Make Scope Decisions Before Optional Syntax

Do not implement these rows from nearby tests alone. Decide the product shape first.

- [ ] `UseCustomContainers`: core extension seam plus optional colon-fenced container parser, child source ownership, renderer seams, and writer support.
- [ ] `UseGridTables`: optional grid-table parser, semantic table model, malformed fallback, source spans, renderer, and writer.
- [ ] `UseListExtras`: inventory Markdig list-extra syntax first, then decide whether it belongs in core list behavior or an optional extension.
- [ ] `UseMathematics`: decide inline/block math delimiter ownership, AST/source/native metadata, writer preservation, and renderer handoff.
- [ ] `UseMediaLinks`: decide provider model and safe renderer policy before parser behavior.
- [ ] `UseFigures`: separate HTML-import figure recovery from Markdown figure syntax.
- [ ] `UseDiagrams`: decide named diagram language mapping and renderer-package ownership.
- [ ] `UseReferralLinks`: keep as renderer policy unless a real parser contract appears.
- [ ] Optional transforms stay optional unless product need changes: `UseEmojiAndSmiley`, `UseJiraLinks`, `UseSmartyPants`.
- [ ] Deferred rows stay deferred until a consumer needs them: `UseCitations`, `UseFooters`, `UseGlobalization`, `UsePragmaLines`.
- [x] Intentional differences stay documented rather than cloned as Markdig switches: `UseAdvancedExtensions`, `UseBootstrap`, `UseSelfPipeline`.

## P3 - Close Editor-Grade AST And Source Parity

This is the difference between "renders like Markdig" and "is a super-duper Markdown app."

- [ ] Canonicalize duplicated semantic/syntax ownership for lists, tables, definition lists, callouts, footnotes, front matter, and extension nodes.
- [ ] Associate syntax nodes with semantic subobjects such as callout titles, list item paragraphs, definition groups/definitions, table rows/cells, and sequence-style inline wrappers.
- [ ] Capture lossless trivia: whitespace, blank lines, tabs, delimiters, raw slices, normalized text, and generated-node diagnostics.
- [ ] Complete delimiter-token coverage for emphasis extras, links/images, escapes/entities, breaks, HTML, footnotes, front matter, tables, and extension nodes.
- [ ] Establish one original-to-normalized mapping story for CRLF/LF/CR, tab expansion, nested containers, transformed nodes, generated nodes, and normalized paragraph text.
- [ ] Broaden `MarkdownRoundtripWriter` beyond unchanged documents and explicit native edits.
- [ ] Add precise fallback diagnostics when exact source preservation is unavailable.

## P4 - Keep Renderer, Writer, Extension, And Security Policy Explicit

- [ ] Build source-aware extension seams for custom blocks, inlines, transforms, renderers, and writers so downstream code does not rescan strings.
- [ ] Separate raw HTML grammar from security policy.
  - [ ] CommonMark raw HTML parsing.
  - [ ] GFM tag filtering.
  - [ ] OfficeIMO allow/strip/escape/sanitize modes.
  - [ ] URL policy.
  - [ ] Source metadata.
  - [ ] Markdown writing.
- [ ] Bound renderer/host rows before parser work starts: `UseDiagrams`, `UseFigures`, `UseMathematics`, `UseMediaLinks`, and `UseReferralLinks`.

## P5 - Proof And Performance Last

- [ ] Add tests only when they prove a parser, AST/source, native projection, renderer, writer, security, profile, extension-seam, or benchmark contract.
- [ ] Regenerate inventories and compatibility matrices after each promoted row.
- [ ] Broaden GFM fixture breadth after covered grammar/source behavior is stable.
- [ ] Run release-mode benchmarks after correctness, source mapping, and writer behavior stop moving.
- [ ] Compare parse, parse-with-syntax, HTML render, Markdown write, transforms, source edits, allocations, and representative README/docs/chat corpora against Markdig.

## No-Loop Execution Rule

Before starting a slice, pick exactly one active lane:

- [ ] One Markdig extension family.
- [ ] One AST/source/lossless architecture gap.
- [ ] One renderer/writer/security/profile contract.
- [ ] One proof-only inventory or benchmark pass after behavior exists.

If behavior is missing, improve the reusable engine first. If behavior exists but is unproven, add focused proof. If scope is unclear, make the scope decision before adding code or tests.
