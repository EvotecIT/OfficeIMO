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
- [ ] Markdig extension parity is not closed: 13 rows are covered, 6 partial, 3 intentional, and 11 gap.
- [ ] AST/source/lossless parity is not closed: full trivia, delimiter tokens, original-to-normalized mapping, broader source edits, and extension-node roundtrip still need work.
- [ ] Performance parity is not known: release-mode benchmark comparison should wait until parser/source/writer behavior stops moving.

## What Is Still Missing From Parity

This is the non-looping checklist. A test-only slice is valid only when the matching engine behavior already exists and the open lane is proof.
Each unchecked item should be treated as exactly one lane before work starts: engine behavior, AST/source model, writer/render behavior, product-scope decision, or proof. If a comparison exposes a behavior mismatch, fix the engine first; do not paper over it with tests.

- [x] `UseDefinitionLists` promotion: mostly engine plus proof.
  - [x] Engine: generated/rebuilt definition body wrappers are marked through `MarkdownSyntaxNode.IsGenerated`, and native projection emits `native.generated-definition-child` diagnostics when definition children are regenerated from semantic content.
  - [x] Engine: broaden remaining lazy-continuation and nested-body parser cases where Markdig keeps lines literal or attaches them to definition bodies.
  - [x] Writer: prove Markdown output reparses to equivalent definition-list behavior for the remaining edge cases.
  - [x] Proof: refresh focused Markdig comparison, native/source-map, writer, and matrix evidence before changing the row to `Covered`.
- [ ] `UseGenericAttributes` promotion: engine plus proof.
  - [ ] Engine: extend attribute ownership from the already-covered shapes to the remaining Markdig-supported block and inline families.
  - [ ] Engine: keep attributes attached to the semantic owner and syntax/source owner through containers such as blockquotes, lists, tables, footnotes, and definition lists.
  - [ ] Writer/render: preserve attributes through HTML rendering and Markdown writing where the row claims support.
  - [ ] Proof: add target-by-target Markdig comparison and source/native fixtures only after each target exists in the engine.
- [x] `UseAlertBlocks` decision: scope first, then engine.
  - [x] Add an opt-in Markdig-style alert HTML fallback for no-title GitHub alert syntax while preserving OfficeIMO titled callouts as richer AST semantics.
  - [x] Decide the titled-callout boundary: OfficeIMO mode keeps rich titled callouts; Markdig-compatible mode treats titled alert markers as ordinary blockquotes.
  - [x] Align the remaining callout/alert syntax fields, native source fields, writer output, and broader comparison fixtures around that boundary.
- [x] `UseCjkFriendlyEmphasis` decision: parser option or intentional gap.
  - [x] Add a Markdig-compatible opt-in delimiter option.
  - [x] Implement delimiter behavior in the inline parser and prove source-token/writer behavior with CJK comparison cases.
- [ ] `UsePreciseSourceLocation` promotion: cross-cutting AST/source architecture.
  - [x] Engine: expose addressable native block and snapshot source-field accessors so repeated fields such as quote/list markers can be selected by occurrence index without consumers rescanning or relying on lossy dictionaries.
  - [x] Diagnostics: expose reason-aware original-source slice failures and include the mapping reason in source-edit roundtrip fallback diagnostics.
  - [x] Engine: expose native inline and inline-metadata source-slice APIs so editor hosts can address inline content, link targets, titles, and similar metadata without rescanning raw markdown.
  - [x] Engine: align native source-slice APIs with source-edit targets for blocks, list item content, table cells, definition-list groups/terms/bodies, reference definitions, and reference-definition fields.
  - [x] Engine: expose paragraph-level native projections, source slices, and source edits for list-item paragraphs so editor hosts can address individual loose-list paragraphs instead of treating the whole item content as one span.
  - [x] Engine: reconcile parsed-compatible list-item paragraph syntax as source-backed when the cached literal still matches the semantic paragraph, so individual loose-list paragraph edits can preserve original source bytes instead of reporting generated-node fallback.
  - [x] Engine: expose document-level blank-line source trivia, snapshots, and source slices so editor hosts can address empty and whitespace-only lines without rescanning raw markdown.
  - [x] Engine: capture document-level leading/trailing horizontal whitespace source trivia, including tabs, with source-order enumeration, position lookup, snapshots, source slices, and tab-expanded source columns.
  - [x] Engine: align generic line/column source-slice fallback with source-map tab-expanded columns.
  - [x] Engine: capture document-level line-ending source trivia and map normalized line-ending spans back to original CRLF, LF, or standalone CR source slices and original-preserving source edits.
  - [x] Engine: expose pipe-table row delimiters as repeated native `tablePipe` source fields with source-order snapshots, caret lookup, and source edits while ignoring escaped and code-span pipes.
  - [x] Engine: expose source-order inline metadata snapshots for link/image fields and delimiter-like inline metadata, including escaped-character markers, decoded entity source text, and hard-break markers.
  - [x] Engine: expose link/image opening, separator, and closing delimiters through source-order inline metadata snapshots and source edits.
  - [x] Engine: expose emphasis-extra opening and closing delimiters for strikethrough, highlight, inserted, superscript, and subscript through source-order native metadata snapshots and source edits.
  - [x] Engine: expose footnote definition opening/separator markers and front-matter opening/closing fences through native source fields, snapshots, and source edits.
  - [x] Engine: expose code and semantic fenced-block opening/closing marker values through native source fields, snapshots, position lookup, and source edits.
  - [ ] Engine: capture lossless trivia beyond current source slices: whitespace, blank lines, tabs, delimiters, raw slices, generated-node roundtrip semantics, and normalized text.
  - [ ] Engine: define one original-to-normalized mapping story for CRLF/LF/CR, tabs, nested containers, transforms, generated nodes, and normalized paragraph text.
  - [ ] Engine: broaden source-edit support beyond the current native field and explicit-edit coverage.
  - [x] Diagnostics: report generated final syntax nodes at parse-result level, including path, index path, source fallback anchor, and associated semantic object details.
  - [x] Diagnostics: original-source slice APIs reject generated syntax-node fallback anchors, including native inline and inline-metadata slices, with a dedicated generated-node failure reason instead of treating them as exact original source.
  - [x] Diagnostics: native source edits created from generated block, inline, list-item, list-item paragraph, table-cell, definition-list, block-field, and metadata targets now carry generated-node original-source failure metadata so roundtrip diagnostics can stay honest while normalized edits still apply.
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
  - [x] Build source-aware extension seams for custom block/inlines, transforms, renderers, and writers.
    - [x] Custom block parser contexts can materialize normalized source slices for claimed spans and line ranges without rescanning raw Markdown.
    - [x] Custom inline parser contexts can materialize normalized source slices for claimed inline ranges without rescanning raw Markdown.
    - [x] Inline transform contexts can materialize normalized source slices for parsed inline nodes and spans without rescanning raw Markdown.
    - [x] Document transforms can materialize normalized and original source slices for parsed model objects, syntax nodes, and source spans without rescanning raw Markdown.
    - [x] Block/inline HTML render contexts and Markdown writer contexts can materialize normalized/original source slices, including original-source failure reasons, without rescanning raw Markdown.
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

Current active row: `UsePreciseSourceLocation`.

The `UseDefinitionLists` promotion below is closed and retained as evidence, not the next active lane.

- [x] Run the final compact Markdig comparison probe for definition-list tail behavior.
  - [x] Paragraph-after-blank tails after plain paragraph, nested unordered list, nested ordered list, and nested blockquote bodies.
  - [x] Two-or-more lazy lines after nested unordered list, nested ordered list, and nested blockquote bodies.
  - [x] Boundary candidates after blank lines: setext/thematic/heading, fenced code, raw HTML, reference-definition-looking text, pipe-table-shaped text, ordered/unordered/task-list starts.
  - [x] Run with pipe tables off and on where the input is table-shaped.
- [x] Fix only real `UseDefinitionLists` engine mismatches found by that probe.
  - [x] Parser ownership: blank-separated nested-list reference-definition-looking lazy tails stay literal inside the definition body like Markdig instead of inheriting root reference definitions.
  - [x] Semantic AST: the tail remains list-item paragraph text, not a consumed reference-definition/link workaround.
  - [x] Syntax/native source: the definition body keeps source-backed spans while normalized native values and writer output encode the colon for stable reparse.
- [x] Fix only real writer/reparse mismatches found by that probe.
  - [x] `ToMarkdown()` output reparses to equivalent HTML/AST behavior for the final compact definition-list probe.
  - [x] Literal text is preserved when Markdig keeps it literal, including setext-looking, table-looking, reference-looking, and escaped-pipe tails.
  - [x] Blank boundaries are inserted only where required to prevent semantic collapse on reparse.
- [x] Close the active nested-body equals-setext literal gap.
  - [x] Nested list lazy paragraph followed by `===` stays literal text when Markdig keeps it literal.
  - [x] Nested blockquote lazy paragraph followed by `===` stays literal text when Markdig keeps it literal.
  - [x] Markdown writer escapes or preserves the shape so OfficeIMO reparse does not create a heading by accident.
  - [x] AST/source/native spans describe the nested paragraph and definition body correctly.
- [x] Broaden remaining lazy-continuation cases.
  - [x] Paragraph-after-blank variants not already covered by the final compact probe.
    - [x] Blank-separated nested blockquote lazy tails preserve Markdig soft-break behavior with syntax/native source spans and writer reparse proof.
  - [x] Multiple lazy lines after nested blocks not already covered by the final compact probe.
    - [x] Multiple lazy lines inside a nested blockquote now stay in the definition body while a following unindented list closes the definition list like Markdig, with syntax/native source spans and writer reparse proof.
  - [x] Remaining list-like and table-like interruption starts, with pipe tables on and off.
    - [x] Compact Markdig comparison now matches for unordered, ordered, task-list-shaped, non-`1` ordered, ordered-parenthesis, escaped-pipe table-shaped, and pipe-table delimiter-mismatch tails across nested paragraph, list, ordered-list, and blockquote definition bodies.
- [x] Broaden nested-body cases.
  - [x] Blockquote source breadth beyond the already-covered heading/thematic/table-shaped cases.
    - [x] Unindented blockquote continuations remain inside active nested blockquotes while unindented fenced code, HTML, and reference-definition-looking lazy text follow Markdig ownership.
  - [x] Fenced-code variants beyond the marker-line and empty-marker boundary cases already closed.
    - [x] Unclosed fenced-code bodies consume lazy-looking trailing lines like Markdig and write a closing fence for stable reparse.
  - [x] List-tail variants after nested body boundaries.
    - [x] Mixed unordered-to-ordered list tails stay inside the definition body as separate list children, with syntax/native source spans and writer reparse proof.
    - [x] Non-`1` ordered lazy tails after nested lists stay inside the definition body, while non-`1` ordered list starts after active nested blockquotes close the definition list like Markdig.
    - [x] Ordered `)` lazy tails after ordered-list bodies split from `.` marker lists like Markdig and write stable Markdown.
    - [x] Unindented blockquote tails after nested list bodies now close the definition list like Markdig, with syntax/native source spans and writer reparse proof.
    - [x] Unindented list tails after nested blockquote bodies now close the definition list like Markdig, with syntax/native source spans and writer reparse proof.
    - [x] Unindented raw HTML after nested list bodies now closes the definition list like Markdig, with syntax/native source spans and writer reparse proof.
- [x] Finish definition-list source mapping.
  - [x] Marker lines are source-backed through parsed `DefinitionMarker` syntax tokens and native `definitionMarker` source fields; generated marker tokens remain source-less by design.
  - [x] Continuation indentation stripped from definition body lines now surfaces as native `definitionContinuationIndent` source fields with precise caret lookup.
  - [x] Blank separators now surface as native `definitionBlankLine` source fields with precise caret lookup while broad `definitionBody` spans remain available.
  - [x] Generated paragraph wrappers are now honest in the final syntax/native model: rebuilt semantic children carry `MarkdownSyntaxNode.IsGenerated`, and definition-list native projection reports `native.generated-definition-child` instead of presenting fallback anchors as exact parsed source.
  - [x] Normalized native `definitionBody` values versus original source spans are now explicit: `definitionBody.Value` stays semantic/normalized while `MarkdownNativeDocument` can materialize normalized or original source slices for the span-backed native field.
- [x] Promote `UseDefinitionLists` only after parser behavior, AST/source/native projection, HTML rendering, Markdown writing, reparse stability, generated inventory, and the compatibility matrix all agree.
  - [x] Focused definition-list tests are green.
  - [x] Broad CommonMark/GFM/Markdig comparison lane is green.
  - [x] Compact Markdig comparison matrix has zero failures for the final definition-list probe.
  - [x] Markdig extension inventory and compatibility docs mark `UseDefinitionLists` as `Covered` with current proof.
  - [x] No new source-map, native-projection, or writer diagnostics are hiding unresolved exactness gaps.

## P1 - Close High-Value Partial Rows

- [ ] Finish `UseGenericAttributes`.
  - [ ] Extend from covered shapes to arbitrary Markdig-supported block families.
  - [ ] Extend source-backed inline attributes across the remaining supported inline families.
  - [ ] Prove container interactions such as blockquotes, lists, tables, footnotes, and definition lists by contract, not incidental HTML output.
    - [x] Standalone generic attributes before pipe tables now target the semantic table, match Markdig HTML, expose syntax/native `attributes` source fields, and support source edits.
    - [x] Standalone generic attributes before image paragraphs match Markdig in portable profiles, and OfficeIMO-default typed image blocks now carry syntax/native/source-edit-backed attributes.
    - [x] Standalone generic attributes before reference-definition-looking lines now match Markdig by producing attributed literal paragraphs without registering reference definitions, with syntax/native/source-edit proof.
    - [x] Standalone generic attribute continuation lines at the end of paragraphs now match Markdig by being consumed without attributes or rendered output, including soft and hard line-break forms.
    - [x] Paragraph-contained attributes embedded at the end of nested link labels, image alt text, linked-image alt text, emphasis content, and strong content now promote to the paragraph owner like Markdig, with syntax/native source proof.
    - [x] No-space paragraph attributes now match Markdig around escaped final punctuation and valid character references, with syntax/native source proof.
    - [x] No-space paragraph attributes now match Markdig for unmatched trailing backtick runs while valid code spans still own inline attributes.
    - [x] ATX heading generic attributes now match Markdig when a closing marker appears before or after the attribute block, with source-backed closing-marker and `attributes` native fields.
    - [x] Fenced-code info-string attributes now parse attribute-only, language-plus-attribute, and opaque-info-prefix forms as metadata; ordinary code blocks render only the explicit `{...}` attribute block on `<code>` like Markdig while preserving opaque fence options for hosts, and expose source-backed native/snapshot `attributes` fields for code and semantic fenced blocks.
    - [x] List-contained ATX and loose nested headings now keep trailing generic attributes literal like Markdig, suppress automatic ids derived from that literal marker, and preserve fenced-code attributes inside list items with native source-field proof.
  - [ ] Keep writer behavior and source edits stable across attributed shapes.
- [x] Decide and close `UseAlertBlocks`.
  - [x] Add focused Markdig comparison proof for no-title note, list, and custom alert rendering through an opt-in Markdig-style HTML fallback.
  - [x] Make titled OfficeIMO callouts an intentional default with an explicit Markdig-compatible boundary mode.
  - [x] Expose GitHub alert header marker tokens (`[!` and `]`) as syntax/native source fields so editor hosts can address the full alert header without raw-string rescans.
  - [x] Prove curated Markdig-compatible alert cases can be written and reparsed back to equivalent Markdig alert HTML.
  - [x] Broaden curated no-title alert fixture coverage across standard GitHub alert kinds, rich inline bodies, nested quotes, fenced code, lists, custom kinds, and multi-paragraph bodies.
  - [x] Cover Markdig alert parser boundaries for empty alerts, lazy-continuation body lines, lowercase kinds, and malformed markers that stay blockquotes.
  - [x] Prove lazy-continuation alert body syntax/native source spans and source edits stay addressable across mixed unquoted and quoted body lines.
  - [x] Run the upstream-style GitHub alert sweep, including all five GitHub Docs examples, separated alert documents, paragraph boundaries, and nested-list blockquote boundaries.
  - [x] Align callout/alert AST fields, source spans, writer output, and broader comparison fixtures around that decision.
- [x] Decide and close `UseCjkFriendlyEmphasis`.
  - [x] Add a Markdig-compatible delimiter option with CJK comparison/source-token proof.
- [ ] Keep `UsePreciseSourceLocation` partial until lossless trivia, original mapping, generated-node roundtrip semantics, and broader source edits are complete.

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
  - [x] Parsed-compatible list-item paragraph syntax is now rebuilt as source-backed when the cached literal still matches the semantic paragraph, closing the generated-wrapper fallback for individual loose-list paragraph source edits.
- [ ] Associate syntax nodes with semantic subobjects such as callout titles, list item paragraphs, definition groups/definitions, table rows/cells, and sequence-style inline wrappers.
- [ ] Capture lossless trivia: whitespace, blank lines, tabs, delimiters, raw slices, normalized text, and generated-node roundtrip semantics.
- [ ] Complete delimiter-token coverage for extension nodes.
  - [x] Raw inline HTML fragments now expose exact native/snapshot `html` metadata and source edits so editor hosts can address the raw tag without rescanning paragraph text.
  - [x] Raw HTML block comments, tag frames, CDATA, declarations, and processing instructions expose source-backed opening/body/closing tag or marker fields with snapshot and source-edit proof.
  - [x] Pipe-table alignment rows now expose per-column `alignmentCell` native/snapshot source fields with occurrence indexes and source edits, so editor hosts can target one alignment marker without rewriting the whole row.
  - [x] Escaped-character markers, decoded entity source text, and hard-break markers now expose native inline metadata, source-order snapshot metadata fields, source slices, position lookup, and source edits.
  - [x] Inline link and image opening, separator, and closing delimiters now expose source-order native metadata snapshots, position lookup, and source edits.
  - [x] Code fences and semantic fenced extension nodes now expose value-bearing opening/closing fence markers in native source fields and snapshots, with position lookup and source-edit proof.
- [ ] Establish one original-to-normalized mapping story for CRLF/LF/CR, tab expansion, nested containers, transformed nodes, generated nodes, and normalized paragraph text.
- [ ] Broaden `MarkdownRoundtripWriter` beyond unchanged documents and explicit native edits.
- [ ] Add precise fallback diagnostics when exact source preservation is unavailable.
- [x] Expose parse-result generated syntax diagnostics so final syntax nodes rebuilt from semantic content are visible without native-projection-specific checks.
- [x] Return a dedicated original-source failure reason for generated syntax nodes so parse/native callers do not treat fallback anchors as byte-exact source.
- [x] Reject generated native inline and inline-metadata original-source slices before span-only mapping so editor hosts do not overclaim exact original bytes for regenerated inline content.
- [x] Carry generated-node original-source failure reasons on native source edits for generated block, inline, list-item, list-item paragraph, table-cell, definition-list, block-field, and metadata targets so roundtrip diagnostics do not overclaim exact original-source edits.
- [x] Carry known original-source failure reasons on native source edits at creation time, including missing preserved trivia and original-input equivalence failures, while keeping roundtrip diagnostics de-duplicated around the primary preserve-trivia message.
- [x] Expose live native-block and UI-safe snapshot source-field accessors so editor hosts can select repeated source fields by name and occurrence index without falling back to raw-string rescans.
- [x] Return exact original-source slice failure reasons for parse/native callers and include those reasons in roundtrip source-edit fallback diagnostics.
- [x] Expose native inline and inline-metadata source-slice APIs for normalized/original text so link targets, titles, formatting content, and similar inline source-backed values do not require raw-string rescans.
- [x] Expose source-slice APIs for native source-edit targets so blocks, list item content, table cells, definition-list objects, reference definitions, and reference-definition fields can be inspected before source edits are applied.
- [x] Expose document-level native abbreviation-definition source fields, snapshots, position lookup, normalized/original source slices, and source edits so abbreviation definitions match reference-definition editor affordances.
- [x] Expose native list-item paragraph projections, snapshots, inline runs, source-slice APIs, and source edits so loose-list paragraph edits can target a stable paragraph object instead of rescanning list item content.
- [x] Expose native document-level blank-line source trivia, snapshots, and normalized/original source-slice APIs so empty and whitespace-only lines are addressable without raw-string rescans.
- [x] Expose native document-level leading/trailing horizontal whitespace trivia, including tabs, with source-order enumeration, position lookup, snapshots, normalized/original source-slice APIs, and source-map-aligned tab-expanded columns.
- [x] Resolve offset-less line/column source slices with source-map-aligned tab-expanded columns so fallback slices do not drift on tabbed input.
- [x] Centralize tab-expanded source-column mapping so source maps, source slices, and document-level trivia share one column model instead of parallel local implementations.
- [x] Align native source-edit offset fallback with the shared tab-expanded column model so offset-less line/column edit spans do not drift on tabbed input.
- [x] Add shared visual-column start-offset mapping so prefix/trivia preservation code can slice text before a tab-expanded column without raw-character drift.
- [x] Expose native document-level line-ending trivia and original-source slices/edits that preserve CRLF, LF, and standalone CR spelling around line-ending-equivalent mapping.
- [x] Expose callout/alert opening and closing marker spans in syntax and native projections so alert header source ownership covers marker, kind, title, and body fields.

## P4 - Keep Renderer, Writer, Extension, And Security Policy Explicit

- [x] Build source-aware extension seams for custom blocks, inlines, transforms, renderers, and writers so downstream code does not rescan strings.
  - [x] Custom block parser contexts expose normalized source-slice helpers for parser-created spans and relative line ranges.
  - [x] Custom inline parser contexts expose normalized source-slice helpers for parser-created inline spans.
  - [x] Inline-transform contexts expose normalized source-slice helpers for source-backed inline nodes and spans.
  - [x] Document-transform contexts expose normalized/original source-slice helpers for associated model objects, syntax nodes, and source spans, sharing the same original-source mapping as parse results.
  - [x] Block/inline HTML renderer contexts and Markdown writer contexts expose normalized/original source-slice helpers plus original-source failure reasons for parsed model objects, syntax nodes, and token spans.
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
