# OfficeIMO.Markdown compatibility

`OfficeIMO.Markdown` provides named CommonMark, GitHub Flavored Markdown, portable, and OfficeIMO profiles. The profile selects grammar and defaults; HTML security and host behavior remain explicit renderer options.

The repository-level [compatibility matrix](../Docs/officeimo.markdown.compatibility-matrix.md) records current coverage and exact limits. Generated [CommonMark](../Docs/officeimo.markdown.commonmark-inventory.md) and [GFM](../Docs/officeimo.markdown.gfm-inventory.md) inventories record the fixture evidence.

## Profiles

### CommonMark

Use `MarkdownReaderOptions.CreateCommonMarkProfile()` when standards grammar and HTML shape matter more than OfficeIMO convenience behavior. The profile disables OfficeIMO-only callouts, task-list promotion, front matter, bare URL/email discovery, and standalone image-block promotion unless the caller opts into them.

### GitHub Flavored Markdown

Use `MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile()` for tables, task lists, strikethrough, autolinks, and the repository's tracked GFM behavior. Pair it with `HtmlOptions.CreateGitHubFlavoredMarkdownProfile()` when GitHub-oriented HTML details and dangerous-tag filtering are required.

### Portable

Use `MarkdownReaderOptions.CreatePortableProfile()` for explicit links, angle autolinks, plain lists, and conservative cross-renderer behavior. Bare URLs, plain-email links, task-list promotion, and OfficeIMO callouts remain disabled.

### OfficeIMO

The default profile enables OfficeIMO conveniences such as typed image blocks, callouts, front matter, semantic fenced blocks, and selected bare-link behavior. These additions remain identifiable in the typed model and can be disabled through reader options.

## Current standards evidence

- 316 CommonMark 0.31.2 examples are pinned as smoke fixtures.
- The generated full CommonMark inventory reports 651 of 652 examples matching.
- The generated GFM inventory reports 52 tracked fixtures and 52 passing.
- Focused tests cover syntax spans, native projection, source edits, HTML rendering, Markdown writing, profiles, extensions, diagnostics, and round trips.

## Source and round-trip behavior

- `MarkdownDoc` is the semantic editing model.
- `MarkdownParseResult.FinalSyntaxTree` is the source-oriented structural model.
- Normalized source slices are available for span-backed nodes.
- `PreserveTrivia` retains original input for lossless workflows.
- `MarkdownRoundtripWriter` preserves unchanged trivia-backed documents byte-for-byte and applies native source edits when every edit maps safely.
- Writer fallbacks are diagnosed; arbitrary tree mutation is not presented as universally byte preserving.

## Extension behavior

The public extension surface supports ordered block parsers, fenced-block parsers, inline parsers, post-inline transforms, type-targeted renderer/writer overrides, and syntax-kind-targeted overrides. Extension contexts can inspect semantic nodes, final syntax nodes, and available normalized/original source slices.

Optional syntax becomes part of a named profile only when parser behavior, semantic ownership, syntax/source mapping, rendering, writing, security policy, and focused evidence agree. Host features such as diagrams and media resolution remain host-owned unless a format-level contract is defined.

## Security boundary

Raw HTML recognition, raw HTML output, URL policy, image/resource policy, escaping, sanitization, and browser hosting are separate choices. Parsing untrusted Markdown does not implicitly authorize raw HTML or remote resource access.

## Known limits

- One CommonMark raw-HTML inventory case remains open.
- Complete trivia/delimiter coverage and original-to-normalized mapping are not available for every node and transform.
- General byte-preserving output after arbitrary multi-node or generated-node edits is not claimed.
- Optional grid tables, math, media, figures, and diagram languages are not default grammar.
