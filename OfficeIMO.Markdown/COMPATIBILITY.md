# OfficeIMO.Markdown Compatibility Notes

This repo has two related pieces:

- `OfficeIMO.Markdown`: Markdown object model (builder), typed Markdown reader (`MarkdownReader`), AST/query helpers, and HTML rendering (`ToHtmlFragment` / `ToHtmlDocument`).
- `OfficeIMO.MarkdownRenderer`: a WebView/browser-oriented layer that wraps `OfficeIMO.Markdown` output into a reusable HTML "shell" (CSS + Prism + Mermaid) and supports fast incremental updates (chat/edit scenarios).

The intent is "GitHub-like" output plus a practical typed reader/AST for documentation and chat scenarios, without pulling in a full CommonMark/GFM engine at runtime.

For the current standards, extension, AST, and editor-readiness scoreboard, see
[`Docs/officeimo.markdown.compatibility-matrix.md`](../Docs/officeimo.markdown.compatibility-matrix.md).

## Recommended Stack (Docs + Chat)

- For docs pipelines (static HTML): use `OfficeIMO.Markdown` directly via `MarkdownReader.Parse(...).ToHtmlDocument(...)` / `ToHtmlFragment(...)`.
- For chat apps (WinUI/WebView2): use `OfficeIMO.MarkdownRenderer`:
  - Navigate WebView2 once to `MarkdownRenderer.BuildShellHtml(...)`.
  - On message updates, compute `bodyHtml = MarkdownRenderer.RenderBodyHtml(markdown, options)` and call `ExecuteScriptAsync(MarkdownRenderer.BuildUpdateScript(bodyHtml))`.
  - Mermaid diagrams are rendered client-side from fenced blocks named `mermaid` (for example, ```` ```mermaid ````).

## What We Support Today (Reader + HTML)

In addition to parsing and HTML rendering, the reader now exposes a typed document/query surface that includes:

- top-level blocks including front matter (`TopLevelBlocks`)
- depth-first traversal (`DescendantsAndSelf()`, `DescendantsOfType<T>()`)
- list-item traversal (`DescendantListItems()`)
- heading traversal and resolved anchors (`DescendantHeadings()`, `GetHeadingInfos()`, `GetHeadingAnchor(...)`, `FindHeading(...)`, `FindHeadings(...)`, `FindHeadingByAnchor(...)`)
- front matter entry/value helpers (`HasDocumentHeader`, `FrontMatterEntries`, `FindFrontMatterEntry(...)`, `TryGetFrontMatterValue<T>(...)`)

Block-level:

- ATX headings (`#` .. `######`)
- Setext headings (`Title` + `---/===` underline)
- Paragraphs
- Horizontal rule
- Fenced code blocks (``` or ~~~), with optional language
- Indented code blocks (4-space indented)
- Blockquotes
- Tables (pipe tables with optional alignment row)
- Lists
- Task lists (`- [ ]` / `- [x]`)
- Definition lists (`Term: Definition`)
- Footnotes (`[^id]` refs and `[^id]:` definitions)
- GitHub/DX-style callouts inside quotes (`> [!WARNING] Title`)
- YAML-ish front matter (`---` ... `---`)
- Limited HTML blocks (optional, controlled by `MarkdownReaderOptions`)

Inline:

- Links (inline + reference-style)
- Images and linked images
- Emphasis: italic, bold, bold+italic
- Strikethrough (`~~text~~`)
- Inline code spans (single or multi-backtick)
- Hard breaks (from explicit line breaks / `<br>` when inline HTML is enabled)
- Angle-bracket autolinks (`<https://...>`, `<ftp://...>`, `<mailto:user@example.com>`, and `<user@example.com>`)
- Literal autolinks in text: `http(s)://...`, `www.example.com`, `user@example.com`, and selected GFM bare schemes such as `mailto:` and `xmpp:` when enabled (configurable via `MarkdownReaderOptions`)

Lists note:

- Supports multi-paragraph list items (blank line + indented continuation) and mixed nesting (ordered lists within unordered list items and vice versa).
- Supports common nested blocks inside list items: fenced code blocks, indented code blocks, blockquotes, and tables (when indented under the list item).
- Supports nested `<details>` blocks inside list items when HTML blocks are enabled.
- CommonMark-style post-marker padding is now handled more accurately, so cases like `-    one` distinguish between shallow continuations that fall out of the list and deep-enough continuations that stay inside the item.
- The CommonMark profile also now handles list items whose first block is indented code more like the official examples, including the extra-leading-space distinction in cases like `1.     indented code` versus `1.      indented code`.
- The CommonMark profile now also handles blank-start and empty list items more like the official examples, including marker-only items, tab-after-marker empty items, and the rule that empty list items cannot interrupt paragraphs.
- The CommonMark profile now treats nested lists and headings as valid first blocks inside list items more like the official examples, and it is stricter about not nesting sublists under wide ordered markers or shallowly indented sibling markers when the continuation indent is too shallow.
- The CommonMark profile now also respects list-type boundaries and loose-list grouping more like the official examples: changing bullet markers or ordered delimiters starts a new list, while blank lines between same-type items keep a single loose list instead of splitting it apart.
- Loose-list HTML is also closer to the official CommonMark examples for empty items: blank loose items now render as empty `<li></li>` nodes instead of paragraph-wrapped empty content.
- Tight-list fenced code items now preserve blank lines inside the fenced content more like the official CommonMark examples, without leaking later sibling list markers into the code block.
- Block-leading list items now also become loose when a blank line separates successive child blocks, so cases like fenced-code-plus-paragraph items render closer to the official CommonMark examples.

Renderer:

- Built-in HTML styles (Clean, GitHub Light/Dark/Auto, Word-ish)
- Optional Prism syntax highlighting (online/offline delivery modes)
- Optional GitHub-style task-list and footnote HTML via `HtmlOptions.GitHubTaskListHtml` and `HtmlOptions.GitHubFootnoteHtml`
- Reader extension seams for block parsers, fenced-block semantic factories, inline parsers, and post-parse inline AST transforms
- Type-targeted block and inline HTML override registration via `HtmlOptions.BlockRenderExtensions` and `HtmlOptions.InlineRenderExtensions`, including context-aware registrations for body/write context and source-span-aware AST rendering
- Type-targeted block and inline Markdown serialization overrides via `MarkdownWriteOptions.BlockRenderExtensions` and `MarkdownWriteOptions.InlineRenderExtensions`, including context-aware registrations for body/write context and source-span-aware AST rendering
- `OfficeIMO.MarkdownRenderer`: Mermaid bootstrap + incremental DOM updates suitable for WebView2

## Known Gaps vs CommonMark / GFM Expectations

These are the main reasons you will see differences compared to typical CommonMark/GFM expectations in real-world README files:

- Tables
  - Headerless tables are intentionally conservative to reduce false positives in the OfficeIMO/default profile: they require outer pipes on every row and at least 2 rows. The GFM profile requires a delimiter row and keeps table cells inline-only, matching cmark-gfm more closely.
  - Escaped pipes (`\\|`) and pipes inside code spans are handled for common cases, but deep edge cases still exist (especially when mixing HTML, backslashes, and unusual backtick fences).
- Lists
  - Continuation lines, multi-paragraph items, wide post-marker padding, and several nested block types are supported, but complex nesting rules are not fully CommonMark-compliant.
- Blockquotes
  - "Lazy continuation" and some nesting/interaction with other blocks differs from CommonMark.
- Inline emphasis rules
  - Delimiter-run rules (nesting, intraword `_`, etc.) are simplified and can differ from CommonMark output.
- Autolinks
  - Literal autolinks cover common cases (`http(s)://...`, `www.*`, plain emails, selected GFM bare schemes such as `mailto:` and `xmpp:`, and angle-bracket absolute URIs like `mailto:`, `ftp://`, `tel:`, and `urn:`) but do not aim for full spec coverage.
  - The GitHub Flavored Markdown profile now matches cmark-gfm's single-tilde strikethrough and `www.*` autolink baseline more closely by treating `~text~` as strikethrough and resolving `www.*` links with `http://`. Double-tilde input now keeps the `~~...~~` delimiter frame instead of degrading into nested single-tilde spans, while longer and mismatched tilde runs stay literal.
  - The GFM parity lane also exercises cmark-gfm-style footnote rendering and the `text![^id]` punctuation case more closely, including paragraph interruption by later footnote definitions. The smoke corpus now also protects escaped pipes, code-span pipes, escaped pipes inside table-cell code spans, broader table backslash escaping, one-column delimiter rows, paragraph-to-table boundaries, reference links inside table cells, adjacent empty cells, compact inline emphasis inside tables, inline formatting in table headers/body cells, non-table pipe-row rejection, minimal header-only tables, raw inline HTML and break tags inside table cells, the cmark-gfm HTML tag filter, nested task lists, uppercase checked task markers, task-marker whitespace boundaries, list/task marker source spans, plus-tag email local parts, invalid email-like tokens, bare `mailto:`/`xmpp:` autolinks, Unicode URL destinations, `www` host underscore rules, quoted/trailing-punctuation autolinks, the upstream ignored malformed-email crash regression, nested emphasis and delimiter-run edge cases inside strikethrough, and footnote ordering by first reference.
  - The CommonMark smoke lane now covers multiline reference link definitions, multiline reference labels, Unicode-aware label folding for shortcut/full references, invalid-inline-link fallback to shortcut references, chained reference precedence/backtracking, percent-encoding of non-ASCII link destinations, more code-span delimiter cases, underscore/digit-adjacent emphasis, broader absolute-URI autolinks, expanded official HTML block behavior for raw tables, type 1 blocks, comments, processing instructions, CDATA, and paragraph interruption, backslash-escaped punctuation, named entities, paragraph blank-line/indentation behavior, hard and soft line breaks, false thematic-break markers, escaped ATX closing marker characters, setext/container interactions, compact nested blockquotes, blockquote lazy-continuation cases, nested blockquote list-continuations, quote marker source spans, list/code boundaries, shallow list indentation, and HTML-comment list boundaries more directly. The final rebuilt AST also no longer leaks definition-source spans into resolved inline link metadata, and final-tree exact-span lookup is more reliable when sibling inline nodes share a boundary.
  - For a more portable baseline, use `MarkdownReaderOptions.CreatePortableProfile()` to turn off bare `http(s)`, `www`, and plain-email autolinking and disable OfficeIMO-only callout/task-list parsing while keeping explicit links, angle autolinks, and plain lists.
- Images
  - The OfficeIMO/default profile promotes standalone markdown image lines into typed `ImageBlock` nodes. CommonMark, GFM, and portable reader profiles now keep those lines as paragraph inline images so the spec-oriented HTML shape stays closer to CommonMark.
  - Parsed image descriptions now flatten inline formatting and nested link/image content down to plain-string HTML `alt` text more like the official CommonMark image examples, while the syntax tree still preserves the raw source form of the image label.
- Extension model
  - The parser/renderer architecture is much cleaner than before, with public block parser, fenced block, inline parser, inline transform, and context-aware type-targeted renderer/writer override seams. It is still not as broad as mature dedicated markdown engines because syntax-node-shape renderers and lossless token/trivia extension points are not first-class yet.
- Spec breadth
- We now cover a much larger compatibility set than the earlier subset reader, and the test suite now includes 172 pinned CommonMark 0.31.2 fixtures, 33 cmark-gfm smoke fixtures, and focused upstream non-render regression coverage with selected AST path/span assertions in addition to curated Markdig 1.3.2 parity cases. A package guardrail keeps the Markdig test and benchmark baselines aligned, but that is still not the same thing as full CommonMark/GFM conformance.
- Syntax-backed parse results now expose normalized source slices for span-backed nodes, including fenced-code opening/info/content/closing tokens with nested blockquote/list source-map coverage, hard-break marker spellings, inline footnote reference delimiters, inline/reference-style link and image opening/separator/closing delimiters, linked-image delimiters, source-backed autolink targets/angle markers, and supported inline HTML tag opening/closing marker tokens with nested inline source-map coverage. `MarkdownReaderOptions.PreserveTrivia` can retain the raw reader input as parse-result metadata for future lossless work. Line-ending-equivalent original input, including CRLF and standalone CR, can also materialize original source slices through line/column coordinates. `MarkdownRoundtripWriter` can preserve unchanged trivia-backed parse results byte-for-byte, apply explicit native source edits to original input when every edit remaps safely, and report diagnostics when it falls back to generated or normalized markdown, but full trivia capture and general byte-preserving edit writing are still not implemented.
  - URL normalization is now closer to the official CommonMark examples for non-ASCII link destinations, but broader URI normalization and edge-case destination parsing still need wider corpus coverage.
- Code blocks
- Some CommonMark edge cases around indentation and list nesting are not fully covered, though the pinned CommonMark lane now includes trickier list-padding/code-boundary examples, shallow-indentation list examples, HTML-comment list boundaries, blockquote/list continuation cases including nested blockquote list items, and list-boundary/loose-list and empty-loose-item cases in addition to the earlier list-item smoke cases (fenced code is still the most reliable form).
- HTML
  - Inline HTML and HTML blocks are intentionally optional; for chat-like untrusted scenarios they should remain disabled.

## Diagrams / Charts Strategy

Markdown itself does not "render diagrams"; it usually encodes them as fenced code blocks plus a client-side renderer.

What we do today:

- Mermaid: supported via `OfficeIMO.MarkdownRenderer` (fenced blocks named `mermaid` are rendered in the WebView).

What to add next (if needed for the chat app):

- Additional diagram engines using the same approach: keep them as fenced code blocks in Markdown, render via client-side JS or a server-side renderer if required.
- More chart formats by agreeing on a fenced block format (JSON/YAML) and adding a renderer that transforms those blocks into `<canvas>` plus JS.

## Suggested Roadmap

If the goal is broader standards coverage rather than just "good enough for GitHub-like content", these are the highest-impact improvements:

1. Expand the pinned CommonMark/GFM corpora beyond the initial official CommonMark 0.31.2 and cmark-gfm smoke lanes plus curated parity cases.
2. Stronger extension/plugin seams for custom parsers and renderers.
3. More delimiter-run / inline edge-case coverage.
4. Benchmarks on representative docs/chat corpora.
5. Continued cleanup of any remaining string-heavy surfaces in the public model.
