# OfficeIMO.Markdown Compatibility Notes

This repo has two related pieces:

- `OfficeIMO.Markdown`: Markdown object model (builder), a lightweight Markdown reader (`MarkdownReader`), and HTML rendering (`ToHtmlFragment` / `ToHtmlDocument`).
- `OfficeIMO.MarkdownRenderer`: a WebView/browser-oriented layer that wraps `OfficeIMO.Markdown` output into a reusable HTML "shell" (CSS + Prism + Mermaid) and supports fast incremental updates (chat/edit scenarios).

The intent is "GitHub-like" output for the subset we use in documentation and a chat app, without pulling in a full CommonMark/GFM engine at runtime.

## Recommended Stack (Docs + Chat)

- For docs pipelines (static HTML): use `OfficeIMO.Markdown` directly via `MarkdownReader.Parse(...).ToHtmlDocument(...)` / `ToHtmlFragment(...)`.
- For chat apps (WinUI/WebView2): use `OfficeIMO.MarkdownRenderer`:
  - Navigate WebView2 once to `MarkdownRenderer.BuildShellHtml(...)`.
  - On message updates, compute `bodyHtml = MarkdownRenderer.RenderBodyHtml(markdown, options)` and call `ExecuteScriptAsync(MarkdownRenderer.BuildUpdateScript(bodyHtml))`.
  - Mermaid diagrams are rendered client-side from fenced blocks named `mermaid` (for example, ```` ```mermaid ````).

## What We Support Today (Reader + HTML)

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
- Angle-bracket autolinks (`<https://...>` and `<user@example.com>`)

Lists note:

- Supports multi-paragraph list items (blank line + indented continuation) and mixed nesting (ordered lists within unordered list items and vice versa).
- Supports common nested blocks inside list items: fenced code blocks, indented code blocks, blockquotes, and tables (when indented under the list item).
- Supports nested `<details>` blocks inside list items when HTML blocks are enabled.

Renderer:

- Built-in HTML styles (Clean, GitHub Light/Dark/Auto, Word-ish)
- Optional Prism syntax highlighting (online/offline delivery modes)
- `OfficeIMO.MarkdownRenderer`: Mermaid bootstrap + incremental DOM updates suitable for WebView2

## Known Gaps vs CommonMark / GFM Expectations

These are the main reasons you will see differences compared to typical CommonMark/GFM expectations in real-world README files:

- Tables
  - Headerless tables are intentionally conservative to reduce false positives: they require outer pipes on every row and at least 2 rows.
  - Escaped pipes (`\\|`) and pipes inside code spans are handled for common cases, but deep edge cases still exist (especially when mixing HTML, backslashes, and unusual backtick fences).
- Lists
  - Continuation lines, multi-paragraph items, and several nested block types are supported, but complex nesting rules are not fully CommonMark-compliant.
- Blockquotes
  - "Lazy continuation" and some nesting/interaction with other blocks differs from CommonMark.
- Inline emphasis rules
  - Delimiter-run rules (nesting, intraword `_`, etc.) are simplified and can differ from CommonMark output.
- Autolinks
  - Literal autolinks are limited (primarily `http(s)://...` scanning).
- Code blocks
  - Some CommonMark edge cases around indentation and list nesting are not fully covered (fenced code is the most reliable form).
- HTML
  - Inline HTML and HTML blocks are intentionally optional; for chat-like untrusted scenarios they should remain disabled.

## Diagrams / Charts Strategy

Markdown itself does not "render diagrams"; it usually encodes them as fenced code blocks plus a client-side renderer.

What we do today:

- Mermaid: supported via `OfficeIMO.MarkdownRenderer` (fenced blocks named `mermaid` are rendered in the WebView).

What to add next (if needed for IntelligenceX.Chat):

- Additional diagram engines (PlantUML, Graphviz) using the same approach: keep them as fenced code blocks in Markdown, render via client-side JS or a server-side renderer if required.
- "Charts" (for example via Chart.js) by agreeing on a fenced block format (JSON/YAML) and adding a renderer that transforms those blocks into `<canvas>` plus JS.

## Suggested Roadmap (To Get Closer to GFM)

If the goal is "good enough for GitHub-like content" (not full spec compliance), these are the highest-impact improvements:

1. Robust table parsing (escaped pipes, code spans, alignment, trimming rules).
2. List continuation lines and multi-paragraph list items.
3. Blockquote + list interactions closer to CommonMark.
4. Improved autolink coverage (angle-bracket and email forms).
5. Inline emphasis delimiter rules (reduce surprises in real-world Markdown).
