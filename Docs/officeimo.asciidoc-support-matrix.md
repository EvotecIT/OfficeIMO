# OfficeIMO.AsciiDoc support matrix

- Status: experimental bounded-profile implementation (Phases 0-1)
- Updated: 2026-07-10
- Runtime dependencies: BCL and existing OfficeIMO project references only

This matrix describes the implemented contract. It is not a claim of complete Asciidoctor compatibility.

## Status meanings

- **Semantic:** source-backed, typed, editable, and covered by preserve/canonical writer behavior.
- **Source-preserved:** retained exactly in the syntax/source model without full evaluation.
- **Converted:** mapped to a typed Markdown or Reader representation.
- **Fallback:** emitted visibly with a diagnostic when the target has no equivalent.
- **Unsupported:** deliberately outside the current profile.

## Source, parser, and writer

| Capability | Native status | Current contract |
|---|---|---|
| Decoded character source | Semantic | `AsciiDocSourceText` retains the complete string and maps offsets to lines and columns. |
| Lossless syntax tree | Semantic | Explicit nodes and trivia cover every source character. Each node retains its exact source slice. |
| Mixed line endings and missing final newline | Semantic | LF, CRLF, CR, and the absence of a final line ending round-trip unchanged. |
| Original encoding and BOM | Unsupported | `Load` decodes text and `Save` writes UTF-8 using .NET defaults; byte-for-byte file reproduction is not promised. |
| Preserve writer | Semantic | An unchanged document is returned character-for-character. An edit regenerates only its owned subtree. |
| Canonical writer | Semantic | Recognized nodes use stable markers and a caller-selected line ending. |
| Recovery | Semantic | Known unterminated delimited blocks are diagnosed and remain source-preserved. |
| Variable-length delimiters | Semantic | Listing, literal, example, sidebar, quote, passthrough, and comment delimiters accept matching repeated-marker lines of four or more characters. This lets generated content contain the normal four-character fence unchanged. |
| Resource budgets | Semantic | Input length, block count, inline-node count, nesting, processing output, resources, and expansion are bounded. |
| Parser implementation | Semantic | Stateful scanners; no regular-expression parser or external process. |

## Blocks and metadata

| AsciiDoc feature | Native status | Markdown/Reader outcome | Current boundary |
|---|---|---|---|
| Document title and section headings | Semantic | Converted | Equals-sign headings through six markers; source-backed inline titles. |
| Document attributes | Semantic | Converted | Set/unset entries and effective values; set values can become Markdown front matter. |
| Paragraphs | Semantic | Converted | Multiline source and typed inline sequences. |
| Block attribute lists | Semantic | Converted where meaningful | Positional, named, ID, role, option, and `subs` values bind to the following block without duplicate writer ownership. |
| Block titles and anchors | Semantic | Converted | Titles have typed inlines; `[[id,reftext]]` metadata binds to its block. |
| Ordered and unordered lists | Semantic | Converted | Marker-derived depth is retained and editable. |
| Description lists | Semantic | Converted | Terms, definitions, depth, and typed inline content are retained. |
| Compound list continuations | Semantic | Converted | `+` continuation nodes bind following paragraphs or delimited blocks to the owning item. |
| Admonition paragraphs | Semantic | Converted | NOTE, TIP, IMPORTANT, WARNING, and CAUTION are typed. |
| Styled admonition blocks | Semantic | Converted | An admonition style on a compatible delimited block is exposed semantically. |
| Listing and literal blocks | Semantic | Converted | Listing/source content maps to code; literal content maps to text code. |
| Quote blocks | Semantic | Converted | Maps to Markdown quote semantics. |
| Example, sidebar, and open blocks | Semantic | Simplified with diagnostics | Source structure is retained; Markdown uses the closest named/container representation available. |
| Passthrough blocks | Semantic | Fallback | Source remains exact; Markdown receives an explicit `asciidoc` fallback. |
| Line and block comments | Semantic | Omitted or fallback | Reader/Markdown inclusion is option-controlled. |
| Block macros | Semantic | Converted or fallback | Image macros convert; unknown macros remain typed, editable, visible, and diagnosed. |
| Unknown source | Source-preserved | Visible fallback/paragraph | Unsupported input is not silently discarded. |

## Inline syntax

| Feature | Native status | Conversion boundary |
|---|---|---|
| Strong, emphasis, monospace | Semantic | Constrained and unconstrained forms, nesting, escaping, and edits are retained and converted. |
| Attribute references | Semantic | Typed source nodes; evaluated values are available through bounded substitution APIs. |
| Cross-references and anchors | Semantic | Typed and editable; mapped to Markdown links/anchors where representable. |
| Links, images, and general inline macros | Semantic | Common link/image forms convert; unknown macro names retain source and produce a fallback diagnostic. |
| Inline STEM | Semantic | Retains AsciiDoc math source and converts through Markdown's semantic math carrier with diagnostics where layout is unavailable. |
| Inline passthrough | Semantic | Exact source is retained; conversion is explicit and may fall back. |
| Superscript, subscript, mark, footnote, UI/callout macros | Source-preserved | Not yet bound to dedicated semantic types. |

## Tables

| Feature | Native status | Current boundary |
|---|---|---|
| PSV tables | Semantic | Default and custom separators, escaped separators, row discovery, and editable cells. |
| CSV and DSV tables | Semantic | Quoting/escaping, row discovery, and safe cell regeneration. TSV is represented by the delimited format model. |
| Column count/specification | Semantic | Common `cols` forms determine the structured grid. Complex width/layout meaning is retained but not typeset. |
| Header option | Semantic | Header rows are marked and mapped to structured Markdown tables. |
| Row/column spans | Semantic | Span prefixes are retained and carried through the structured Markdown bridge. Reverse conversion counts logical columns, including column spans. |
| Cell alignment and style | Semantic | Parsed and retained; target renderers may simplify unsupported layout/style. |
| Nested AsciiDoc blocks inside `a` cells | Source-preserved | Cell source is retained, but recursive block parsing inside a cell is not implemented. |
| Footer and advanced table layout | Source-preserved | No dedicated semantic behavior yet. |

## Processing and extensibility

Processing is explicit: `AsciiDocDocument.Parse` never reads another file or executes an extension merely because the source requests it.

| Feature | Status | Current contract |
|---|---|---|
| Attribute set/unset and references | Semantic | Case-insensitive effective set, configurable undefined behavior, cycle detection, and expansion limits. |
| Ordered substitution plans | Semantic | Named defaults and `subs` overrides retain the mandated substitution order. The API exposes the plan; it does not emulate every Asciidoctor replacement. |
| `ifdef`, `ifndef`, `ifeval`, `endif` | Semantic processing | Explicit preprocessing with deterministic diagnostics. |
| Includes | Semantic processing, disabled by default | Requires a caller-supplied resolver. The built-in file resolver is root-confined and rejects URI, absolute, traversal, and symbolic-link escapes by default. |
| Include selection | Semantic processing | Line selection, named tags, `*`/`**` wildcards, named and wildcard exclusions, nested tag markers, and level offsets, with cycle/depth/count/byte/output limits. |
| Registered directives | Semantic processing | Callers may register bounded in-process .NET processors. Built-ins are reserved. Documents cannot load code or assemblies. |
| General extension ecosystem | Unsupported | No Ruby/JavaScript processors, dynamic plugins, or document-controlled assembly loading. |

## Integration surfaces

| Surface | Status | Contract |
|---|---|---|
| `OfficeIMO.AsciiDoc` | Implemented | Native BCL-only engine; it does not depend on Markdown or Reader. |
| `OfficeIMO.AsciiDoc.Markdown` | Implemented both directions | Forward conversion is loss-aware. Reverse conversion generates canonical AsciiDoc, chooses collision-free delimited-block fences, and reparses it through the lossless engine. |
| `OfficeIMO.Reader.AsciiDoc` | Implemented | Modular `.adoc`, `.asciidoc`, and `.asc` path/stream handler with hierarchy, source locations, tables, compound-list ownership, whole-document projection without duplicated attached blocks, and diagnostics. |
| Word, HTML, and PDF | Available through Markdown | Semantic OfficeIMO output, not Asciidoctor rendering parity. Unsupported source remains diagnosed or visible according to conversion options. |

## Deliberate limits

The current profile does not promise full Asciidoctor substitution parity, every inline macro, generated indexes/lists, callout correlation, recursive AsciiDoc table cells, remote includes, ecosystem extensions, or Asciidoctor-identical HTML/PDF layout. Those features require individual semantic, security, writer, and conversion contracts before they can be marked supported.
