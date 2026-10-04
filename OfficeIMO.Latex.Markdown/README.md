# OfficeIMO.Latex.Markdown

This package maps the bounded `OfficeIMO.Latex` profile to and from `OfficeIMO.Markdown`. It reports source fallbacks and simplifications, especially for TeX math layout, package-specific commands, bibliography formatting, and unknown environments.

`LatexDocumentProfile.PreserveOnly` retains the complete visible source in a diagnosed LaTeX code block, including preamble declarations and source after `\end{document}`. Setting `PreserveUnsupportedAsSource` to `false` omits that source and reports its full span. Comment suppression still applies. The active OfficeIMO profile projects the document body and ignores the inert trailer.

Literal control words follow TeX delimiter rules: `A\textasciitilde B` becomes `A~B`, while `A\textasciitilde{} B` keeps the space after the group. Text, metadata, captions, URLs and image paths share this decoding rule.

```csharp
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;

LatexDocument latex = LatexDocument.Parse(source);
LatexToMarkdownResult converted = latex.ToMarkdownDocumentResult();
string markdown = converted.Value.ToMarkdown();
```

The cancellation-token overload cooperatively stops source rebinding, block and inline projection, and source fallback processing:

```csharp
LatexToMarkdownResult converted = latex.ToMarkdownDocumentResult(null, cancellationToken);
```

Reverse conversion creates canonical bounded-profile LaTeX and reparses the generated source through the lossless engine:

```csharp
MarkdownToLatexResult generated = markdownDocument.ToLatexDocumentResult();
string source = generated.Source;
LatexDocument parsed = generated.Value;
```

The bridge maps front matter, headings, inline formatting and links, lists and definitions, footnotes, images/figures, table captions/labels and common spans, theorem callouts with required declarations, verbatim/code, and math transport. Canonical output escapes TeX arguments and deterministically encodes labels. Unrepresented figure/table container source remains visible with diagnostics. Conversions rebind edited native source before projection. Plain source, the `PreserveOnly` profile, and unsupported table containers receive visible fallbacks and fidelity diagnostics. Literal code and link destinations decode the bridge's escapes. URLs and image paths retain literal tildes; a tilde in ordinary prose maps to spacing. Generated optional titles and terms protect embedded brackets. Required arguments bind brace groups or single character/control-sequence tokens; missing arguments remain visible with diagnostics. Graphics options and counter-based references report their unevaluated layout semantics.

List items, description definitions and quotations retain ordered child blocks, including multiple paragraphs, nested lists and quotations. Footnote references map to typed Markdown references and definitions whose bodies use the same block projection. Explicit TeX marks report a counter simplification; nested footnotes stay visible as diagnosed source. Reader keeps the note's input location and heading path, and PDF includes the projected note body.

Display metadata, figure and table captions, and optional theorem titles decode escaped characters and retain visible text. Nested formatting that a scalar target cannot represent produces a simplification diagnostic. Moving a theorem label onto its callout preserves enclosing supported formatting and math source. Labels inside unsupported command arguments stay in the source fallback. Metadata after `\end{document}` cannot create a title or front matter.

List setup and content without an item remain visible with a source-fallback diagnostic. Commands wrapping a heading, table, list or multiple paragraphs retain their complete source instead of activating only the nested block. Common `l`, `c` and `r` table columns map to Markdown alignments and back; widths, modifiers, repetition and vertical rule styling report their simplifications. Generated TeX retains effective alignment for spanning cells and explicit cell alignment overrides.

Reverse conversion retains formatting and identifiers when promoting a heading to the title, preserves loose-list paragraphs, and writes fragment links as explicit `\hyperref[label]{visible text}` navigation links. Strikethrough uses `ulem` with `normalem` so ordinary emphasis remains italic. Custom list numbering, task markers and highlighted text report their simplifications. Ordinary callouts retain their kind, title and body in a quote. Image alternate text, titles, links and layout hints, and link tooltip/HTML metadata that the bounded target cannot represent report omissions. Recognized verbatim closing-delimiter variants inside code are escaped and diagnosed so literal code remains opaque.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Latex` and `OfficeIMO.Markdown`; the bridge owns bounded-profile mapping, canonical generation, and diagnostics.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Latex.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
