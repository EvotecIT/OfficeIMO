# OfficeIMO.Latex.Markdown

This package maps the bounded `OfficeIMO.Latex` profile to and from `OfficeIMO.Markdown`. It reports source fallbacks and simplifications, especially for TeX math layout, package-specific commands, bibliography formatting, and unknown environments.

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

The bridge maps front matter, headings, inline formatting and links, lists and definitions, images/figures, table captions/labels and common spans, theorem callouts with required declarations, verbatim/code, and math transport. Canonical output escapes TeX arguments and deterministically encodes labels. Unrepresented figure/table container source remains visible with diagnostics. Conversions rebind edited native source before projection. Plain source, the `PreserveOnly` profile, and unsupported table containers receive visible fallbacks and fidelity diagnostics. Literal code and link destinations decode the bridge's escapes. URLs and image paths retain literal tildes; a tilde in ordinary prose maps to spacing. Generated optional titles and terms protect embedded brackets. Required arguments must be braced in the bounded profile; missing or unbraced arguments remain visible with diagnostics. Graphics options and counter-based references report their unevaluated layout semantics.

Display metadata, figure and table captions, and optional theorem titles decode escaped characters and retain visible text. Nested formatting that a scalar target cannot represent produces a simplification diagnostic. Moving a theorem label onto its callout preserves enclosing inline formatting and math source.

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
