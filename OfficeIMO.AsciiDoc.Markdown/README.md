# OfficeIMO.AsciiDoc.Markdown

`OfficeIMO.AsciiDoc.Markdown` is the thin, loss-aware conversion bridge between the native `OfficeIMO.AsciiDoc` model and `MarkdownDoc`.

```csharp
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;

AsciiDocDocument source = AsciiDocDocument.Parse(asciiDoc);
AsciiDocToMarkdownResult result = source.ToMarkdownDocumentResult();

string markdown = result.Value.ToMarkdown();
bool hasLoss = result.Report.HasLoss;
```

Reverse conversion generates canonical AsciiDoc, selects longer delimited-block fences when content contains the normal fence, and reparses it through the lossless native engine:

```csharp
MarkdownToAsciiDocResult generated = markdownDocument.ToAsciiDocDocumentResult();
string asciiDoc = generated.Source;
```

The bridge maps typed inline content, metadata, lists and compound children, definitions, admonitions, structured tables and spans, images, code metadata, anchors, and STEM where the target model can carry them. Constructs without a safe equivalent are preserved visibly or omitted according to options, with source-located diagnostics.

The adapter never participates in native AsciiDoc parsing or round-trip writing.

## References and compound content

Forward conversion uses the attributes in effect at each source block, including set/unset changes inside compound blocks. Example, sidebar, open, quote, and admonition containers project all their child blocks. AsciiDoc-style table cells expose their parsed block bodies; ordinary cells expose editable inline content.

Labeled URLs become links. Named and anonymous footnotes become Markdown references with one definition per used note. Explicit anchors, bibliography anchors, and generated section IDs form a document reference catalog; a reference without a label uses the catalog's reference text when available. Heading IDs reach the Markdown model and HTML output. Sections with `sectids` disabled receive no automatic HTML ID.

For a single block, supply the catalog of its containing document so definitions elsewhere remain available:

```csharp
AsciiDocReferenceCatalog references = AsciiDocReferenceCatalog.Create(source);
AsciiDocBlockContext context = source.GetBlockContexts().First();
AsciiDocToMarkdownResult chunk = context.Block.ToMarkdownDocumentResult(
    context.Attributes,
    new AsciiDocToMarkdownOptions { References = references });
```

Create a new catalog after editing titles, anchors, attributes, footnotes, or document order. Conversion emits only the footnote definitions used by the selected block. Callout lists retain their numbers visibly when Markdown cannot express the numbering directly; inspect `Report.Diagnostics` for those presentation approximations and other lossy mappings.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.AsciiDoc` and `OfficeIMO.Markdown`; the bridge owns typed mapping, canonical generation, and source-located diagnostics.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.AsciiDoc.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
