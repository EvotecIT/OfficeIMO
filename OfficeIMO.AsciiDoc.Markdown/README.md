# OfficeIMO.AsciiDoc.Markdown

`OfficeIMO.AsciiDoc.Markdown` is the thin, loss-aware conversion bridge between the native `OfficeIMO.AsciiDoc` model and `MarkdownDoc`.

```csharp
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;

AsciiDocDocument source = AsciiDocDocument.Parse(asciiDoc).Document;
AsciiDocMarkdownConversionResult result = source.ToMarkdownDocument();

string markdown = result.Document.ToMarkdown();
```

Reverse conversion generates canonical AsciiDoc, selects longer delimited-block fences when content contains the normal fence, and reparses it through the lossless native engine:

```csharp
MarkdownAsciiDocConversionResult generated = markdownDocument.ToAsciiDocDocument();
string asciiDoc = generated.Source;
```

The bridge maps typed inline content, metadata, lists and compound children, definitions, admonitions, structured tables and spans, images, code metadata, anchors, and STEM where the target model can carry them. Constructs without a safe equivalent are preserved visibly or omitted according to options, with source-located diagnostics.

The adapter never participates in native AsciiDoc parsing or round-trip writing.
