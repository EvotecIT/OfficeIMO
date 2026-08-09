# OfficeIMO.Word.OpenDocument

`OfficeIMO.Word.OpenDocument` converts between the OfficeIMO Word and ODT object models. Conversion is explicit and returns a feature-mapping report so callers can inspect approximated, skipped, or unsupported source features.

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using OfficeIMO.OpenDocument;

using WordDocument word = WordDocument.Load("input.docx", readOnly: true);
OdfConversionResult<OdtDocument> result = word.ToOpenDocumentResult();
result.Value.Save("output.odt");

foreach (OdfConversionMapping mapping in result.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter maps ordered body blocks, headings, paragraphs, alignment, indentation, spacing, shading, font family, common run formatting, hyperlinks, lists, tables and merges, embedded inline images, page layout, page breaks, bookmarks, and default headers and footers. Mixed ODT text, spans, hyperlinks, images, and bookmark markers are consumed in document order. Nested inline markup that does not have an exact typed mapping is flattened with an explicit `inline-formatting` approximation instead of being reported as exact.

The report calls out omitted table and image-layout details as well as tracked changes, section-specific layout, alternate headers/footers, footnotes, fields, charts, content controls, and other source features that cannot be represented directly. Use `ToOpenDocumentResult` or `ToWordDocumentResult` for evidence-bearing conversion, and set the options' `LossPolicy` to `ThrowOnAnyLoss` for strict workflows.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Word` and `OfficeIMO.OpenDocument`; the adapter only owns feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
