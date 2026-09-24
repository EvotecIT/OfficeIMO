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

The adapter maps ordered body blocks, headings, paragraphs, alignment, indentation, spacing, shading, font family, common run formatting, hyperlinks, lists, tables and merges, embedded inline images, page layout, page breaks, bookmarks, default headers and footers, and footnotes and endnotes. Native ODT notes retain their kind, inline reference order, and plain paragraph text through DOCX conversion. Mixed ODT text, spans, hyperlinks, images, bookmarks, and note references are consumed in document order. Nested inline markup that does not have an exact typed mapping is flattened with an explicit `inline-formatting` approximation instead of being reported as exact.

The report calls out omitted table and image-layout details as well as tracked changes, section-specific layout, alternate headers/footers, fields, charts, content controls, and other source features that cannot be represented directly. Note body formatting is flattened with an explicit approximation; multiple ODT note-body paragraphs become one Word note paragraph, and custom citation text is replaced by Word's automatic numbering. Note bodies outside the supported paragraph subset are reported as unsupported. Use `ToOpenDocumentResult` or `ToWordDocumentResult` for evidence-bearing conversion, and set the options' `LossPolicy` to `ThrowOnAnyLoss` for strict workflows.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Word` and `OfficeIMO.OpenDocument`; the adapter only owns feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
