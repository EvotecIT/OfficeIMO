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

The adapter maps ordered body blocks, headings, paragraphs, alignment, indentation, spacing, shading, font family, common run formatting, hyperlinks, lists, tables and merges, embedded inline images, page layout, page breaks, bookmarks, default and first-page headers and footers, even-page Word headers and footers to ODT left-page variants, and body footnotes and endnotes. The header and footer mapping uses the first Word section and the first ODT master page. When ODT defines only one alternate story in a first-page or left-page pair, the missing story uses its default header or footer. Body notes retain their kind, inline reference order, and plain paragraph text through DOCX conversion. Bare Word `PAGE`, `NUMPAGES`, `DATE`, and `TIME` simple fields map to native ODT fields with cached display text, and the same basic ODT fields map back to Word. Mixed ODT text, spans, hyperlinks, images, fields, bookmarks, and note references are consumed in document order. Nested inline markup without an exact typed mapping is flattened with an explicit `inline-formatting` approximation.

The report calls out omitted table and image-layout details as well as tracked changes, section-specific layout, later-section or inactive alternate headers and footers, page-number restarts, header/footer tables and other nonparagraph blocks, unsupported field instructions or properties, charts, content controls, and other source features that cannot be represented directly. ODT header and footer content marked `style:display="false"` remains hidden in Word and is reported as skipped content. First-page header/footer stories require ODF 1.4 for both package and flat XML output. Unsupported fields retain cached display text where available and report `fields` loss; lost direct formatting on Word field results is reported separately. Note body formatting and named paragraph styles are flattened with an explicit approximation; multiple ODT note-body paragraphs become one Word note paragraph. Custom Word marks and ODT citation labels are replaced by automatic numbering. Direct Word anchor formatting and styled, linked, or paragraph-formatted ODT references are reported as approximated when replaced by the destination's default reference. Authored Word note numbering, placement, and customized separators are reported as approximated because this conversion does not carry them into ODT. Repeated Word references to one note, ODT note configuration, notes in headers or footers, nested image or other nontext note body content, and block content inside Word notes are reported as unsupported. Use `ToOpenDocumentResult` or `ToWordDocumentResult` for evidence-bearing conversion, and set the options' `LossPolicy` to `ThrowOnAnyLoss` for strict workflows.

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
