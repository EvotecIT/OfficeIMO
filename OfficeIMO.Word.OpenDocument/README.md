# OfficeIMO.Word.OpenDocument

`OfficeIMO.Word.OpenDocument` converts between the OfficeIMO Word and ODT object models. Conversion is explicit and returns a feature-mapping report so callers can inspect approximated, skipped, or unsupported source features.

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

using WordDocument word = WordDocument.Load("input.docx", readOnly: true);
OdfConversionResult<OdtDocument> result = word.ToOpenDocument();
result.Document.Save("output.odt");

foreach (OdfConversionMapping mapping in result.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter currently maps ordered body blocks, headings, paragraphs, basic inline formatting, hyperlinks, lists, tables and merges, embedded inline images, page layout, page breaks, bookmarks, and default headers and footers. The report calls out omitted run, paragraph, table, and image-layout details as well as tracked changes, section-specific layout, alternate headers/footers, footnotes, fields, charts, content controls, and other source features that cannot be represented directly.
