# OfficeIMO.DjVu.Pdf

Convert scanned DjVu pages into PDFs with source physical size, display rotation, a searchable stored-text layer, and supported outline navigation. This adapter uses `OfficeIMO.DjVu` and `OfficeIMO.Pdf` in process.

```xml
<ProjectReference Include="../OfficeIMO/OfficeIMO.DjVu.Pdf/OfficeIMO.DjVu.Pdf.csproj" />
```

```csharp
using OfficeIMO.DjVu;
using OfficeIMO.DjVu.Pdf;

var document = DjVuDocument.Load("book.djvu");
var result = document.ToPdfDocumentResult(new DjVuToPdfOptions {
    PageNumbers = new[] { 2, 1 }
});
var saved = result.Save("selected.pdf");
```

`ToPdfBytes`, `SaveAsPdf`, and asynchronous equivalents share the same conversion. Native DPI is the default. Lower render DPI is an explicit reported approximation; physical page size still comes from the source. The PDF retains scanned pixels rather than reconstructing editable paragraphs or source typography.

The result carries a `DjVuPdfConversionReport` alongside PDF-stage reports. Per-page evidence records source/output page numbers, original text status, `StoredText`/`Ocr`/`None` provenance, pixel size, physical size, and embedded character count. Conversion reports distinguish omitted annotations, corrupt text, unsupported navigation targets, approximate geometry, and render qualifications. Use the existing result/report `RequireNoLoss()` gate when reported loss is unacceptable.

For OCR, supply an existing `IOcrEngine` through `DjVuToPdfOptions.OcrEngine` and call an asynchronous entrypoint. Only selected pages with `Absent` or `Empty` stored text are recognized. Present text is reused and corrupt text remains a diagnostic. OCR uses the existing Reader execution limits and records new provider/model/confidence evidence; text without usable provider geometry is reported and not placed into an invented searchable rectangle. OfficeIMO does not install an OCR engine.

Options bound selected pages, decoded image pixels and working memory, individual and aggregate PNG bytes, searchable spans and characters, and serialized PDF bytes. Stream output stays open; seekable output is replaced and rewound. Serialization failure can leave partial output. The shared `djvu-pdf` workflow stages, reopens, and publishes output atomically, and retains combined source/PDF diagnostics. Workflow OCR remains a separate explicit operation.

See the [native support matrix](../OfficeIMO.DjVu/SUPPORT.md) for exact codecs and independent qualification.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.DjVu.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
