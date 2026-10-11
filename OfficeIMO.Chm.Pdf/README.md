# OfficeIMO.Chm.Pdf

Render each selected compiled-help topic through `OfficeIMO.Html.Pdf`, then combine its pages using `OfficeIMO.Pdf`. Independent topic rendering preserves separate CSS contexts and starts each topic at a page boundary.

```csharp
using OfficeIMO.Chm;
using OfficeIMO.Html.Pdf;

ChmDocument book = ChmDocument.Load("manual.chm");
ChmConversionResult<byte[]> result = book.ToPdfBytesResult(
    new ChmConversionOptions { MaxOutputBytes = 32L * 1024 * 1024 },
    new HtmlToPdfOptions());
File.WriteAllBytes("manual.pdf", result.RequireValue());
foreach (var diagnostic in result.Report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

`ToPdfDocumentResult()` and its async counterpart return the existing `PdfDocumentConversionResult`, with CHM source evidence attached. Pass renderer options for page geometry, fonts and the supported PDF policy. All resources resolve inside the archive; conversion never fetches remote resources or opens external files. Encryption, when requested, applies once to the completed PDF.

Multi-topic PDFs are untagged because the shared merger cannot preserve a merged structure tree. Cross-topic links and the keyword index are not rebuilt. These losses are reported. Select one topic to retain the renderer's tagging behavior. This adapter does not claim PDF/UA or browser-equivalent layout. See [CHM support](../OfficeIMO.Chm/SUPPORT.md) and [HTML PDF rendering](../OfficeIMO.Html.Pdf/README.md).

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Chm.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
