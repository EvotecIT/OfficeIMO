# OfficeIMO.Publisher.Pdf

Convert recovered Publisher pages to PDF through the shared OfficeIMO drawing
and PDF engines. The adapter references `OfficeIMO.Publisher` and `OfficeIMO.Pdf`;
it does not decode `.pub` records itself or require Publisher to be installed.

```csharp
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

var publication = PublisherDocument.Load("brochure.pub");
OfficeConversionReport report = publication.SaveAsPdf("brochure.pdf");

foreach (var diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.LossKind}: {diagnostic.Message}");
Console.WriteLine($"Fidelity loss: {report.HasLoss}");
```

Document pages retain publication order and physical dimensions with zero added
margins. Master definitions are already applied by the Publisher reader and do
not become separate PDF pages. Text and simple shapes remain drawing elements;
pictures use the PDF engine's managed image handling.

The report combines the `PublisherReadReport` diagnostics with PDF rendering
and serialization findings. Inspect it before accepting a conversion; successful
saving does not imply native fidelity. Set `FailOnLoss` on
`PublisherPdfExportOptions` to make reported loss an `OfficeConversionException`
before anything is written.

`ToPdf(options, out report)` returns the PDF model, `ToPdfAsync` returns it with
the report, and `SaveAsPdf`/`SaveAsPdfAsync` write a path or a caller-owned
stream, which stays open. All routes accept cancellation; the options and their
`PdfCreateOptions` are cloned for the operation. The async forms run the
CPU-bound conversion on the thread pool and write the output asynchronously. Refer to the [Publisher support contract](../OfficeIMO.Publisher/SUPPORT.md)
for profile and reconstruction limits, including linked text flow and native
metafile placeholders.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Publisher.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
