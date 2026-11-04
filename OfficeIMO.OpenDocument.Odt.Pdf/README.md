# OfficeIMO.OpenDocument.Odt.Pdf

Bidirectional ODT and PDF conversion without Excel or PowerPoint dependencies.

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;

OdtDocument source = OdtDocument.Load("proposal.odt");
OfficeConversionReport pdfReport = source.SaveAsPdf("proposal.pdf", new OdtPdfExportOptions { FailOnLoss = true });

PdfDocument pdf = PdfDocument.Load("proposal.pdf");
OfficeConversionReport odtReport = pdf.SaveAsOdt("reconstructed.odt");
```

Forward conversion takes `OdtPdfExportOptions` (`ConversionOptions` for the OpenDocument import, `PdfOptions` for the PDF engine, `FailOnLoss`); `ToPdf(options, out report)`, `ToPdfAsync`, `SaveAsPdf` and `SaveAsPdfAsync` return the report of both stages. Reverse conversion takes `PdfOdtExportOptions` (`PdfOptions`, `OpenDocumentOptions`, `FailOnLoss`) on a `PdfDocument` or `PdfDocumentReadResult`: `ToOdt(options, out report)`, `ToOdtAsync`, `SaveAsOdt` and `SaveAsOdtAsync`. The reverse report is a `PdfOdtConversionReport` with the PDF and OpenDocument stage reports. PDF to ODT is semantic reconstruction through the PDF-to-Word and Word-to-ODT adapters. Inspect both stage reports when fidelity matters.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument.Odt.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
