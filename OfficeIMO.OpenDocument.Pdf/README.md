# OfficeIMO.OpenDocument.Pdf

`OfficeIMO.OpenDocument.Pdf` provides direct, loss-aware PDF conversion for OpenDocument text documents, spreadsheets, and presentations.

The adapter does not introduce another renderer. It projects each OpenDocument model through its existing Office semantic adapter and then through the corresponding first-party PDF engine. The returned `PdfDocumentConversionResult` combines diagnostics from both stages, so callers can inspect approximations or omissions before saving.

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Pdf;
using OfficeIMO.Pdf;

OdtDocument document = OdtDocument.Load("proposal.odt");
PdfDocumentConversionResult result = document.ToPdfDocumentResult();

result.Report.RequireNoLoss();
result.Save("proposal.pdf");
```

The same façade is available on `OdsDocument` and `OdpPresentation`, including byte-array, path, stream, and asynchronous save entry points.
