# OfficeIMO.OpenDocument.Odt.Pdf

Bidirectional ODT and PDF conversion without Excel or PowerPoint dependencies.

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;

OdtDocument source = OdtDocument.Load("proposal.odt");
PdfDocumentConversionResult pdfResult = source.ToPdfDocumentResult();
pdfResult.RequireNoLoss();
pdfResult.Save("proposal.pdf");

PdfDocument pdf = PdfDocument.Load("proposal.pdf");
PdfOdtConversionResult odtResult = pdf.ToOdtDocumentResult();
odtResult.Value.Save("reconstructed.odt");
```

Forward results expose the typed ODT-to-Word report in `SourceConversionReports` and PDF-layout diagnostics in `Report`; `ConversionReports`, `HasLoss`, and `RequireNoLoss()` cover both stages. PDF to ODT is semantic reconstruction through the PDF-to-Word and Word-to-ODT adapters. Inspect both stage reports when fidelity matters.
