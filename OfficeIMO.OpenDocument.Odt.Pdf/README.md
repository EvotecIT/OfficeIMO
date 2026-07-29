# OfficeIMO.OpenDocument.Odt.Pdf

Bidirectional ODT and PDF conversion without Excel or PowerPoint dependencies.

```csharp
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Open("proposal.pdf");
PdfOdtConversionResult result = pdf.ToOdtDocumentResult();
result.Value.Save("proposal.odt");
```

PDF to ODT is semantic reconstruction through the PDF-to-Word and Word-to-ODT adapters. Inspect both stage reports when fidelity matters.
