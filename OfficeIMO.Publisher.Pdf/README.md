# OfficeIMO.Publisher.Pdf

Convert recovered Publisher pages to PDF through the shared OfficeIMO drawing
and PDF engines. The adapter references `OfficeIMO.Publisher` and `OfficeIMO.Pdf`;
it does not decode `.pub` records itself or require Publisher to be installed.

```csharp
using OfficeIMO.Publisher;
using OfficeIMO.Publisher.Pdf;

var publication = PublisherDocument.Load("brochure.pub");
var converted = publication.ToPdfDocumentResult();

foreach (var diagnostic in publication.ReadReport.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.LossKind}: {diagnostic.Message}");

var saved = converted.Save("brochure.pdf");
Console.WriteLine($"Saved: {saved.Succeeded}, fidelity loss: {saved.HasLoss}");
```

Document pages retain publication order and physical dimensions with zero added
margins. Master definitions are already applied by the Publisher reader and do
not become separate PDF pages. Text and simple shapes remain drawing elements;
pictures use the PDF engine's managed image handling.

`ToPdfDocumentResult()` returns the existing `PdfDocumentConversionResult` with
`PublisherReadReport` in `SourceConversionReports`. The PDF report captures
rendering and serialization findings. Inspect source and output reports before
accepting a conversion; successful saving does not imply native fidelity.

`ToPdfDocument()`, `ToPdfBytes()` and `SaveAsPdf()` provide convenience routes.
Path and caller-stream saves return `PdfSaveResult`; stream output remains open.
All routes accept cancellation and optional `PdfOptions`, which are cloned for
the operation. Refer to the [Publisher support contract](../OfficeIMO.Publisher/SUPPORT.md)
for profile and reconstruction limits, including linked text flow and native
metafile placeholders.
