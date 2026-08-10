# OfficeIMO.PowerPoint.OpenDocument

Explicit conversion between `OfficeIMO.PowerPoint` presentations and native `OfficeIMO.OpenDocument` presentations.

```csharp
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.OpenDocument;

using PowerPointPresentation presentation = PowerPointPresentation.Load("deck.pptx");
OdfConversionResult<OdpPresentation> conversion = presentation.ToOpenDocumentResult();
conversion.Value.Save("deck.odp");

foreach (var mapping in conversion.Report.Mappings) {
    Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count})");
}
```

The adapter maps slide size and order, hidden slides, text boxes, ordered mixed text/run/hyperlink content, common run formatting, images, tables and merges, basic shapes, solid backgrounds, common transitions, and plain speaker notes. Nested inline markup without an exact typed mapping is flattened with an explicit approximation. Masters, complex geometry, charts, SmartArt, media, advanced animations, unsupported transition families, and other detected features are called out in the conversion report. Set the conversion options' `LossPolicy` to `ThrowOnAnyLoss` when approximated, skipped, or unsupported content must reject the conversion.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.PowerPoint` and `OfficeIMO.OpenDocument`; the adapter owns feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
