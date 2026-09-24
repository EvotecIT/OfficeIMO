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

The adapter maps slide size and order, hidden slides, text boxes, ordered mixed text/run/hyperlink content, common run formatting, images, tables and merges, basic shapes, solid backgrounds, common transitions, and plain speaker notes. PPTX-to-ODP conversion also retains each slide's master and layout association, direct RGB master backgrounds without effects, and common title, subtitle, body, and object placeholder roles. ODP-to-PPTX conversion retains those placeholder roles and flattens inherited solid master backgrounds onto the affected slides. An unsupported slide background override does not expose the mapped master color as the slide's own background. Nested inline markup without an exact typed mapping is flattened with an explicit approximation. Master and layout drawing content, placeholder geometry and indexes, theme or effect shape appearance, table and cell styles, complex shapes, charts, SmartArt, media, advanced animations, unsupported transition families, and other detected features remain in the conversion report as loss. Set the conversion options' `LossPolicy` to `ThrowOnAnyLoss` when approximated, skipped, or unsupported content must reject the conversion.

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.PowerPoint` and `OfficeIMO.OpenDocument`; the adapter owns feature mapping and fidelity reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.PowerPoint.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
