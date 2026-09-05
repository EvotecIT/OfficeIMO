# OfficeIMO.Visio.Pdf

`OfficeIMO.Visio.Pdf` adds document-shaped PDF entry points to
`OfficeIMO.Visio`. Visio projects directly into the dependency-free
`OfficeDocumentModel` from `OfficeIMO.Core`, and `OfficeIMO.Pdf` owns the PDF
projection policy and composition. Reader packages are not part of this
conversion boundary.

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

VisioDocument diagram = VisioDocument.Load("architecture.vsdx");
diagram.SaveAsPdf("architecture.pdf");

byte[] pdf = diagram.ToPdfBytes();
```

The conversion produces searchable diagram text and topology. It reports when
the shared projection uses semantic fallback instead of claiming native Visio
page-rendering fidelity.

## Dependency footprint

- `OfficeIMO.Core` owns the neutral document model.
- `OfficeIMO.Visio` owns VSDX inspection and projection into that model.
- `OfficeIMO.Pdf` owns loss-aware PDF composition.
- `OfficeIMO.Reader.Visio` and `OfficeIMO.Reader.Pdf` are not dependencies.
