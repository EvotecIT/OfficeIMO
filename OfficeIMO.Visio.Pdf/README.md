# OfficeIMO.Visio.Pdf

`OfficeIMO.Visio.Pdf` adds document-shaped PDF entry points to
`OfficeIMO.Visio`. It reuses the shared Visio reader projection and
`OfficeIMO.Reader.Pdf` engine, including explicit semantic-fallback diagnostics.

```powershell
dotnet add package OfficeIMO.Visio.Pdf
```

```csharp
using OfficeIMO.Visio;
using OfficeIMO.Visio.Pdf;

VisioDocument diagram = VisioDocument.Load("architecture.vsdx");
diagram.SaveAsPdf("architecture.pdf");

byte[] pdf = diagram.ToPdf();
```

The conversion produces searchable diagram text and topology. It reports when
the shared projection uses semantic fallback instead of claiming native Visio
page-rendering fidelity.
