# OfficeIMO.Reader.PowerPoint

PowerPoint presentation support for `OfficeIMO.Reader.Core`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.PowerPoint;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPowerPointHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("slides.pptx");
```

The package supports PPTX/PPTM plus OfficeIMO.PowerPoint's legacy PPT/POT/PPS import path. It depends only on Reader.Core and OfficeIMO.PowerPoint.
