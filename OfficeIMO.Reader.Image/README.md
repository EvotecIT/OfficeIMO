# OfficeIMO.Reader.Image

`OfficeIMO.Reader.Image` adds standalone image files to an isolated `OfficeDocumentReader`. It uses the existing header-only `OfficeIMO.Drawing` identification API to expose format, dimensions, DPI, a materializable source asset, and optional OCR readiness without decoding pixels or running OCR.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Image;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddImageHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("diagram.png");
Console.WriteLine(result.Markdown);
Console.WriteLine(result.Assets[0].MediaType);
```

PNG, JPEG, GIF, BMP, TIFF, SVG, EMF, WMF, ICO, PCX, and WebP extensions are registered. Identification is local and header-only. OCR execution remains an explicit host choice through the core `IOfficeOcrEngine` contract.

The handler applies a 128 MiB default input ceiling when `ReaderOptions.MaxInputBytes` is not set. Set `IncludePayload = false` together with `CreateOcrCandidate = false` when only metadata is required.
