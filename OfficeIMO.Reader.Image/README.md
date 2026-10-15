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

PNG, JPEG, GIF, BMP, TIFF, SVG, EMF, WMF, ICO, PCX, and WebP extensions are registered. Identification is local and header-only. OCR execution remains an explicit host choice through `OfficeIMO.Reader.Ocr` and the shared `OfficeIMO.Ocr.IOcrEngine` contract.

The handler applies a 128 MiB default input ceiling when `ReaderOptions.MaxInputBytes` is not set. Set `IncludePayload = false` together with `CreateOcrCandidate = false` when only metadata is required.

Enable detailed inspection to project the shared Core metadata evidence into located Reader metadata:

```csharp
OfficeDocumentReadResult inspected = reader.ReadDocument("photo.jpg", new ReaderOptions {
    Inspection = new ReaderInspectionOptions()
});
foreach (OfficeDocumentMetadataEntry property in inspected.Metadata) {
    Console.WriteLine($"{property.Category}: {property.Name} = {property.Value}");
}
```

The `image.exif`, `image.exif.timestamp`, `image.exif.gps`, `image.xmp`, `image.icc` and `image.resolution` categories preserve source counts/offsets, raw timestamp components, timezone uncertainty, RDF collection/language qualifiers and separate ICC validity/transform-support facts. Coverage and diagnostics use the existing Reader document contract. `ReaderInspectionOptions` applies operation-wide record and character limits; `ReaderImageOptions.MetadataInspection` controls Core's finite image/profile limits. Its default detailed-metadata input limit is 16 MiB, independently of the handler's general input ceiling. Detailed metadata does not execute OCR or decode image pixels, and explicitly reports unsupported BigTIFF, private MakerNote and additional-IFD evidence.

`ReaderOptions.Projection = ReaderProjection.TextAndMetadata` returns image identification as plain text chunks and metadata without constructing an asset, visual, Markdown or OCR candidate. Enable `Inspection` for the detailed metadata above. The default `Full` projection retains the asset and optional OCR-readiness behavior.
