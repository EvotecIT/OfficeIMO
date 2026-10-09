# OfficeIMO.Reader.DjVu

Project native DjVu stored text, word geometry, source pages, outlines, and optional page images into the shared Reader contract. The adapter uses `OfficeIMO.DjVu` in process.

```xml
<ProjectReference Include="../OfficeIMO/OfficeIMO.Reader.DjVu/OfficeIMO.Reader.DjVu.csproj" />
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.DjVu;

var reader = new OfficeDocumentReaderBuilder()
    .AddDjVuHandler(new ReaderDjVuOptions {
        ImageMode = ReaderDjVuImageMode.MissingTextPages
    })
    .Build();
var result = reader.ReadDocument("book.djvu");
Console.WriteLine(result.Kind); // DjVu
```

`OfficeIMO.Reader.All` registers this handler with its default text-only policy. `ImageMode.None` avoids decoding images. `MissingTextPages` emits PNG assets for absent or empty text; `AllPages` emits every selected page. Assets default to 150 DPI and have per-page and aggregate limits. `PageNumbers` controls selection and order, preserving each native source page number.

Stored text retains UTF-16 offsets, rotated geometry in PDF points, source identity, and the primary input hash. Its historical OCR engine is unknown, so the adapter does not invent recognition provenance. Absent and empty text produce explicit OCR candidates; corrupt text produces a parsing diagnostic. Executing OCR requires the separate Reader OCR contract and an explicitly supplied engine.

Reader stream ingestion follows the shared Reader convention: seekable input is read from the beginning and its original position is restored; non-seekable input is read from its current position. Streams remain open. The native `DjVuDocument.Load(Stream)` API instead reads from the current position for both stream kinds.

DjVu uses `ReaderInputKind.DjVu` (`27`) and transport schema v11. Older schemas remain available for their supported input kinds. The [native support matrix](../OfficeIMO.DjVu/SUPPORT.md) defines the codec and indirect-document limits.
