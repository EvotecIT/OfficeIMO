# OfficeIMO.DjVu

Read `.djvu` and `.djv` scanned documents, reuse stored text, and render pages with managed OfficeIMO codecs. No external decoder, native library, or executable is required. PDF export and Reader ingestion live in separate adapter packages.

Reference the owning project from a source checkout:

```xml
<ProjectReference Include="../OfficeIMO/OfficeIMO.DjVu/OfficeIMO.DjVu.csproj" />
```

```csharp
using OfficeIMO.DjVu;

var document = DjVuDocument.Load("book.djvu");
var page = document.Pages[0];
var stored = page.GetText();
Console.WriteLine($"{page.Number}: {stored.Status}");
Console.WriteLine(stored.Text);

var rendered = page.Render(new DjVuRenderOptions { Dpi = 150 });
// rendered.Image is an owned OfficeRasterImage; use Core's image encoder to save it.
foreach (var diagnostic in rendered.FidelityDiagnostics)
    Console.WriteLine(diagnostic.Message);
```

`Pages` keeps source order. Page selection in adapters uses one-based source numbers. Text states are `Absent`, `Empty`, `Present`, and `Corrupt`; a malformed layer is reported separately from missing text. Zone byte offsets refer to the original UTF-8 text, while character offsets refer to the returned .NET string. Native rectangles use unrotated pixels with a bottom-left origin; `GetDisplayBounds` maps them to the rotated display view.

Byte inputs and options are copied. `Load(Stream)` reads from its current position, supports non-seekable streams, and leaves the stream open. A file handle is closed before `Load(string)` returns. The source hash and length describe the primary input, excluding resolved indirect components.

Indirect documents require a caller-supplied `DjVuReadOptions.ComponentResolver`. Its component IDs are opaque: OfficeIMO does not interpret them as paths or URLs. Return bytes from an explicitly approved component map. Resolved bytes are copied and count toward the aggregate source limit; recursive includes and foreign component forms are rejected.

Rendering defaults to native DPI, applies display rotation, and returns an RGBA raster. `Region` selects an unrotated native rectangle before scaling and rotation. Read and render options bound input, expansion, page pixels, codec memory, symbols, and output; cancellation is observed during parsing and decoding. `RequireNoLoss()` rejects reported rendering qualifications.

See the [support and independent evidence contract](SUPPORT.md), [Reader adapter](../OfficeIMO.Reader.DjVu/README.md), and [searchable PDF adapter](../OfficeIMO.DjVu.Pdf/README.md).
