# OfficeIMO.Xps

Read, create, edit, and save XPS and OpenXPS documents without Windows desktop APIs
or an external conversion process. The package references only `OfficeIMO.Core`.
`OfficeIMO.Xps.Pdf` provides the optional bridge to the existing PDF engine.

## Read and convert

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Load("report.oxps");
XpsPage page = document.Pages[0];
string text = page.ExtractText();
File.WriteAllText("page-1.svg", page.ToSvg().Svg);
File.WriteAllBytes("page-1.png", page.ExportImage(OfficeImageExportFormat.Png).Bytes);
```

Coordinates are XPS units: 96 units per inch. Pages follow the native document
sequence, including multiple fixed documents and repeated page references.
Repeated references share a single editable page: editing that native part updates
all occurrences.
`ExtractText()` returns `UnicodeString` runs in markup order; it does not infer
paragraphs or recover missing Unicode from glyph IDs.

SVG export outlines embedded TrueType glyphs at their native positions, including
explicit glyph IDs, advances, offsets, and clusters. It does not reshape printed
text or substitute installed fonts. SVG navigation uses `page-N.svg` filenames
and `xps-` prefixes for native named targets; keep that naming when exporting a
whole document. The original native navigation remains in the saved XPS package.

Conversion is strict by default. A feature outside the supported rendering
profile throws rather than silently disappearing. To inspect a partial SVG and
its losses explicitly:

```csharp
XpsSvgResult result = page.ToSvg(allowPartial: true);
foreach (string diagnostic in result.Diagnostics)
    Console.WriteLine(diagnostic);
```

Drawing, image, and PDF conversion also enforce the shared engines' resource
limits and reject reported SVG import losses. See the [support matrix](SUPPORT.md)
for the precise rendering and preservation boundaries.

## Create and edit

```csharp
using OfficeIMO.Xps;

XpsDocument document = XpsDocument.Create(XpsFormat.OpenXps);
string font = document.AddFont(File.ReadAllBytes("licensed-font.ttf"));
XpsPage page = document.AddPage(816, 1056);
page.AddPath("M48,48 H768 V1008 H48 Z", "#FFF2F5FA");
page.AddText("Quarterly report", font, 28, 72, 112);
string image = document.AddResource("Resources/chart.png",
    File.ReadAllBytes("chart.png"), "image/png");
page.AddImage(image, 72, 160, 400, 240);
document.Save("report.oxps");
```

Use `XpsFormat.Xps` for Microsoft's original dialect. File extensions do not
change dialects. Fonts are obfuscated by default; callers must have permission
to embed them. PNG/JPEG placement uses the image's declared resolution.

`GetMarkup()` returns a detached `XElement`. Edit it and call `ReplaceMarkup()`
to change an existing page, including native features outside the rendering
profile. This retains the format's XML rather than reconstructing it from a
rendered approximation. `AddPage()` appends to newly created documents; loaded
sequences retain their existing structure. `AddResource()` adds a new part and
`GetPartBytes()` returns a copy; neither exposes mutable internal buffers.

## Native brush rendering

Native page markup can use linear and radial gradients, PNG/JPEG image brushes,
and visual brushes containing paths, glyphs, or canvases. Image and visual brushes
map absolute viewboxes into viewports and support tiled and mirrored repetition.
Canvas, path, and glyph opacity masks use brush alpha. These features flow through
`ToSvg()`, `ToDrawing()`, image export, and the PDF bridge within the
[documented rendering limits](SUPPORT.md). Conversion rejects unsupported drawing
features instead of silently dropping them.

## I/O and preservation

`Load` accepts paths, byte arrays, and readable streams, including non-seekable
streams. It reads from the current stream position and leaves caller streams
open. `XpsReadOptions` bounds compressed input, expanded bytes, part count, page
count, and XML depth. Cancellation is accepted by loading and conversion APIs.
External resource URIs and XML DTDs are rejected; resources are never fetched.

Saving keeps the original dialect, opaque package parts, and native page markup.
It rebuilds ZIP packaging and content types; it does not preserve original ZIP
metadata or byte-identical XML serialization. Repeated saves of the same document
are deterministic on the same runtime. File saves stage and atomically replace
the destination. Stream saves leave the stream open and can leave partial output
if the destination fails or the operation is cancelled.

Digitally signed packages may be inspected, but saving them is rejected to avoid
presenting changed content with stale signatures. This package does not validate
signatures or decrypt protected packages.

Targets: .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2 on Windows.
