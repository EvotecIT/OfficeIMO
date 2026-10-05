# OfficeIMO.Reader.Xps

`OfficeIMO.Reader.Xps` registers `.xps` and `.oxps` ingestion through the native
`OfficeIMO.Xps` engine. It reads literal Unicode, native story order, page identity
and tables without rendering or requiring a desktop XPS viewer.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Xps;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddXpsHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("document.oxps");
foreach (OfficeDocumentBlock block in result.EnumerateBlocks())
    Console.WriteLine($"Page {block.Location.Page}: {block.Text}");
```

An already loaded `XpsDocument` supports `ToOfficeDocumentReadResult()` without
serialization or reopening. `XpsDocument.ToOfficeDocumentModel()` owns the reusable
projection, including recursive native structure, list markers and cell spans.
Reader retains bounded text chunks, physical page locations, native links, rectangular tables,
diagnostics and optional SVG preview assets. Spanned table positions are blank in
the rectangular projection; the native structure retains the span values.
Grids exceeding 1000 columns or containing row spans beyond the available rows
are omitted with a diagnostic; their text and recursive native structure remain available.
Page Markdown emits each represented cell once, including nested tables. Rows
outside a truncated grid remain as text. Logical tables that span pages remain
at document scope; each physical page retains its own text instead of receiving
another page's cells.

`ReaderXpsOptions.ReadOptions` controls package expansion, part, page and XML limits.
Registration snapshots these limits, and `ReaderOptions.MaxInputBytes` can tighten
them. Stream input follows Reader's whole-stream snapshot contract, restores a
seekable stream's position and leaves the caller's stream open. Source hashes cover
the captured package bytes. A mutable loaded document has a stable logical source
ID but no inferred package hash.

```csharp
var readerWithPreviews = new OfficeDocumentReaderBuilder()
    .AddXpsHandler(new ReaderXpsOptions { IncludeSvgPreviewAssets = true })
    .Build();
```

Previews use strict native SVG conversion; unsupported paint fails instead of
producing an incomplete preview. Preview projection is bounded to 512 pages and
128 MiB of asset payload. The native package's rendering support matrix applies.

Native story order can cross physical pages. `ReaderLocation.LogicalOrder` retains
that order through canonical block/content traversal and JSON transport. Missing
structure falls back to page/markup order with a diagnostic. Unreferenced Unicode
follows declared stories. Headers and footers retain their own block kinds. Glyph
IDs are not reverse-mapped to Unicode, and text inside brush visuals is excluded.
Overlapping text ownership in native structure is rejected.

Content-based detection recognizes bounded atomic OPC root relationships. Packages
with interleaved root relationships use their `.xps` or `.oxps` extension with the
default detection mode; their contents are assembled by the native package loader.

The adapter references only `OfficeIMO.Reader.Core` and `OfficeIMO.Xps`.
See the [native support matrix](../OfficeIMO.Xps/SUPPORT.md) and
[PDF export API](../OfficeIMO.Xps.Pdf/README.md).
