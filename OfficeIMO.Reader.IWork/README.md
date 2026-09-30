# OfficeIMO.Reader.IWork

`OfficeIMO.Reader.IWork` reads Apple Pages (`.pages`), Numbers (`.numbers`), and Keynote (`.key`) packages through the bounded `OfficeIMO.IWork` semantic model. It emits Reader chunks and a rich result with logical pages, text blocks, tables, links, image assets, and source diagnostics.

## Install

```powershell
dotnet add package OfficeIMO.Reader.IWork
```

## Read a document

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddIWorkHandler()
    .Build();

OfficeDocumentReadResult document = reader.ReadDocument("report.pages");
Console.WriteLine(document.Markdown);
foreach (ReaderTable table in document.Tables) {
    Console.WriteLine($"{table.Title}: {table.Rows.Count} projected rows");
}
```

For a stream, pass its source name: `reader.ReadDocument(stream, "report.numbers")`. The extension identifies the expected iWork kind. `OfficeIMO.Reader.All` also registers this handler. Content detection recognizes direct IWA indexes and inspects up to 16 MiB of a nested `Index.zip` to identify renamed packages. Keep a `.pages`, `.numbers`, or `.key` source name for larger nested indexes.

Pages content is projected as text, drawable tables and images, plus headers and footers. Numbers sheets become logical pages with tables and text boxes in their shared source order. Keynote slides become logical pages with text boxes, tables, images, and presenter notes. Rich text becomes Markdown and linked runs become link entries. Formula cells expose their cached display values in Reader tables; use `OfficeIMO.IWork` directly when formula syntax and typed cells are needed.

`ReaderOptions.MaxTableRows` and `ReaderIWorkOptions.MaximumTableColumns` bound each table, while `ReaderIWorkOptions.MaximumProjectedTableCells` bounds dense table cells across the whole result. `ReaderIWorkOptions.ReadOptions` controls source package and semantic limits. Image bytes are omitted by default; set `IncludeImagePayloads` when the caller needs them. Truncation, omitted tables after the cell budget is exhausted, and unsupported visual details are reported as diagnostics. Reader tables are flat grids; a diagnostic records source header-column and footer-row counts when their roles cannot be represented. Tables with source geometry expose column profiles and geometry diagnostics. If a sparse source grid exceeds 32-bit diagnostic counts, expected and missing counts saturate at `int.MaxValue` and a diagnostic records the exact logical cell count.

Chunks keep Unicode surrogate pairs together. A chunk may exceed `ReaderOptions.MaxChars` by one UTF-16 code unit when that limit would split a character. Rich Markdown may be split independently of plain text; concatenate chunk Markdown in order to recover the complete markup. Markdown preserves bold, italic, strikethrough, and safe links. Diagnostics identify source paragraph and run formatting, explicit page, section, and layout breaks, drawable rotation, and table merges that Reader output cannot represent. Image asset filenames are unique within a result, including when an embedded image is reused.

The Reader path accepts ZIP-form iWork packages as files or streams and directory bundles with a `.pages`, `.numbers`, or `.key` extension. Sync and async path reading treat a bundle as one document. Folder ingestion and `EnumerateDocumentPaths` identify registered bundles without descending into their resources. Normal directories still use folder ingestion; an invalid bundle fails as a document rather than falling back to folder traversal.

Bundle reads retain the iWork owner's entry/path/byte limits, physical-root and regular-file checks, and cancellation. `ReaderOptions.MaxInputBytes` and folder aggregate byte limits also apply. `Source.LengthBytes` counts captured physical file bytes before nested expansion. When `ComputeHashes` is enabled, `Source.SourceHash` hashes the captured normalized entries, including paths, bytes, and nested-index entries; it is a package-content identity rather than a ZIP-file checksum. File and stream sources retain Reader's existing file/payload hash contract. Discovery with an aggregate byte limit parses a bundle through its bounded handler to obtain its source size.

Reader does not paginate Pages layouts or render Keynote slides. Its logical page labels represent a Pages document, Numbers sheet, or Keynote slide, not a rendered Pages page count.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows. License: MIT. Runtime dependencies are `OfficeIMO.Reader.Core` and `OfficeIMO.IWork`.
