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

For a stream, pass its source name: `reader.ReadDocument(stream, "report.numbers")`. The extension identifies the expected iWork kind. `OfficeIMO.Reader.All` also registers this handler.

Pages content is projected as text, drawable tables and images, plus headers and footers. Numbers sheets become logical pages with tables and text boxes. Keynote slides become logical pages with text boxes, tables, images, and presenter notes. Rich text becomes Markdown and linked runs become link entries. Formula cells expose their cached display values in Reader tables; use `OfficeIMO.IWork` directly when formula syntax and typed cells are needed.

`ReaderOptions.MaxTableRows` and `ReaderIWorkOptions.MaximumTableColumns` bound dense table materialization. `ReaderIWorkOptions.ReadOptions` controls source package and semantic limits. Image bytes are omitted by default; set `IncludeImagePayloads` when the caller needs them. Truncation and unsupported visual details are reported as diagnostics.

Chunks keep Unicode surrogate pairs together. A chunk may exceed `ReaderOptions.MaxChars` by one UTF-16 code unit when that limit would split a character. Markdown preserves bold, italic, strikethrough, and safe links; diagnostics identify source paragraph and run formatting that Markdown cannot represent. Image asset filenames are unique within a result, including when an embedded image is reused.

The Reader path accepts ZIP-form iWork packages as files or streams. It does not paginate Pages layouts or render Keynote slides. Its logical page labels represent a Pages document, Numbers sheet, or Keynote slide, not a rendered Pages page count.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows. License: MIT. Runtime dependencies are `OfficeIMO.Reader.Core` and `OfficeIMO.IWork`.
