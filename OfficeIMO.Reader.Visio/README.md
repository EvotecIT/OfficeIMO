# OfficeIMO.Reader.Visio - Visio reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Visio)](https://www.nuget.org/packages/OfficeIMO.Reader.Visio)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Visio?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Visio)

`OfficeIMO.Reader.Visio` provides a Visio adapter for `OfficeIMO.Reader.Core` using `OfficeIMO.Visio` inspection snapshots.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Visio
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddVisioHandler()
    .Build();

IReadOnlyList<ReaderChunk> chunks = reader
    .Read("diagram.vsdx")
    .ToList();
```

## Examples

### Inspect pages, shapes, and Shape Data

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Visio;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddVisioHandler()
    .Build();

foreach (var chunk in reader.Read("architecture.vsdx")) {
    Console.WriteLine($"Page {chunk.Location.Page}: {chunk.Id}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);

    foreach (var table in chunk.Tables ?? Array.Empty<ReaderTable>()) {
        Console.WriteLine($"Shape Data table: {table.Rows.Count} row(s)");
    }
}
```

### Read Visio files from a folder with other formats

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Visio;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddPdfHandler()
    .AddVisioHandler()
    .Build();

var chunks = reader.ReadFolder("Architecture",
    new ReaderFolderOptions {
        Extensions = new[] { ".vsdx", ".pdf" },
        Recurse = true,
        DeterministicOrder = true
    },
    new ReaderOptions {
        MaxChars = 8_000
    }).ToList();
```

### Read topology and geometry

```csharp
OfficeDocumentReadResult document =
    VisioReaderAdapter.ReadDocument("network.vsdx");

foreach (ReaderVisual topology in document.Visuals) {
    Console.WriteLine($"Page {topology.Location.Page}: {topology.Content}");
}

foreach (OfficeDocumentPage page in document.Pages) {
    Console.WriteLine($"{page.Name}: {page.Width} x {page.Height} points");
}
```

The topology visual is deterministic JSON describing page shapes and connector edges. Shared regions and page dimensions are expressed in points. A configured reader uses the same native mapping through `reader.ReadDocument("network.vsdx")`.

## What it emits

- Page-aware chunks for `.vsdx`, `.vsdm`, `.vstx`, and `.vstm` files.
- Shape Data as `ReaderTable` rows.
- Pages, shapes, connectors, hyperlinks, and optional preview asset metadata in the shared read result envelope.
- Point-based geometry and one topology `ReaderVisual` per page for graph-aware consumers.

## Boundaries

- Reader adapter configuration belongs here.
- Visio package parsing and inspection belongs in `OfficeIMO.Visio`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** `System.IO.Packaging` only through `OfficeIMO.Visio`.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Visio` own diagram inspection, topology, chunks, tables, and visuals.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
