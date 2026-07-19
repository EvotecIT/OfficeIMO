# OfficeIMO.Reader package family

OfficeIMO.Reader is the read-only extraction API shared by selective format adapters. `OfficeIMO.Reader.Core` owns
contracts, routing, limits, normalized results, processors, and custom-handler registration. It deliberately contains
no Word, Excel, PowerPoint, Email, PDF, image, or other format engine.

## Choose packages by data type

Install the adapters an application actually uses. Every adapter brings Core and its owning format package:

| Need | Package |
| --- | --- |
| Contracts, plain text, custom handlers | `OfficeIMO.Reader.Core` |
| Word | `OfficeIMO.Reader.Word` |
| Excel | `OfficeIMO.Reader.Excel` |
| PowerPoint | `OfficeIMO.Reader.PowerPoint` |
| Markdown | `OfficeIMO.Reader.Markdown` |
| Email, Outlook stores, and OAB | `OfficeIMO.Reader.Email` |
| PDF | `OfficeIMO.Reader.Pdf` |
| Every local managed adapter | `OfficeIMO.Reader.All` |

Other `OfficeIMO.Reader.*` packages follow the same rule. The former `OfficeIMO.Reader` convenience package is
retired; there is no compatibility package with that identity. Public namespaces remain `OfficeIMO.Reader`.

## Build a selective reader

```powershell
dotnet add package OfficeIMO.Reader.Word
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .WithMaxConcurrentReads(4)
    .Build();

foreach (ReaderChunk chunk in reader.Read("Policy.docx", new ReaderOptions {
    MaxChars = 8_000,
    MaxTableRows = 200,
    ComputeHashes = true
})) {
    Console.WriteLine(chunk.Location.HeadingPath ?? chunk.Location.Path);
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

`OfficeDocumentReader.Default` is an empty immutable Core reader. A format becomes available only after its adapter
package is installed and registered on a builder. This keeps reader instances isolated and prevents a supposedly
small Core install from hiding a complete document-engine graph.

## Build a broad reader intentionally

```powershell
dotnet add package OfficeIMO.Reader.All
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddAllOfficeIMOHandlers()
    .WithMaxConcurrentReads(4)
    .Build();

foreach (ReaderSourceDocument source in reader.ReadFolderDocuments(
    "KnowledgeBase",
    new ReaderFolderOptions {
        Recurse = true,
        DeterministicOrder = true,
        MaxFiles = 10_000,
        MaxTotalBytes = 500L * 1024 * 1024
    },
    new ReaderOptions { MaxChars = 4_000, ComputeHashes = true })) {
    Console.WriteLine($"{source.Path}: {source.ChunksProduced} chunks");
}
```

All composes every in-repository local managed adapter. OCR engines, external processes, network clients, hosted
providers, and native tools remain explicit host choices and are not pulled into All.

## Result contracts

- `Read(...)` returns deterministic `ReaderChunk` sequences for indexing and RAG.
- `ReadDocument(...)` returns the stable rich result with pages, blocks, tables, links, forms, assets, visuals, OCR
  candidates, metadata, and structured diagnostics.
- Async path, stream, byte, folder, and bounded batch APIs preserve cancellation, ordering, input limits, and
  caller-owned stream lifetime.
- `ReadStructured(...)` and `ReadHierarchical(...)` provide bounded structured records and hierarchy-aware RAG leaves.
- Ordered processors can normalize or filter every rich result in one immutable reader instance.

The detailed API guide is [OfficeIMO.Reader.Core/README.md](../OfficeIMO.Reader.Core/README.md). Package ownership,
migration, dependency, and release gates are recorded in
[officeimo.reader.modular-roadmap.md](officeimo.reader.modular-roadmap.md).
