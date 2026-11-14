---
title: Read Whole Folders
description: "Read every supported document in a folder tree with file, size and traversal limits, and collect per-file results with OfficeIMO.Reader."
order: 58
---

Folder reading walks a directory, skips files without a registered handler, and reads the rest within explicit limits: maximum files, total bytes, and traversal entries. Each file reports whether it was parsed and any warnings, so one damaged file does not stop the run.

## Read a folder into chunks

### C#

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Word;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddWordHandler()
    .AddPdfHandler()
    .AddMarkdownHandler()
    .WithMaxConcurrentReads(4)
    .Build();

ReaderIngestResult result = reader.ReadFolderDetailed(
    folderPath: "Knowledge",
    folderOptions: new ReaderFolderOptions {
        Recurse = true,
        MaxFiles = 1_000,
        Extensions = new[] { ".docx", ".pdf", ".md" }
    },
    options: ReaderOptions.CreateSafeIngestion(),
    includeChunks: true);

Console.WriteLine($"{result.FilesParsed} of {result.FilesScanned} files, {result.ChunksProduced} chunks");
foreach (ReaderIngestFileResult file in result.Files.Where(f => !f.Parsed)) {
    Console.WriteLine($"Skipped {file.Path}: {string.Join("; ", file.Warnings)}");
}
```

To process chunks as they are produced instead of collecting them, use `reader.EnumerateFolder(...)`, which yields `ReaderChunk` values one file at a time.

### PowerShell

```powershell
$result = Read-OfficeDocument -As Ingest -Path .\Knowledge -Recurse `
    -Extension docx,pdf,md -MaximumFiles 1000 -MaximumTotalBytes 2GB
```

### CLI

Write one Markdown or JSON file per input document:

```sh
officeimo reader folder ./Knowledge --output ./knowledge-md --format markdown --concurrency 4 --max-files 1000
officeimo reader folder ./Knowledge --output ./knowledge-json --format json --no-recursive
```

## Limits and reference

**Limits:** filtering is by extension; there are no glob patterns. Defaults allow 500 files per folder read; raise `MaxFiles` and `MaxTotalBytes` deliberately for large trees.

**Reference:** [OfficeIMO.Reader API](/api/reader/) · [Read-OfficeDocument](/api/powershell/read-officedocument/)
