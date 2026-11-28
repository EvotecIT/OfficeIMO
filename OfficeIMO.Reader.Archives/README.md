# OfficeIMO.Reader.Archives - archive reader adapter

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Reader.Archives)](https://www.nuget.org/packages/OfficeIMO.Reader.Archives)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Reader.Archives?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Reader.Archives)

`OfficeIMO.Reader.Archives` reads documents inside archives for `OfficeIMO.Reader.Core`: ZIP, 7z, Microsoft Cabinet, TAR plain or compressed, single compressed files and, on request, Windows Installer packages. Archives nested inside archives are read through the same path and limits. It replaces `OfficeIMO.Reader.Zip`.

## Install

```powershell
dotnet add package OfficeIMO.Reader.Archives
```

## Register

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Archives;
using OfficeIMO.Archives;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddArchiveHandler(new ArchiveTraversalOptions {
        MaxEntries = 1000,
        MaxPhysicalEntries = 5000,
        MaxDepth = 8,
        MaxTotalUncompressedBytes = 100L * 1024L * 1024L
    })
    .Build();
```

## Examples

### Read supported files inside an archive

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Archives;
using OfficeIMO.Archives;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddCsvHandler()
    .AddJsonHandler()
    .AddArchiveHandler(
        zipOptions: new ArchiveTraversalOptions {
        MaxEntries = 500,
        MaxTotalUncompressedBytes = 200L * 1024L * 1024L
        },
        readerArchiveOptions: new ReaderArchiveOptions {
        ReadNestedArchives = true,
        MaxNestedDepth = 2
        })
    .Build();

foreach (var chunk in reader.Enumerate("evidence.zip", new ReaderOptions {
    MaxInputBytes = 200L * 1024L * 1024L
})) {
    Console.WriteLine($"{chunk.Kind}: {chunk.Location.Path}");
    Console.WriteLine(chunk.Markdown ?? chunk.Text);
}
```

### Surface archive warnings

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Archives;
using OfficeIMO.Archives;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddArchiveHandler(
        zipOptions: new ArchiveTraversalOptions {
        MaxEntries = 100,
        MaxCompressionRatio = 50,
        MaxEntryUncompressedBytes = 10L * 1024L * 1024L
        },
        readerArchiveOptions: new ReaderArchiveOptions {
        ReadNestedArchives = false
        })
    .Build();

var chunks = reader.Enumerate("upload.zip").ToList();
foreach (string warning in chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>())) {
    Console.WriteLine(warning);
}
```

## What it reads

| Input | Registered extensions | Notes |
| --- | --- | --- |
| ZIP | `.zip` | Traversal warnings, duplicate names kept apart, verified members |
| 7z | `.7z` | Members are listed first; only wanted members within the limits are decoded, once per solid folder |
| Microsoft Cabinet | `.cab` | Stored, MSZIP and LZX folders |
| TAR | `.tar`, `.tgz`, `.tzst`, `.txz`, `.tlz` | Plain or compressed with gzip, Zstandard, xz or lzip, detected from content |
| Single compressed file | `.gz`, `.zst`, `.xz`, `.lz` | gzip uses its stored name; the others are named after the input without the suffix |
| Windows Installer | `.msi` | Only with `ReaderArchiveOptions.ReadInstallerPackages`; members are named by their installed path |

The format is chosen from the content, so a renamed archive still reads. RAR is not read yet.

- Warning chunks for skipped or failed entries.
- Path and stream dispatch, including non-seekable stream support.
- Nested archives of every kind above, including `.msi` when installer reading is on, are read with `ReadNestedArchives`, `MaxNestedDepth` and `MaxNestedArchiveBytes`, and draw on the same cumulative budget. A nested archive that is too deep, too large or damaged becomes a warning, and the rest of its parent is still read.

TAR members, 7z members and single compressed files are verified by `OfficeIMO.Archives` (TAR headers, 7z CRCs, codec checksums; plain TAR has no content checksum), and only files a registered handler can read, or nested archives, are materialized. Symbolic links, hard links, devices and directories are never read. `MaxArchiveBytes`, `MaxPhysicalEntries`, `MaxEntryUncompressedBytes` and `MaxTotalUncompressedBytes` bound every format as they do ZIP. A corrupt or unsafe top-level archive fails the read instead of producing per-entry warnings.

The adapter checks compressed size and every physical ZIP entry before opening top-level or nested archives. Path and stream inputs are copied into bounded private snapshots so validation and parsing see the same bytes. Archives borrows that snapshot for inventory and verified reads; it does not create another input copy. The snapshot is released when enumeration completes, fails, is cancelled or is disposed early. Its default source limit is 512 MiB and its default physical entry limit is 10,000. `ReaderOptions.MaxInputBytes` can lower the source limit; raise `ArchiveTraversalOptions.MaxArchiveBytes` or `MaxPhysicalEntries` deliberately for larger trusted archives. A top-level archive above either limit fails before metadata loading; a nested archive is skipped with a warning.

## Boundaries

- Reader adapter configuration belongs here.
- Archive parsing, verification and traversal policy belong in `OfficeIMO.Archives`.
- Shared extraction contracts belong in `OfficeIMO.Reader.Core`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.

## Dependency footprint

- **External:** None beyond platform compression APIs.
- **OfficeIMO:** `OfficeIMO.Reader.Core` and `OfficeIMO.Archives` own safe traversal, nested-archive limits, extraction projection, and warnings.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

## Verified payloads and cumulative archive limits

Reader opens ZIP inventory through `OfficeIMO.Archives` and dispatches a member only after complete length, exact Deflate framing and CRC checks pass. Corrupt ordinary or nested members produce warnings instead of document content. Archive quota exhaustion is translated to the terminal `ReaderResourceLimitException` with its native exception preserved; it is not downgraded to a parse warning.

When `ArchiveTraversalOptions.OperationBudget` is null, each read creates a new cumulative archive budget shared with nested archives. A budget supplied when registering `AddArchiveHandler` is shared by every read through the built reader and is not reset between reads. Use that shared instance to apply explicit tree limits or coordinate related archive operations. The existing Reader document/chunk budgets continue to apply. ZIP input above 64 MiB uses private temporary storage; returned verified members remain buffered and consume memory beyond logical byte limits. Pass `ZipLegacyCodePage` only when the producer's encoding is known; strict ASCII/flagged-UTF-8 admission is the default, and UTF-8 declarations take precedence.
