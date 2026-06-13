# OfficeIMO.Zip - safe ZIP traversal primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Zip)](https://www.nuget.org/packages/OfficeIMO.Zip)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Zip?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Zip)

`OfficeIMO.Zip` provides dependency-light ZIP traversal primitives for ingestion scenarios.

## Install

```powershell
dotnet add package OfficeIMO.Zip
```

## Quick start

```csharp
using OfficeIMO.Zip;

ZipTraversalResult result = ZipTraversal.Traverse("archive.zip", new ZipTraversalOptions {
    MaxEntries = 1000,
    MaxDepth = 8,
    MaxTotalUncompressedBytes = 100L * 1024L * 1024L
});

foreach (ZipEntryDescriptor entry in result.Entries) {
    Console.WriteLine($"{entry.FullName} ({entry.UncompressedLength} bytes)");
}

foreach (ZipTraversalWarning warning in result.Warnings) {
    Console.WriteLine($"{warning.EntryPath}: {warning.Warning}");
}
```

## What it does

- Enumerates entries deterministically.
- Applies path safety guards for relative traversal, absolute paths, and drive paths.
- Enforces depth, entry-count, uncompressed-size, per-entry-size, and compression-ratio limits.
- Reports traversal warnings for rejected or limited entries.

## Boundaries

- This package owns ZIP traversal policy primitives.
- Reader integration belongs in `OfficeIMO.Reader.Zip`.
- It is not a general archive extraction framework.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
