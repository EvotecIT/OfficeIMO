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

## Examples

### Enumerate accepted files only

```csharp
using OfficeIMO.Zip;

IReadOnlyList<ZipEntryDescriptor> entries = ZipTraversal.Enumerate("evidence.zip",
    new ZipTraversalOptions {
        DeterministicOrder = true,
        IncludeDirectoryEntries = false,
        MaxEntries = 2500
    });

foreach (var entry in entries.Where(entry => !entry.IsDirectory)) {
    Console.WriteLine($"{entry.FullName} depth={entry.Depth} bytes={entry.UncompressedLength}");
}
```

### Traverse a stream with defensive limits

```csharp
using OfficeIMO.Zip;

await using var upload = File.OpenRead("upload.zip");

ZipTraversalResult result = ZipTraversal.Traverse(upload, new ZipTraversalOptions {
    MaxDepth = 6,
    MaxEntries = 500,
    MaxEntryUncompressedBytes = 20L * 1024L * 1024L,
    MaxTotalUncompressedBytes = 100L * 1024L * 1024L,
    MaxCompressionRatio = 100
});

if (result.Warnings.Count > 0) {
    foreach (var warning in result.Warnings) {
        Console.WriteLine($"{warning.EntryPath}: {warning.Warning}");
    }
}
```

### Use traversal output before extraction

```csharp
using OfficeIMO.Zip;

var traversal = ZipTraversal.Traverse("incoming.zip");
var safeJsonEntries = traversal.Entries
    .Where(entry => entry.FullName.EndsWith(".json", StringComparison.OrdinalIgnoreCase))
    .ToList();

foreach (var entry in safeJsonEntries) {
    Console.WriteLine($"Queue {entry.FullName} for a separate extraction step.");
}
```

## Boundaries

- This package owns ZIP traversal policy primitives.
- Reader integration belongs in `OfficeIMO.Reader.Zip`.
- It is not a general archive extraction framework.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None beyond platform compression APIs.
- **OfficeIMO:** `OfficeIMO.Drawing`. Traversal policy, safety limits, descriptors, and warnings are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
