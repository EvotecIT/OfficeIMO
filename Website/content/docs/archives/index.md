---
title: "ZIP, CAB and MSI Archives"
description: "Inspect, verify, read and safely extract ZIP, Microsoft Cabinet and Windows Installer payloads in managed .NET with explicit resource limits."
order: 58
meta.seo_title: "Read and extract ZIP, CAB and MSI files in C#"
verified: 2026-11-14
---

`OfficeIMO.Archives` inspects ZIP and Microsoft Cabinet archives and ordinary Windows Installer (MSI) databases, reads complete verified members into memory and extracts payloads to disk. It is fully managed, needs no native archive tools and never runs installer actions, scripts or custom actions.

```shell
dotnet add package OfficeIMO.Archives
```

Targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2.

## Inspect and read members

`ArchiveInspector.Open` builds an ordinal-preserving inventory without decoding payloads. Reads then return complete, length-verified streams, checking every ZIP CRC and each CAB frame checksum that is present (zero-checksum CAB frames have nothing to check); a failed read returns nothing partial.

```csharp
using OfficeIMO.Archives;

using var inspection = ArchiveInspector.Open("AdministrativeTemplates.msi", new ArchiveInspectionOptions {
    MaxInputBytes = 128L * 1024 * 1024,
    MaxEntryBytes = 16L * 1024 * 1024,
    MaxExpandedBytes = 256L * 1024 * 1024
});

var ordinals = inspection.Entries
    .Where(entry => entry.CanRead && entry.AuthoredPath?.EndsWith(".admx", StringComparison.OrdinalIgnoreCase) == true)
    .Select(entry => entry.Ordinal).ToArray();

var plan = inspection.PlanRead(ordinals);
Console.WriteLine($"Selected {plan.SelectedBytes} bytes, decoding {plan.DecodedBytes} bytes");

using var batch = inspection.ReadMembers(ordinals, maxSelectedBytes: 64L * 1024 * 1024);
foreach (var member in batch.Members)
    Console.WriteLine($"{member.Entry.AuthoredPath}: {member.Content.Length} bytes");
```

Entries keep ZIP central-directory order, CAB file order or MSI File-table order. Duplicate names and MSI aliases stay separate and are selected by ordinal. `PlanRead` estimates cost without decoding, and `ReadMembers` decodes each selected cabinet folder once. `Verify()` returns `Verified`, `Unsupported` or `Invalid` per member, and `CopyTo` writes one verified member to a stream.

## Extract to disk

```csharp
using OfficeIMO.Archives;

var result = ArchiveExtractor.Extract("AdministrativeTemplates.msi", "extracted-templates",
    new ArchiveExtractionOptions {
        IncludeFile = path => path.EndsWith(".admx", StringComparison.OrdinalIgnoreCase)
                           || path.EndsWith(".adml", StringComparison.OrdinalIgnoreCase),
        MaxInputBytes = 128L * 1024 * 1024,
        MaxExpandedBytes = 256L * 1024 * 1024
    });

foreach (var file in result.Files)
    Console.WriteLine($"{file.RelativePath}: {file.Length} bytes");
```

The format is detected from bytes, not the file name. The destination must be absent or empty: files are written to a sibling staging directory, verified, then moved into place. Paths are checked for traversal, absolute paths, alternate data streams, reserved device names and case collisions before anything is written. A failure or cancellation removes staging output.

## Read ZIP content in memory

```csharp
using OfficeIMO.Archives;

var files = ArchiveReader.ReadZip(File.ReadAllBytes("package.zip"), new ArchiveReadOptions {
    IncludeFile = path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase),
    MaxEntryBytes = 8L * 1024 * 1024
});
foreach (var file in files)
    Console.WriteLine($"{file.RelativePath}: {file.Length} bytes");
```

`ReadZip` applies the same name, path and CRC-32 rules as extraction and creates no files. `ArchiveReader.TraverseZip` enumerates metadata with ingestion warnings without verifying content.

## What it covers

| Container | Supported profile |
|---|---|
| ZIP | Single-disk archives with stored or Deflate members, ZIP64, UTF-8 names; legacy code pages only through an explicit `ZipLegacyCodePage` |
| CAB | Single cabinets with Store, MSZIP and LZX compression |
| MSI | Ordinary installer databases with Directory, Component, File and Media tables and embedded cabinets; raw storage streams are listed |

Every operation is bounded: input size, entry count, per-member and aggregate expansion, decompression work, MSI metadata and path length and depth. Pass one `ArchiveOperationBudget` to related operations, and `CreateNested()` to child containers, to share those limits. Quota failures throw `OfficeInspectionResourceLimitException` with the exact limit name; malformed content throws `InvalidDataException`.

## Limits

- Encrypted ZIP members, other compression methods, multi-disk archives and prepended self-extracting stubs are rejected. Ambiguous ZIP layouts fail rather than being guessed.
- Cabinet sets, continued records and Quantum compression are not supported.
- External MSI cabinets and loose source files are listed with diagnostics but never opened. Transforms, patches, features, conditions and properties are not evaluated, so extraction shows the authored layout rather than a simulated installation.
- CAB checksums and ZIP CRC-32 detect corruption; they do not prove authenticity.
- Limits bound logical data, not process peak memory.

See the [OfficeIMO.Archives README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Archives/README.md) and the [OfficeIMO.Archives support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Archives/SUPPORT.md).
