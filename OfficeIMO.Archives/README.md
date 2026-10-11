# OfficeIMO.Archives

OfficeIMO.Archives extracts Microsoft Cabinet archives and payloads embedded in ordinary Windows Installer databases. It runs in managed .NET and uses OfficeIMO.Core for compound storage, CAB decoding and shared archive validation. Extraction does not execute an installer, scripts or custom actions.

## Install

```powershell
dotnet add package OfficeIMO.Archives
```

## Extract policy files

```csharp
using OfficeIMO.Archives;

var result = ArchiveExtractor.Extract(
    "AdministrativeTemplates.msi",
    "extracted-templates",
    new ArchiveExtractionOptions {
        IncludeFile = path => path.EndsWith(".admx", StringComparison.OrdinalIgnoreCase)
                           || path.EndsWith(".adml", StringComparison.OrdinalIgnoreCase),
        MaxInputBytes = 128L * 1024 * 1024,
        MaxExpandedBytes = 256L * 1024 * 1024
    });

foreach (var file in result.Files)
    Console.WriteLine($"{file.RelativePath}: {file.Length} bytes");
```

The destination must be absent or empty. Member paths and collisions are checked before files are written. The extractor writes selected files into a fresh sibling staging directory, verifies complete lengths and applicable checksums, then moves the completed directory into place. A failure or cancellation removes the staging directory and preserves an occupied destination.

Results preserve authored filename case and use forward slashes. MSI results describe the authored target layout; standard Windows folder properties and installation conditions are not evaluated. For example, a package may expose `PFiles/Product/Templates/en-US/Example.adml`. The caller chooses any application-specific root or layout after extraction.

Each call opens one container. A nested archive remains a file in that result. The caller owns ZIP processing, any deliberate nesting and overall collection limits; OfficeIMO.Zip provides shared ZIP traversal primitives.

`ArchiveExtractionOptions` bounds input bytes, member count, per-file and aggregate expansion, CAB decompression work, MSI metadata and rows, embedded cabinet count, and path length/depth. MSI extraction indexes compound storage once and accepts at most 256 embedded cabinets by default, including cabinets with empty or unselected files. `IncludeFile` receives each normalized relative file path; safety and collision checks still apply to unselected members. CAB folder decompression counts toward the expansion and work bounds even when its files are filtered out. The reader buffers bounded input, one decoded folder at a time and selected entry arrays. Cancellation is observed during parsing, decoding and file writes.

`MaxDecompressionWorkBytes` defaults to 2 GiB and is separate from `MaxExpandedBytes`. Before processing each CAB data frame, the extractor charges the greatest of 32 KiB, the complete frame record size, and the expanded bytes plus any MSZIP history replay. One budget covers every folder and embedded MSI cabinet, including repeated data offsets, gaps and unselected files. A cabinet with no data frames consumes no decompression work. For example, four one-byte MSZIP frames consume 128 KiB of work despite expanding to four bytes. The default permits at most 65,536 frames; larger frames or dictionary replay can exhaust it sooner. These units bound processing and frame setup rather than measuring exact CPU time.

See [SUPPORT.md](SUPPORT.md) for supported compression and MSI profiles. There are no native archive executables to install or download.
