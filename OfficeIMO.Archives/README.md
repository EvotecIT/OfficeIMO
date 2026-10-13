# OfficeIMO.Archives

OfficeIMO.Archives extracts ZIP archives, Microsoft Cabinet archives and payloads embedded in ordinary Windows Installer databases, applying one safety contract to all three. It runs in managed .NET and uses OfficeIMO.Core for compound storage, CAB decoding, ZIP central-directory checks and shared archive validation. Extraction does not execute an installer, scripts or custom actions.

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

`AllowedFormats` optionally restricts the formats detected from bytes. A disallowed container raises `NotSupportedException` before payload parsing or staging; `null` permits every supported format. This lets callers allow only ZIP and CAB for nested payloads while accepting MSI at the top level.

The destination must be absent or empty. Member paths are checked before files are written. The extractor writes selected files into a fresh sibling staging directory, verifies complete lengths, applicable checksums and MSI alias equality, then moves the completed directory into place. A failure or cancellation removes staging and alias-comparison files and preserves an occupied destination.

Results preserve authored filename case and use forward slashes. MSI files that share an exact target or differ only by ASCII letter case in the final filename are coalesced when every payload has the same declared length and identical bytes. Their parent directory spelling must match exactly. The output uses the ordinal-smallest authored path in the complete alias group, independently of member or cabinet order, and reports that logical file once. Conflicting bytes, directory-case differences and Unicode-normalization collisions fail the entire extraction. Ordinary CAB member collisions remain errors.

MSI results describe the authored target layout; standard Windows folder properties and installation conditions are not evaluated. For example, a package may expose `PFiles/Product/Templates/en-US/Example.adml`. The caller chooses any application-specific root or layout after extraction. MSI source compression follows the Word Count Summary property and each File row's compressed or noncompressed override. Loose source files remain unsupported, including when excluded by the caller.

Required MSI tables may contain additional persistent columns; extraction projects the required fields while accounting for the complete row layout. Each call opens one container, detected from its bytes rather than its file name. A nested archive remains a file in that result. The caller owns any deliberate nesting and overall collection limits. OfficeIMO.Zip remains the read-only traversal API for ZIP-based document formats; use `ArchiveExtractor` when ZIP members must be written to disk.

ZIP extraction accepts single-disk archives whose members are stored or Deflate-compressed and whose ZIP signature starts at byte zero. Prepended self-extracting stubs are unsupported. The central directory is scanned with bounded reads before any member is opened: the member count, multi-disk layout, encryption and compression method are checked first. Layouts that different ZIP readers could resolve to different directories fail: an end-of-central-directory signature inside the comment, ZIP64 values that disagree with unmasked classic fields, file headers beyond the declared directory size, and a classic 65,535-entry archive containing an earlier ZIP64 locator that .NET Framework could follow. ZIP64 member sizes, local-header offsets and disk numbers are resolved before extraction. Every local header must match the central name, flags and compression method, including excluded files. Selected payloads are read through bounded compressed windows so the declared-length check observes the actual decoded end of the stream. Deflate data must contain a final block and end exactly at that block; truncated streams and trailing compressed bytes fail. Every member name, including directory records, passes the same path, device-name, portable-segment and collision checks as CAB and MSI members; directory records publish no empty directories. Duplicate or case-colliding names fail instead of overwriting. Selected declared lengths are charged against the expansion limits before writing; no member writes more than its declared length, a shorter member fails, and every published file matches its CRC-32. Names containing non-ASCII characters must carry the UTF-8 flag and be valid UTF-8, because unflagged names decode through a host-dependent code page. Encrypted members, patched data sets and other compression methods raise `NotSupportedException`.

`ArchiveExtractionOptions` bounds input bytes, member count, per-file and aggregate expansion, CAB decompression work, MSI metadata bytes, per-table rows and string-pool entries, embedded cabinet count, and path length/depth. `MaxTableRows` and `MaxStringPoolEntries` are independent limits, each defaulting to 200,000. Compound structure has separate bounds of 65,536 directory slots and 32,768 streams. MSI extraction indexes compound storage once, checks disjoint regular and mini-sector allocation with exact chain termination for every declared stream, and accepts at most 256 embedded cabinets by default, including cabinets with empty or unselected files. `IncludeFile` receives each normalized relative file path, once per logical target for MSI aliases. Every mapped CAB member length is checked against MSI FileSize even when excluded. Alias payloads are compared even when their logical file is excluded; every source row, member, cabinet and decoded byte retains its original resource charge. Results count only unique published files and their bytes.

CAB folder decompression counts toward expansion and work bounds even when its files are filtered out. The reader buffers bounded input, one decoded folder at a time and selected or compared entry arrays. MSI alias representatives use bounded sibling comparison files rather than retaining arrays across cabinets; those files are removed after comparison or failure. Cancellation is observed during parsing, decoding, comparisons and file writes.

`MaxDecompressionWorkBytes` defaults to 2 GiB and is separate from `MaxExpandedBytes`. Before processing each CAB data frame, the extractor charges the greatest of 32 KiB, the complete frame record size, and the expanded bytes plus any MSZIP history replay. One budget covers every folder and embedded MSI cabinet, including repeated data offsets, gaps and unselected files. A cabinet with no data frames consumes no decompression work. For example, four one-byte MSZIP frames consume 128 KiB of work despite expanding to four bytes. The default permits at most 65,536 frames; larger frames or dictionary replay can exhaust it sooner. These units bound processing and frame setup rather than measuring exact CPU time.

See [SUPPORT.md](SUPPORT.md) for supported compression and MSI profiles. There are no native archive executables to install or download.
