# OfficeIMO.Archives

OfficeIMO.Archives inspects ZIP and Microsoft Cabinet archives and ordinary Windows Installer databases, reads complete verified members, and extracts ZIP, CAB and MSI payloads to disk. It runs in managed .NET and uses OfficeIMO.Core for compound storage, CAB decoding and shared ZIP validation. Extraction does not execute an installer, scripts or custom actions.

## Install

```powershell
dotnet add package OfficeIMO.Archives
```

## Inspect and read a member without extraction

```csharp
using OfficeIMO.Archives;
using System.Linq;

using var inspection = ArchiveInspector.Open("AdministrativeTemplates.msi",
    new ArchiveInspectionOptions {
        MaxInputBytes = 128L * 1024 * 1024,
        MaxEntryBytes = 16L * 1024 * 1024,
        MaxExpandedBytes = 256L * 1024 * 1024
    });

foreach (var entry in inspection.Entries) {
    Console.WriteLine($"{entry.Ordinal}: {entry.RawName}, {entry.Length} bytes, {entry.Compression}");
    if (!entry.CanRead) {
        Console.WriteLine(entry.Diagnostic?.Message);
        continue;
    }
}

var ordinals = inspection.Entries
    .Where(entry => entry.CanRead && entry.AuthoredPath?.EndsWith(".admx", StringComparison.OrdinalIgnoreCase) == true)
    .Select(entry => entry.Ordinal).ToArray();
using var batch = inspection.ReadMembers(ordinals, maxSelectedBytes: 64L * 1024 * 1024);
foreach (var member in batch.Members) {
    // Each stream contains a complete length-validated member for parsing or hashing.
    Console.WriteLine($"Read {member.Content.Length} bytes from {member.Entry.AuthoredPath}");
}
```

`Entries` preserves ZIP central-directory order, CAB CFFILE order or MSI File-table row order. ZIP entries retain raw decoded paths, compressed lengths, compression methods, directory flags and DOS timestamps. `RawName` retains the ZIP/CAB member name or MSI File identifier. MSI entries also retain the unmodified `FileName` field, File sequence, Media cabinet reference and assembled authored target path. Aliases and duplicate CAB names stay separate and are selected by ordinal. Inspection preserves names such as traversal paths because it never uses member names to write files. The filesystem safety checks below apply to extraction.

The inventory checks bounded directory structure, declared sizes, member ranges and embedded MSI File/Media mappings. It does not decode payloads or verify their checksums. `CanRead` means the declared source and compression are supported. `OpenRead` verifies the complete containing folder, applicable checksums, decoded boundaries and member length before returning a read-only stream. `FolderFrameCount` and `FolderChecksumFrameCount` describe the containing folder and its nonzero checksum declarations. A successful read checks exactly those nonzero frame checksums; zero-checksum frames have no checksum verification. CAB checksums detect corruption and do not prove authenticity. It returns no prefix or partial member on failure. A caller can hash the complete returned stream; the inventory supplies no content hash.

`ReadMembers(ordinals, maxSelectedBytes, cancellationToken)` reads selected members together. It decodes each selected cabinet folder once, including its gaps and unselected siblings, and preserves the requested ordinal order. The selected payload byte sum, distinct ordinals, supported sources and remaining expansion allowance are checked before decoding. Duplicate names and aliases retain their separate inventory identities; repeating the same ordinal in a request is an error. A corruption, cancellation or work-limit failure returns no batch or partial member. Each call retains its complete selected payloads in memory under the explicit `maxSelectedBytes` bound. `SelectedBytes`, `DecodedFolderBytes` and `DecompressionWorkBytes` report retained payload, complete folder expansion and charged frame-processing work respectively. These limits bound parser buffers and retained payloads; they do not establish a hard process peak-memory limit.

External Media cabinet references and loose MSI source rows remain visible with `MSI_EXTERNAL_CABINET` or `MSI_LOOSE_SOURCE` diagnostics. Their paths and URLs are never opened. Unsupported CAB compression, including Quantum or unsupported LZX parameters, remains visible with a per-member diagnostic. Cabinet sets, continued records, unsupported name encodings and unsupported MSI metadata profiles still fail explicitly.

For MSI, `StorageStreams` exposes exact raw compound names and declared sizes, including metadata, Binary content and embedded cabinets. `OpenStorageStream(ordinal)` reads a complete inert raw stream under the same per-member and cumulative byte limits. It does not decode an embedded cabinet or interpret Binary content. Encoded MSI compound names remain encoded in this raw inventory.

Inspection snapshots a bounded readable input. ZIP input above 64 MiB spills to private temporary storage; smaller ZIP inputs remain in memory. CAB/MSI inputs and decoded folders retain their managed buffer limits. For seekable streams it reads from the beginning, restores the caller's position on success or failure, and leaves the source open. Nonseekable streams are consumed from their current position through EOF under the same input limit and are left open. Subsequent reads use the immutable snapshot and their own cancellation tokens. Dispose the inspection to release its input and indexes; already returned streams remain caller-owned and usable. Calls on one inspection are serialized. The caller also owns memory held by returned streams.

Dispose a returned batch to close all its member streams and release its member collection. Its streams remain usable after inspection disposal and may also be closed individually. Captured inventory records and batch byte/work totals remain available after disposal. Reading the same folder in another batch consumes expansion and work again; there is no retained decoded-folder cache.

`ArchiveInspectionOptions` bounds input, entries, cabinets, individual member bytes, declared aggregate expansion, MSI metadata, table rows, string-pool entries, path structure and CAB directory records. `MaxDirectoryRecords` defaults to 1,000,000 and covers CAB folder, file and CFDATA records across all inventoried cabinets; MSI compound storage additionally retains its 65,536-slot and 32,768-stream bounds. Inventory checks file-byte and folder-byte sums separately against `MaxExpandedBytes`. Subsequent reads share a fresh cumulative expansion allowance of that size and one `MaxDecompressionWorkBytes` allowance. Each `OpenRead` charges its entire folder again; each batch charges each selected folder once, including gaps and sibling files. Raw storage reads charge their complete stream length. Failed attempted decodes retain their folder charge, while pre-canceled or preflight-rejected reads do not consume bytes. `ReadBytes` reports that cumulative charge; inspection `DecompressionWorkBytes` reports frame work actually charged, including attempts that fail after charging. No temporary extraction or comparison files are created.

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

Results preserve filename case and use forward slashes. MSI directory and file names use their Windows-installed spelling: trailing ASCII spaces and dots are removed before validation, selection and collision checks. A directory authored as `.` still means its parent; names that normalize entirely away are rejected. Authored long names are limited to 255 characters, short/long pairs to 511, and target/source directory definitions to 1,023 before splitting or trimming, including when the MSI schema permits unbounded strings.

MSI files that share a normalized target or differ only by ASCII letter case in the final filename are coalesced when every payload has the same declared length and identical bytes. Their normalized parent directory spelling must match exactly. The output uses the ordinal-smallest normalized path in the complete alias group, independently of member or cabinet order, and reports that logical file once. Conflicting bytes, directory-case differences and Unicode-normalization collisions fail the entire extraction. Ordinary CAB member collisions remain errors and CAB names retain strict portable-segment validation.

MSI results describe the target layout with installed name spelling; standard Windows folder properties and installation conditions are not evaluated. For example, a package may expose `PFiles/Product/Templates/en-US/Example.adml`. The caller chooses any application-specific root or layout after extraction. MSI source compression follows the Word Count Summary property and each File row's compressed or noncompressed override. Loose source files remain unsupported, including when excluded by the caller.

Required MSI tables may contain additional persistent columns; extraction projects the required fields while accounting for the complete row layout. Each call opens one container, detected from its bytes rather than its file name. A nested archive remains a file in that result. The caller owns any deliberate nesting and overall collection limits. Use `ArchiveReader.TraverseZip` for ZIP metadata and `ArchiveExtractor` when ZIP members must be written to disk.

ZIP extraction accepts single-disk archives whose members are stored or Deflate-compressed and whose ZIP signature starts at byte zero. Prepended self-extracting stubs are unsupported. The central directory is scanned with bounded reads before any member is opened: the member count, multi-disk layout, encryption and compression method are checked first. Layouts that different ZIP readers could resolve to different directories fail: an end-of-central-directory signature inside the comment, ZIP64 values that disagree with unmasked classic fields, file headers beyond the declared directory size, and a classic 65,535-entry archive containing an earlier ZIP64 locator that .NET Framework could follow. ZIP64 member sizes, local-header offsets and disk numbers are resolved before extraction. Every local header must match the central name, flags and compression method, including excluded files. Selected payloads are read through bounded compressed windows so the declared-length check observes the actual decoded end of the stream. Deflate data must contain a final block and end exactly at that block; truncated streams and trailing compressed bytes fail. Every member name, including directory records, passes the same path, device-name, portable-segment and collision checks as CAB and MSI members; directory records publish no empty directories. Duplicate or case-colliding names fail instead of overwriting. Selected declared lengths are charged against the expansion limits before writing; no member writes more than its declared length, a shorter member fails, and every published file matches its CRC-32. By default, names containing non-ASCII characters must carry the UTF-8 flag and be valid UTF-8. Set `ZipLegacyCodePage` explicitly on inspection, read, extraction or traversal options to admit unflagged legacy names, for example `1250` for a known Windows Central European producer or `65001` for known unflagged UTF-8. The UTF-8 flag always takes precedence. No host-locale guessing or global encoding-provider registration occurs; decoded extraction paths still undergo collision and portable-name checks. Encrypted members, patched data sets and other compression methods raise `NotSupportedException`.

`ArchiveExtractionOptions` bounds input bytes, member count, per-file and aggregate expansion, CAB decompression work, MSI metadata bytes, per-table rows and string-pool entries, embedded cabinet count, and path length/depth. `MaxTableRows` and `MaxStringPoolEntries` are independent limits, each defaulting to 200,000. Compound structure has separate bounds of 65,536 directory slots and 32,768 streams. MSI extraction indexes compound storage once, checks disjoint regular and mini-sector allocation with exact chain termination for every declared stream, and accepts at most 256 embedded cabinets by default, including cabinets with empty or unselected files. `IncludeFile` receives each normalized relative file path, once per logical target for MSI aliases. Every mapped CAB member length is checked against MSI FileSize even when excluded. Alias payloads are compared even when their logical file is excluded; every source row, member, cabinet and decoded byte retains its original resource charge. Results count only unique published files and their bytes.

CAB folder decompression counts toward expansion and work bounds even when its files are filtered out. The reader buffers bounded input, one decoded folder at a time and selected or compared entry arrays. MSI alias representatives use bounded sibling comparison files rather than retaining arrays across cabinets; those files are removed after comparison or failure. Cancellation is observed during parsing, decoding, comparisons and file writes.

`MaxDecompressionWorkBytes` defaults to 2 GiB and is separate from `MaxExpandedBytes`. Before processing each CAB data frame, the extractor charges the greatest of 32 KiB, the complete frame record size, and the expanded bytes plus any MSZIP history replay. One budget covers every folder and embedded MSI cabinet, including repeated data offsets, gaps and unselected files. A cabinet with no data frames consumes no decompression work. For example, four one-byte MSZIP frames consume 128 KiB of work despite expanding to four bytes. The default permits at most 65,536 frames; larger frames or dictionary replay can exhaust it sooner. These units bound processing and frame setup rather than measuring exact CPU time.

See [SUPPORT.md](SUPPORT.md) for supported compression and MSI profiles. There are no native archive executables to install or download.

## Read ZIP content in memory

`ArchiveReader.ReadZip` returns complete selected files after validating their lengths, Deflate framing and CRC-32. It uses the same strict name, path, collision and payload rules as extraction. It creates no filesystem output and executes no content.

```csharp
using OfficeIMO.Archives;

var files = ArchiveReader.ReadZip(packageBytes, new ArchiveReadOptions {
    IncludeFile = path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase),
    MaxInputBytes = 32L * 1024 * 1024,
    MaxEntryBytes = 8L * 1024 * 1024,
    MaxExpandedBytes = 64L * 1024 * 1024
});
foreach (var file in files)
    Console.WriteLine($"{file.RelativePath}: {file.Length} bytes");
```

The stream overload reads from the current position through EOF into a bounded snapshot, supports nonseekable streams and leaves the input open. Returned arrays belong to the caller and remain usable after the input closes. A failure returns no partial result. This read API accepts ZIP; use `ArchiveInspector` for CAB and MSI inventory and in-memory member reads. Limits bound logical input and decoded content; buffers and result copies add managed-memory overhead.

For metadata enumeration with ingestion warnings, use `ArchiveReader.TraverseZip` and `ArchiveTraversalOptions`. Traversal can skip individual entries and report warnings; it does not verify selected file contents. `ArchiveReader.ValidateZipSource` preflights a seekable source before a caller opens its own `ZipArchive`; keep those bytes immutable until that archive has been read. Stream/path traversal overloads take their own bounded snapshot.

Configured quota exhaustion throws `OfficeIMO.OfficeInspectionResourceLimitException` with `Owner = "archives"`, the stable limit property or parameter name in `LimitName`, and its effective `Maximum`. This includes `maxSelectedBytes` for a batch and fixed MSI compound ceilings (`MaxDirectoryRecords` and `MaxStorageStreamCount`). A reduced native remaining allowance can report zero; an array-backed read reports the smaller of its requested byte limit and `Int32.MaxValue`. Native structured failures remain in `InnerException`. Catch this type separately when resource failures must stop a larger operation. `InvalidDataException` continues to mean malformed, inconsistent or unsafe content; `NotSupportedException` identifies unsupported source profiles. Callers that previously caught `InvalidDataException` for quotas must update that catch. Quota failures preserve extraction rollback. Inspection stream overloads restore the caller’s position; `ArchiveReader.ReadZip(Stream, ...)` advances by the bytes consumed, including on failure. ZIP traversal retains its separate warning and admission behavior.

## Verify, copy and budget related containers

```csharp
var budget = new ArchiveOperationBudget(new ArchiveOperationLimits {
    MaxInputBytes = 256L * 1024 * 1024,
    MaxExpandedBytes = 512L * 1024 * 1024,
    MaxSelectedBytes = 128L * 1024 * 1024,
    MaxDepth = 3
});
using var inventory = ArchiveInspector.Open("bundle.zip", new ArchiveInspectionOptions {
    OperationBudget = budget
});
var verification = inventory.Verify();
foreach (var member in verification.Members)
    Console.WriteLine($"{member.Entry.Ordinal}: {member.Status} {member.Code}");

// Use budget.CreateNested() when opening a returned member as another container.
// OpenRead and ReadMembers also work for ZIP, preserving duplicate source ordinals.
```

`Verify` returns structured `Verified`, `Unsupported` and `Invalid` outcomes, retaining each entry and stable diagnostic code. CAB selections sharing a folder decode that folder once. ZIP verification streams to a discard sink. A failed folder marks all selected members of that folder invalid; it does not claim which member caused the corruption. Cancellation and quota exhaustion throw instead of returning an incomplete report. Verification checks supported content integrity, not cryptographic authenticity.

`inventory.CopyTo(ordinal, destinationStream, cancellationToken)` writes a verified member at the destination's current position and leaves that stream open. ZIP content is decoded to a private temporary file and completely verified before the destination receives bytes, avoiding a complete decoded-member array. CAB/MSI use their existing complete in-memory member reads. Corrupt content leaves the destination untouched. Cancellation or a destination I/O failure during the final copy can leave a verified prefix; the caller owns output rollback. Temporary files are removed when the operation exits. ZIP input spills above 64 MiB and remains bounded by `MaxInputBytes`; CAB/MSI input and complete in-memory member reads retain managed array limits.

Pass the same `ArchiveOperationBudget` to related operations and `CreateNested()` to child containers. It tracks cumulative compressed input, physical entries, attempted expanded bytes, selected payload admission, processing work and maximum depth across inspection, reading, verification, copying and extraction. MSI File rows and embedded CAB records both count as physical entries. Repeated and failed attempts keep charges already reserved; disposing results does not refund them. These are logical admission limits, not measurements of live retained memory or process RSS. Existing per-container limits still apply. Without an explicit shared budget, standalone Archives calls retain their existing per-container allowances.

`ResourceUsage` on an inspection, verification result or extraction result reports a detached snapshot of the shared counters. Quota exceptions use `Operation.MaxInputBytes`, `Operation.MaxEntryCount`, `Operation.MaxExpandedBytes`, `Operation.MaxSelectedBytes`, `Operation.MaxDecompressionWorkBytes` or `Operation.MaxDepth` for shared-limit identity. CAB processing charges each reached frame as described above, including unselected solid-folder work. ZIP reads reserve compressed bytes plus twice declared expansion to account for framing validation and decoding; these units are not exact CPU time. Metadata-only traversal charges input and physical entries, and traversal of an existing inspection does not charge its input twice.

See [the opt-in qualification consumer](../Build/Archives/README.md) for independent producer and package validation, including evidence boundaries.
