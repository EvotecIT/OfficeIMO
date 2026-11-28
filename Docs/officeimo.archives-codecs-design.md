# Archive codecs and containers design

`OfficeIMO.Archives` reads ZIP, 7z, CAB, MSI, TAR and gzip, writes deterministic ZIP, TAR and gzip, and owns extraction safety, budgets, read plans and verified copies. This page sets the target codec and container set beyond that, where each piece lives, the order it lands in and how it is proved. Open work lives in the [roadmap](ROADMAP.md#archive-containers); delivered behaviour moves to the [archive support matrix](../OfficeIMO.Archives/SUPPORT.md).

## Goals

1. **Every codec in managed code.** No native libraries, P/Invoke or new runtime dependencies. Codecs are written from their primary specifications; no code is copied, ported or adapted from other projects.
2. **Hostile input is the normal case.** Limits on input, output, expansion, entries, metadata, dictionary/window size and decode work are on by default and reported through the existing `OfficeInspectionResourceLimitException`. Decoder memory is admitted from the declared window or dictionary before allocation. Malformed input fails with a typed error that names the format, the byte offset and what was expected; no truncated result is returned.
3. **Deterministic output.** Every encoder writes the same bytes for the same input and options on every platform, target framework, CPU and thread count, as the ZIP, TAR and gzip writers already do. SHA-256 pins of reference outputs guard it. Encrypted output is the exception: each encrypted entry gets a fresh random salt (WinZip AES requires a different salt per file, and a reused salt shows whether two encrypted files are identical), so pins cover unencrypted output and encrypted entries are checked by decrypting them.
4. **One API shape.** One-shot `Compress`/`Decompress` methods, a span-based encoder/decoder that reports `OperationStatus`, and a `Stream` wrapper, each taking an options object. Containers keep the existing `ArchiveInspector` / `ArchiveReader` / `ArchiveExtractor` / writer surfaces rather than gaining a parallel facade.
5. **Portable.** `netstandard2.0`, `net472` and current .NET; NativeAOT, trimming and browser WebAssembly. A codec nobody calls is trimmed away: no static registry roots every codec.

## Where OfficeIMO starts

| Area | Today |
| --- | --- |
| Containers read | ZIP (stored, Deflate, ZIP64, descriptors), 7z (Copy, LZMA, LZMA2, encoded headers, solid folders), CAB (Store, MSZIP, LZX), MSI (CFB database, embedded cabinets), TAR (ustar, pax, GNU, v7), .tar.gz, gzip |
| Containers written | ZIP, TAR, .tar.gz, gzip, all byte-identical across platforms and targets |
| Codecs | Deflate encoder (managed, deterministic), Deflate validation in `OfficeIMO.Core` with runtime inflate, LZMA and LZMA2 decoders (7z), LZX decoder (CAB, shared in Core), MSZIP, Zstandard decoder |
| Checksums | CRC-32 (several private copies across Core, Email, OneNote and Png), Adler-32 inside zlib |
| Safety | Portable-path checks, staged atomic extraction, symlink/reparse refusal, per-member and aggregate byte limits, decode-work accounting, shared `ArchiveOperationBudget` across related containers |
| Beyond the usual archive library | MSI payload extraction with File/Directory/Media resolution and alias proof; passive inventory that keeps raw names and ordinals; `PlanRead` cost plans before decoding; atomic grouped reads; verified copies through a private spool; `ArchiveResourceUsage` reports; nested-archive traversal in `OfficeIMO.Reader.Zip` |

## Scope reference

The open-source OneArchive project (packages RecursiveAI.OneCompress and RecursiveAI.OneArchive, MIT or Apache-2.0, `net10.0` only) sets a useful bar for breadth: a single managed library with Zstandard, LZ4, Deflate/zlib/gzip, Deflate64, Brotli, LZMA/LZMA2/xz, lzip, bzip2, PPMd and LZX, plus zip, tar, 7z, cpio, ar, ISO 9660, cab, RAR 5 (stored entries) and compound files. The table lists what that covers against OfficeIMO today. Its README and design notes were read for capabilities only; none of its code is used.

| Capability | OneArchive 1.2 | OfficeIMO today | Plan |
| --- | --- | --- | --- |
| Deflate decode / encode | Yes / yes, levels | Runtime inflate behind managed validation / managed deterministic encoder, one level | Managed inflate (WASM and determinism-independent), encoder levels |
| zlib, gzip | Yes | gzip read and write; zlib internal only | Public zlib |
| Deflate64 (ZIP method 9) | Decode and encode | No | Decision: no formal specification |
| Zstandard | Decode, encode (levels 1–22), dictionaries, training, long-distance matching, multithreaded jobs, seekable format | Decode with dictionaries (`ZstandardArchive`) | Container integration, encoder, training, seekable format |
| LZ4 | Block and frame, levels, dictionaries | No | Decode and encode |
| Brotli | Decode and encode (qualities 0–11) | Runtime `BrotliStream` on .NET 8+ only (WOFF2, CSV) | Managed decode on every target, then encode |
| LZMA, LZMA2 | Decode and encode | Decode (7z) | Encoder |
| xz | CRC-32, CRC-64, SHA-256 checks; delta and branch filters; read and write | No | Read and write; delta filter; branch filters are a decision |
| lzip | Read and write | No | Read and write |
| bzip2 | Read and write | No | Decision: informal specification only |
| PPMd (variants H and I) | Read and write | No | Decision: no specification beyond source code |
| LZX | Decode | Decode (CAB) | Done |
| Checksums | CRC-32, CRC-32C, CRC-64, Adler-32, xxHash32, xxHash64 | CRC-32 (duplicated), Adler-32 | One public checksum set, retiring the private copies |
| ZIP editing | Not offered | Write new archives only | Edit in place: add, remove, rename and replace entries while copying untouched entries' bytes unchanged, with correct zip64 above 4 GiB |
| ZIP | zip64, descriptors, methods store, Deflate, Deflate64, bzip2, Zstandard, PPMd (read and write) and LZMA, xz (read); WinZip AES read and write, ZipCrypto read, damaged-file recovery, lazy central directory, raw entry copy | zip64, descriptors, stored and Deflate, no encryption; inventory spills large input to private storage | AES read and write, ZipCrypto read, recovery mode, new methods as their codecs land (Zstandard 93, LZMA 14, xz 95), raw member copy |
| TAR | ustar, pax, GNU, sparse | ustar, pax, GNU, v7; sparse rejected | Sparse read, `.tar.zst/.xz/.bz2/.lz4/.lz/.br` |
| 7z | Read and write, solid blocks, encoded headers, AES-256; Copy, Deflate, Deflate64, bzip2 (read), LZMA, LZMA2, Zstandard (non-standard method id), PPMd, Delta and branch filters; BCJ2 unsupported | Read: Copy, LZMA, LZMA2, encoded headers, solid; no filters, no encryption, no writer | Delta, BCJ filters, AES read, then a deterministic writer |
| cpio, ar | newc, crc, odc cpio; common, GNU and BSD ar (`.deb` reads as ar); read and write | No | Read, then write; `.deb` through ar |
| ISO 9660 | Joliet, Rock Ridge | No | Read, then write |
| CAB | Stored, MSZIP, LZX read; no Quantum, no writing | Same | Quantum decoding and CAB creation (MSZIP, then LZX), so PowerShell users no longer need `makecab` or `expand.exe` |
| Compound files | Versions 3 and 4, read and write, recovery mode on read | Read and write in Core (MSI, legacy Office) | Recovery reading where a consumer needs it |
| RAR 5 | Listing and stored entries | No | Separate work in progress |
| Small messages | Frame-per-message and flushed-stream modes for Zstandard and LZ4 with trained dictionaries | No | With Zstandard dictionaries |
| Single compressed files | `.gz .bz2 .zst .lz4 .xz .lz`, `.br` by name, `.tar.*` detection | gzip through `ArchiveExtractor`; standalone Zstandard streams through `ZstandardArchive` (not detected automatically yet) | Content-sniffed single-stream reading for every codec that has a signature; Brotli only with an explicit format or a `.br` name, since a Brotli stream has no signature (a single `0x06` byte is a valid empty stream) |
| Format detection | From bytes, then extension | Caller chooses the API | `ArchiveInspector` sniffing across every container and codec with a signature; Brotli is never guessed from content |
| Typed errors | One exception hierarchy with format, offset and expectation | Resource limits typed; `ArchiveCorruptionException` (format, offset, expectation) for new codecs; older readers use `InvalidDataException` | Move the older readers to `ArchiveCorruptionException` |
| Determinism | Byte-identical across OS, CPU, runtime, threads | Byte-identical across OS and target frameworks | Extend to every encoder and to thread count |
| Targets | `net10.0`; AOT, trim, browser WASM | `netstandard2.0`, `net472`, `net8.0`, `net10.0`; AOT smoke | Keep the wider target set; add a WASM smoke |

Where OfficeIMO goes further and keeps going: CAB creation and Quantum decoding (planned), in-place ZIP editing (planned), MSI payloads, `PlanRead` cost plans, atomic grouped reads, verified copies, shared budgets across nested containers, decode-work accounting, passive raw inventory, staged atomic extraction, `.NET Framework` and `netstandard2.0` support, and nested-archive traversal through the Reader. New codecs plug into all of these instead of bypassing them: a Zstandard member inside a ZIP inside a TAR is planned, budgeted and verified like a Deflate member today.

## Package placement

No new packages. Codecs live in `OfficeIMO.Archives` and `OfficeIMO.Core`:

- A codec starts as internal types in `OfficeIMO.Archives`. Single-stream facades such as `ZstandardArchive` sit next to `GZipArchive` in the `OfficeIMO.Archives` namespace. Lower-level public codec types (span encoders and decoders, checksums) go in `OfficeIMO.Archives.Compression`.
- A codec moves into `OfficeIMO.Core` (internal, like LZX and Deflate validation today) when a second package needs it. Brotli goes straight to Core: WOFF2 font decoding in Core and Brotli CSV files only work on .NET 8 and later today, because they use the runtime `BrotliStream`. LZ4 and Zstandard would follow if Arrow IPC buffer compression is added.

## Specifications

| Codec or container | Primary source | Status |
| --- | --- | --- |
| Deflate, zlib, gzip | RFC 1951, 1950, 1952 | Written specification |
| Zstandard, dictionaries | RFC 8878 | Written specification |
| Zstandard seekable format | `zstd/contrib/seekable_format/zstd_seekable_compression_format.md` | Written specification |
| LZ4 block and frame | `lz4_Block_format.md`, `lz4_Frame_format.md` | Written specification |
| xxHash32, xxHash64 | `xxhash_spec.md` | Written specification |
| Brotli | RFC 7932 (static dictionary from the RFC appendix, recorded in `Build/third-party-material.json`) | Written specification |
| xz | The .xz file format 1.2.1 | Written; delta filter specified, branch filters only named |
| LZMA, LZMA2 | `lzma-specification.txt`, 7z format notes; xz format for LZMA2 chunks | Already decoded; encoders need only a valid stream |
| lzip | lzip manual, file format section | Written specification |
| CRC-32C, CRC-64 | RFC 3720 (Castagnoli), ECMA-182 | Written specification |
| ZIP, zip64, ZipCrypto | PKWARE APPNOTE 6.3.10 | Written specification |
| WinZip AES | WinZip AE-1/AE-2 specification | Written specification |
| ISO 9660, Joliet, Rock Ridge | ECMA-119, Joliet specification, IEEE P1281 SUSP and P1282 RRIP | Written specification |
| TAR, sparse | POSIX pax/ustar, GNU tar manual (sparse formats 0.0, 0.1, 1.0) | Written specification |
| cpio | POSIX odc, SVR4 newc/crc descriptions | Written specification |
| ar | System V/GNU and BSD `ar` format descriptions | Written specification |
| 7z, AES | `7zFormat.txt` | Container written; AES key derivation not specified |
| bzip2 | No formal specification; informal format descriptions only | Decision |
| Deflate64 | APPNOTE names method 9 without defining it | Decision |
| PPMd H and I | Only the reference source | Decision |
| BCJ/branch filters (x86, ARM, ARM64, PowerPC, SPARC, IA-64, RISC-V) | Named in xz and 7z formats, defined only by source | Decision |
| Quantum (CAB) | No Microsoft specification; only independent descriptions of the format | Planned; which description counts as the source is part of the open specification decision |
| CAB creation | [MS-CAB] for the container and MSZIP; for LZX, the LZX data compression format document from Microsoft's Cabinet SDK ([MS-PATCH] LZX DELTA is a derivative, not the CAB format itself) | MSZIP: written specification. LZX: confirm that document as the source before the encoder starts |
| RAR | Separate work | Out of scope here |

## Decisions needed

Package placement is settled above. Open: which sources are acceptable for formats defined only by another project's source code or by informal descriptions.

1. **bzip2:** implement from informal format descriptions and validate only against the reference tool's output, or leave out.
2. **Deflate64:** same question; the only definition is behavioural (64 KiB window, extended length and distance codes).
3. **PPMd:** no specification exists apart from source code. Leave out unless a written description is accepted.
4. **Branch filters and 7z AES key derivation:** defined only by source code. Leave out, or accept independent descriptions.
5. **CAB Quantum:** no Microsoft specification; independent descriptions exist.

These stay out until the decision is made.

## Delivery order

Each step is one PR with interop fixtures produced by the reference tools through a committed script that records tool versions and hashes, malformed-input tests that assert the typed error and offset, default-on limits, and determinism pins for encoders.

1. Zstandard decoder with `ArchiveCorruptionException`: frames, skippable frames, dictionaries, window limits.
2. ZIP encrypted-member reading: traditional ZipCrypto (APPNOTE) and WinZip AES (AE-1 and AE-2), password supplied through options.
3. TAR wrappers as each codec lands: `.tar.zst`, `.tar.xz`, `.tar.lz` now, `.tar.lz4` with step 6, `.tar.br` (by name or explicit format) with step 7, and `.tar.bz2` after the decision, read alongside `.tar.gz`, written once the encoder exists; content-sniffed single-stream detection in the extractor.
4. `OfficeIMO.Reader.Zip` becomes `OfficeIMO.Reader.Archives` for 4.0: 7z, CAB and RAR stored entries read automatically like ZIP, TAR and gzip, MSI opt-in, nested archives of any supported kind through the same path and limits.
5. xz and lzip readers over the existing LZMA/LZMA2 decoders: checks, delta filter, multi-stream and padding.
6. LZ4 frame and block decoder, then encoder.
7. Managed Brotli decoder in Core on every target, taking WOFF2 and CSV off the .NET 8 requirement.
8. Typed corruption in the existing readers and a shared checksum set (CRC-32, CRC-32C, CRC-64, Adler-32, xxHash32, xxHash64).
9. Managed inflate, so decoding no longer depends on the runtime's zlib.
10. Encoders: Zstandard levels, LZMA/LZMA2 and xz, Brotli, Deflate levels, all deterministic.
11. ZIP editing in place, ZIP AES writing, recovery mode for damaged ZIPs, additional ZIP methods.
12. CAB creation (MSZIP, then LZX) and Quantum decoding (after the decision).
13. Containers: cpio and ar, ISO 9660 with Joliet and Rock Ridge, TAR sparse reading.
14. 7z delta filter, AES read (after the decision), deterministic 7z writer.
15. Zstandard dictionary training, seekable format and multithreaded compression with thread-count-independent output.

## Benchmarks

`OfficeIMO.Archives.Benchmarks` already runs a OneArchive 1.2.0 lane on `net10.0` for ZIP traversal. Codec comparisons join that project, which stays outside `OfficeIMO.slnx` so comparison packages never enter normal restore:

- One lane per codec and direction (decompress, then compress at matched levels), on a pinned corpus (`silesia.tar` plus small-message and archive workloads), with OneArchive and the runtime codecs where they exist.
- Validation before timing: every lane must decode to the same SHA-256, and compress lanes must round-trip through the reference tool. Lanes doing different work (for example skipping checksum verification) are not compared.
- Compression rows compare size and speed together; a level-for-level comparison is not meaningful across libraries.
- Tests that measure time or memory carry `[Trait("Category", "Performance")]`; CI never runs them. Evidence runs happen on an idle machine and record the machine, runtime and package versions.
