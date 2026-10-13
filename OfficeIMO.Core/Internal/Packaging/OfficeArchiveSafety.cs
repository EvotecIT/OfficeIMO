using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Core.Internal;

/// <summary>
/// Centralizes archive-entry safety rules shared by ZIP-backed OfficeIMO format owners.
/// </summary>
internal static partial class OfficeArchiveSafety {
    private const uint CentralDirectoryFileHeaderSignature = 0x02014b50U;
    private const uint CentralDirectoryDigitalSignature = 0x05054b50U;
    private const uint EndOfCentralDirectorySignature = 0x06054b50U;
    private const uint Zip64EndOfCentralDirectorySignature = 0x06064b50U;
    private const uint Zip64EndOfCentralDirectoryLocatorSignature = 0x07064b50U;

    /// <summary>Normalizes an archive entry name to forward-slash notation.</summary>
    internal static string NormalizeEntryName(string? fullName) {
        string value = fullName ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;

        string normalized = value.Replace('\\', '/').Trim();
        while (normalized.StartsWith("./", StringComparison.Ordinal)) {
            normalized = normalized.Substring(2);
        }

        return normalized;
    }

    /// <summary>Returns true when an entry name is absolute, traverses parents, or contains a NUL.</summary>
    internal static bool IsUnsafePath(string fullName) {
        if (fullName.Length == 0) return true;
        if (fullName[0] == '/' || fullName[0] == '\\') return true;
        if (ContainsNull(fullName)) return true;
        if (fullName.Length >= 2 && char.IsLetter(fullName[0]) && fullName[1] == ':') return true;

        int segmentStart = 0;
        for (int index = 0; index <= fullName.Length; index++) {
            if (index < fullName.Length && fullName[index] != '/') continue;

            int segmentLength = index - segmentStart;
            if (segmentLength == 1 && fullName[segmentStart] == '.') return true;
            if (segmentLength == 2
                && fullName[segmentStart] == '.'
                && fullName[segmentStart + 1] == '.') {
                return true;
            }
            segmentStart = index + 1;
        }

        return false;
    }

    /// <summary>Returns the logical path depth of an archive entry.</summary>
    internal static int ComputeDepth(string fullName, bool isDirectory) {
        string normalized = isDirectory ? fullName.TrimEnd('/') : fullName;
        if (normalized.Length == 0) return 0;

        int depth = 1;
        for (int i = 0; i < normalized.Length; i++) {
            if (normalized[i] == '/') depth++;
        }

        return depth;
    }

    /// <summary>Reads an entry's declared uncompressed length without allowing metadata failures to escape.</summary>
    internal static bool TryGetLength(ZipArchiveEntry entry, out long length) {
        try {
            length = entry.Length;
            return true;
        } catch {
            length = 0;
            return false;
        }
    }

    /// <summary>Checks an entry's declared expansion ratio.</summary>
    internal static bool IsCompressionRatioExceeded(ZipArchiveEntry entry, long uncompressedLength, double maxRatio) {
        if (maxRatio <= 0 || uncompressedLength <= 0) return false;

        long compressedLength;
        try {
            compressedLength = entry.CompressedLength;
        } catch {
            return false;
        }

        if (compressedLength <= 0) return false;
        return (double)uncompressedLength / compressedLength > maxRatio;
    }

    /// <summary>
    /// Reads exactly the declared entry length and rejects truncated or
    /// over-expanding payloads before materializing bytes beyond that bound.
    /// </summary>
    internal static byte[] ReadEntryBytes(Stream source, long declaredLength, long maximumLength) =>
        ReadEntryBytes(source, declaredLength, maximumLength, CancellationToken.None);

    internal static byte[] ReadEntryBytes(Stream source,
        long declaredLength, long maximumLength, CancellationToken cancellationToken) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (declaredLength < 0 || declaredLength > maximumLength
            || maximumLength < 0 || declaredLength > int.MaxValue) {
            throw new InvalidDataException(
                "The archive entry has an invalid declared length.");
        }

        using var output = new MemoryStream(checked((int)declaredLength));
        var buffer = new byte[81920];
        long remaining = declaredLength;
        while (remaining > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int requested = (int)Math.Min(buffer.Length, remaining);
            int read = source.Read(buffer, 0, requested);
            if (read <= 0 || read > requested) {
                throw new InvalidDataException(
                    "The archive entry is shorter than its declared length.");
            }
            output.Write(buffer, 0, read);
            remaining -= read;
        }

        cancellationToken.ThrowIfCancellationRequested();
        if (source.Read(buffer, 0, 1) != 0) {
            throw new InvalidDataException(
                "The archive entry exceeds its declared expansion length.");
        }
        return output.ToArray();
    }

    /// <summary>
    /// Scans ZIP central-directory records without materializing entry metadata.
    /// </summary>
    internal static ZipCentralDirectoryScanResult ScanZipCentralDirectory(
        byte[] bytes, int entryLimit) => ScanZipCentralDirectory(bytes, 0,
        bytes?.Length ?? 0, entryLimit);

    /// <summary>
    /// Scans ZIP central-directory records from a seekable stream using bounded buffers.
    /// The source position is restored before this method returns.
    /// </summary>
    internal static ZipCentralDirectoryScanResult ScanZipCentralDirectory(
        Stream source, long archiveLength, int entryLimit,
        CancellationToken cancellationToken = default) =>
        ScanZipCentralDirectory(source, archiveLength, entryLimit, cancellationToken, null);

    /// <summary>
    /// Scans like <see cref="ScanZipCentralDirectory(Stream, long, int, CancellationToken)"/>, appends each file header
    /// in directory order, and also rejects layouts that another ZIP reader could resolve to a different directory:
    /// a later end-of-central-directory signature, ZIP64 values that disagree with unmasked classic fields, and file
    /// headers continuing past the declared directory size. Extractors use it so these records describe exactly the
    /// entries that <see cref="ZipArchive"/> opens.
    /// </summary>
    internal static ZipCentralDirectoryScanResult ScanZipCentralDirectoryForExtraction(
        Stream source, long archiveLength, int entryLimit, ICollection<ZipCentralDirectoryRecord> records,
        CancellationToken cancellationToken = default) {
        if (records == null) throw new ArgumentNullException(nameof(records));
        return ScanZipCentralDirectory(source, archiveLength, entryLimit, cancellationToken, records);
    }

    private static ZipCentralDirectoryScanResult ScanZipCentralDirectory(
        Stream source, long archiveLength, int entryLimit,
        CancellationToken cancellationToken, ICollection<ZipCentralDirectoryRecord>? records) {
        bool strict = records != null;
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek) throw new ArgumentException(
            "ZIP central-directory inspection requires a readable seekable stream.", nameof(source));
        if (archiveLength < 0 || archiveLength > source.Length - source.Position) {
            throw new ArgumentOutOfRangeException(nameof(archiveLength));
        }
        if (entryLimit < 0) throw new ArgumentOutOfRangeException(nameof(entryLimit));

        long originalPosition = source.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (archiveLength < 22) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "The ZIP end-of-central-directory record was not found.");
            }

            int tailLength = checked((int)Math.Min(archiveLength, 22L + ushort.MaxValue));
            var tail = new byte[tailLength];
            ReadExactlyAt(source, originalPosition + archiveLength - tailLength, tail, tailLength);
            int endOfCentralDirectory = FindEndOfCentralDirectory(tail, 0, tail.Length);
            if (endOfCentralDirectory < 0) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "The ZIP end-of-central-directory record was not found.");
            }
            // Readers that take the last signature without checking the comment length would select a later record.
            if (strict && HasLaterEndOfCentralDirectorySignature(tail, endOfCentralDirectory)) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "The ZIP comment contains another end-of-central-directory signature.");
            }

            long endOfCentralDirectoryOffset = archiveLength - tailLength + endOfCentralDirectory;
            ushort diskNumber = ReadZipUInt16(tail, endOfCentralDirectory + 4);
            ushort centralDirectoryDisk = ReadZipUInt16(tail, endOfCentralDirectory + 6);
            ushort entriesOnDisk16 = ReadZipUInt16(tail, endOfCentralDirectory + 8);
            ushort totalEntries16 = ReadZipUInt16(tail, endOfCentralDirectory + 10);
            uint centralDirectorySize32 = ReadZipUInt32(tail, endOfCentralDirectory + 12);
            uint centralDirectoryOffset32 = ReadZipUInt32(tail, endOfCentralDirectory + 16);

            ulong declaredEntries = 0;
            ulong centralDirectorySize = 0;
            ulong centralDirectoryOffset = 0;
            bool zip64Required = centralDirectorySize32 == uint.MaxValue
                || centralDirectoryOffset32 == uint.MaxValue;
            bool zip64 = false;
            bool hasZip64Locator = HasZip64Locator(source,
                originalPosition, endOfCentralDirectoryOffset);
            // .NET Framework searches backwards through the whole file for a locator when a classic field is masked,
            // so a masked layout must keep its locator in the standard position immediately before this record.
            if (strict && !hasZip64Locator && (zip64Required || totalEntries16 == ushort.MaxValue
                || entriesOnDisk16 == ushort.MaxValue || diskNumber == ushort.MaxValue
                || centralDirectoryDisk == ushort.MaxValue)) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "A masked ZIP end-of-central-directory field requires a ZIP64 locator immediately before it.");
            }
            if (zip64Required || hasZip64Locator) {
                bool parsedZip64 = TryReadZip64Directory(source, originalPosition, archiveLength,
                    endOfCentralDirectoryOffset,
                    out uint zip64DiskNumber,
                    out uint zip64CentralDirectoryDisk,
                    out ulong zip64EntriesOnDisk,
                    out ulong zip64TotalEntries,
                    out ulong zip64CentralDirectorySize,
                    out ulong zip64CentralDirectoryOffset,
                    out string? zip64Error);
                bool supportedZip64 = parsedZip64
                    && zip64DiskNumber == 0 && zip64CentralDirectoryDisk == 0
                    && zip64EntriesOnDisk == zip64TotalEntries;
                if (!supportedZip64 && (zip64Required || strict)) {
                    return ZipCentralDirectoryScanResult.Invalid(zip64Error
                        ?? "The ZIP64 central directory is malformed.");
                }
                // Readers that consult ZIP64 data only for masked classic fields must reach the same directory.
                if (strict && (totalEntries16 != ushort.MaxValue && totalEntries16 != zip64TotalEntries
                    || entriesOnDisk16 != ushort.MaxValue && entriesOnDisk16 != zip64EntriesOnDisk
                    || centralDirectorySize32 != uint.MaxValue && centralDirectorySize32 != zip64CentralDirectorySize
                    || centralDirectoryOffset32 != uint.MaxValue && centralDirectoryOffset32 != zip64CentralDirectoryOffset
                    || diskNumber != ushort.MaxValue && diskNumber != 0
                    || centralDirectoryDisk != ushort.MaxValue && centralDirectoryDisk != 0)) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP64 directory disagrees with the classic end-of-central-directory fields.");
                }
                if (supportedZip64) {
                    zip64 = true;
                    declaredEntries = zip64TotalEntries;
                    centralDirectorySize = zip64CentralDirectorySize;
                    centralDirectoryOffset = zip64CentralDirectoryOffset;
                }
            }
            if (!zip64) {
                if (diskNumber != 0 || centralDirectoryDisk != 0
                    || entriesOnDisk16 != totalEntries16) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "Multi-disk ZIP packages are not supported.");
                }
                declaredEntries = totalEntries16;
                centralDirectorySize = centralDirectorySize32;
                centralDirectoryOffset = centralDirectoryOffset32;
            }

            if (declaredEntries > (ulong)entryLimit) {
                return ZipCentralDirectoryScanResult.Exceeded(checked((long)
                    Math.Min(declaredEntries, (ulong)long.MaxValue)));
            }
            if (centralDirectoryOffset > (ulong)archiveLength
                || centralDirectorySize > (ulong)archiveLength - centralDirectoryOffset) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "The ZIP central-directory bounds exceed the package.");
            }

            long cursor = checked((long)centralDirectoryOffset);
            long end = checked(cursor + (long)centralDirectorySize);
            long actualEntries = 0;
            bool foundDigitalSignature = false;
            var header = new byte[46];
            while (cursor < end) {
                cancellationToken.ThrowIfCancellationRequested();
                if (cursor > end - 4) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP central directory ends inside a record signature.");
                }
                ReadExactlyAt(source, originalPosition + cursor, header, 4);
                uint signature = ReadZipUInt32(header, 0);
                if (signature == CentralDirectoryFileHeaderSignature) {
                    if (foundDigitalSignature) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "The ZIP central directory contains a file header after its digital-signature record.");
                    }
                    if (cursor > end - 46) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "A ZIP central-directory file header is truncated.");
                    }
                    ReadExactlyAt(source, originalPosition + cursor, header, header.Length);
                    long recordLength = 46L
                        + ReadZipUInt16(header, 28)
                        + ReadZipUInt16(header, 30)
                        + ReadZipUInt16(header, 32);
                    if (recordLength > end - cursor) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "A ZIP central-directory file header exceeds the declared directory bounds.");
                    }
                    actualEntries++;
                    if (actualEntries > entryLimit) {
                        return ZipCentralDirectoryScanResult.Exceeded(actualEntries);
                    }
                    if (records != null) {
                        var name = new byte[ReadZipUInt16(header, 28)];
                        ReadExactlyAt(source, originalPosition + cursor + 46, name, name.Length);
                        records.Add(new ZipCentralDirectoryRecord(ReadZipUInt16(header, 8),
                            ReadZipUInt16(header, 10), ReadZipUInt32(header, 16),
                            ReadZipUInt32(header, 20), ReadZipUInt32(header, 24), name));
                    }
                    cursor += recordLength;
                } else if (signature == CentralDirectoryDigitalSignature) {
                    if (foundDigitalSignature) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "The ZIP central directory contains more than one digital-signature record.");
                    }
                    if (cursor > end - 6) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "The ZIP central-directory digital signature is truncated.");
                    }
                    ReadExactlyAt(source, originalPosition + cursor, header, 6);
                    long recordLength = 6L + ReadZipUInt16(header, 4);
                    if (recordLength > end - cursor) {
                        return ZipCentralDirectoryScanResult.Invalid(
                            "The ZIP central-directory digital signature exceeds the declared directory bounds.");
                    }
                    cursor += recordLength;
                    foundDigitalSignature = true;
                } else {
                    return ZipCentralDirectoryScanResult.Invalid(
                        $"The ZIP central directory contains unexpected signature 0x{signature:X8}.");
                }
            }

            if (actualEntries != checked((long)declaredEntries)) {
                return ZipCentralDirectoryScanResult.Invalid(
                    $"The ZIP central directory declares {declaredEntries} entries but contains {actualEntries}.");
            }
            // Readers that stop at the first non-file-header signature must not find more file headers here.
            if (strict && end <= archiveLength - 4) {
                ReadExactlyAt(source, originalPosition + end, header, 4);
                if (ReadZipUInt32(header, 0) == CentralDirectoryFileHeaderSignature) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "ZIP central-directory file headers continue beyond the declared directory size.");
                }
            }
            return ZipCentralDirectoryScanResult.Valid(actualEntries);
        } finally {
            source.Position = originalPosition;
        }
    }

    // Any later signature counts: .NET 10 also considers positions too close to the end for a complete record,
    // so such bytes would make readers disagree or fail on some targets. Rejecting them keeps every target consistent.
    private static bool HasLaterEndOfCentralDirectorySignature(byte[] tail, int endOfCentralDirectory) {
        for (int offset = endOfCentralDirectory + 1; offset <= tail.Length - 4; offset++) {
            if (ReadZipUInt32(tail, offset) == EndOfCentralDirectorySignature) return true;
        }
        return false;
    }

    /// <summary>
    /// Scans a ZIP slice without copying it or materializing entry metadata.
    /// </summary>
    internal static ZipCentralDirectoryScanResult ScanZipCentralDirectory(
        byte[] bytes, int archiveOffset, int archiveLength,
        int entryLimit) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (archiveOffset < 0 || archiveLength < 0
            || archiveOffset > bytes.Length - archiveLength) {
            throw new ArgumentOutOfRangeException(nameof(archiveOffset));
        }
        if (entryLimit < 0) throw new ArgumentOutOfRangeException(nameof(entryLimit));
        int endOfCentralDirectory = FindEndOfCentralDirectory(bytes,
            archiveOffset, archiveLength);
        if (endOfCentralDirectory < 0) {
            return ZipCentralDirectoryScanResult.Invalid(
                "The ZIP end-of-central-directory record was not found.");
        }

        ushort diskNumber = ReadZipUInt16(bytes, endOfCentralDirectory + 4);
        ushort centralDirectoryDisk = ReadZipUInt16(bytes,
            endOfCentralDirectory + 6);
        ushort entriesOnDisk16 = ReadZipUInt16(bytes,
            endOfCentralDirectory + 8);
        ushort totalEntries16 = ReadZipUInt16(bytes,
            endOfCentralDirectory + 10);
        uint centralDirectorySize32 = ReadZipUInt32(bytes,
            endOfCentralDirectory + 12);
        uint centralDirectoryOffset32 = ReadZipUInt32(bytes,
            endOfCentralDirectory + 16);

        ulong declaredEntries = 0;
        ulong centralDirectorySize = 0;
        ulong centralDirectoryOffset = 0;
        bool zip64Required = centralDirectorySize32 == uint.MaxValue
            || centralDirectoryOffset32 == uint.MaxValue;
        bool zip64 = false;
        bool hasZip64Locator = HasZip64Locator(bytes,
            archiveOffset, endOfCentralDirectory);
        if (zip64Required || hasZip64Locator) {
            bool parsedZip64 = TryReadZip64Directory(bytes, archiveOffset,
                archiveLength, endOfCentralDirectory,
                out uint zip64DiskNumber,
                out uint zip64CentralDirectoryDisk,
                out ulong zip64EntriesOnDisk,
                out ulong zip64TotalEntries,
                out ulong zip64CentralDirectorySize,
                out ulong zip64CentralDirectoryOffset,
                out string? zip64Error);
            bool supportedZip64 = parsedZip64
                && zip64DiskNumber == 0 && zip64CentralDirectoryDisk == 0
                && zip64EntriesOnDisk == zip64TotalEntries;
            if (!supportedZip64 && zip64Required) {
                return ZipCentralDirectoryScanResult.Invalid(zip64Error
                    ?? "The ZIP64 central directory is malformed.");
            }
            if (supportedZip64) {
                zip64 = true;
                declaredEntries = zip64TotalEntries;
                centralDirectorySize = zip64CentralDirectorySize;
                centralDirectoryOffset = zip64CentralDirectoryOffset;
            }
        }
        if (!zip64) {
            if (diskNumber != 0 || centralDirectoryDisk != 0
                || entriesOnDisk16 != totalEntries16) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "Multi-disk ZIP packages are not supported.");
            }
            declaredEntries = totalEntries16;
            centralDirectorySize = centralDirectorySize32;
            centralDirectoryOffset = centralDirectoryOffset32;
        }

        if (declaredEntries > (ulong)entryLimit) {
            return ZipCentralDirectoryScanResult.Exceeded(checked((long)
                Math.Min(declaredEntries, (ulong)long.MaxValue)));
        }
        if (centralDirectoryOffset > (ulong)archiveLength
            || centralDirectorySize > (ulong)archiveLength
                - centralDirectoryOffset) {
            return ZipCentralDirectoryScanResult.Invalid(
                "The ZIP central-directory bounds exceed the package.");
        }

        long cursor = checked(archiveOffset
            + (long)centralDirectoryOffset);
        long end = checked(cursor + (long)centralDirectorySize);
        long actualEntries = 0;
        bool foundDigitalSignature = false;
        while (cursor < end) {
            if (cursor > end - 4) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "The ZIP central directory ends inside a record signature.");
            }

            uint signature = ReadZipUInt32(bytes, checked((int)cursor));
            if (signature == CentralDirectoryFileHeaderSignature) {
                if (foundDigitalSignature) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP central directory contains a file header after its digital-signature record.");
                }
                if (cursor > end - 46) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "A ZIP central-directory file header is truncated.");
                }
                int headerOffset = checked((int)cursor);
                long recordLength = 46L
                    + ReadZipUInt16(bytes, headerOffset + 28)
                    + ReadZipUInt16(bytes, headerOffset + 30)
                    + ReadZipUInt16(bytes, headerOffset + 32);
                if (recordLength > end - cursor) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "A ZIP central-directory file header exceeds the declared directory bounds.");
                }
                cursor += recordLength;
                actualEntries++;
                if (actualEntries > entryLimit) {
                    return ZipCentralDirectoryScanResult.Exceeded(
                        actualEntries);
                }
            } else if (signature == CentralDirectoryDigitalSignature) {
                if (foundDigitalSignature) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP central directory contains more than one digital-signature record.");
                }
                if (cursor > end - 6) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP central-directory digital signature is truncated.");
                }
                long recordLength = 6L + ReadZipUInt16(bytes,
                    checked((int)cursor + 4));
                if (recordLength > end - cursor) {
                    return ZipCentralDirectoryScanResult.Invalid(
                        "The ZIP central-directory digital signature exceeds the declared directory bounds.");
                }
                cursor += recordLength;
                foundDigitalSignature = true;
            } else {
                return ZipCentralDirectoryScanResult.Invalid(
                    $"The ZIP central directory contains unexpected signature 0x{signature:X8}.");
            }
        }

        if (actualEntries != checked((long)declaredEntries)) {
            return ZipCentralDirectoryScanResult.Invalid(
                $"The ZIP central directory declares {declaredEntries} entries but contains {actualEntries}.");
        }
        return ZipCentralDirectoryScanResult.Valid(actualEntries);
    }

    private static int FindEndOfCentralDirectory(byte[] bytes,
        int archiveOffset, int archiveLength) {
        if (archiveLength < 22) return -1;
        int archiveEnd = checked(archiveOffset + archiveLength);
        int minimumOffset = Math.Max(archiveOffset,
            archiveEnd - (22 + ushort.MaxValue));
        for (int offset = archiveEnd - 22;
             offset >= minimumOffset; offset--) {
            if (ReadZipUInt32(bytes, offset)
                != EndOfCentralDirectorySignature) continue;
            ushort commentLength = ReadZipUInt16(bytes, offset + 20);
            if ((long)offset + 22L + commentLength == archiveEnd) {
                return offset;
            }
        }
        return -1;
    }

    private static bool TryReadZip64Directory(byte[] bytes,
        int archiveOffset, int archiveLength, int endOfCentralDirectory,
        out uint diskNumber,
        out uint centralDirectoryDisk, out ulong entriesOnDisk,
        out ulong totalEntries, out ulong centralDirectorySize,
        out ulong centralDirectoryOffset, out string? error) {
        diskNumber = centralDirectoryDisk = 0;
        entriesOnDisk = totalEntries = centralDirectorySize =
            centralDirectoryOffset = 0;
        error = null;
        int locatorOffset = endOfCentralDirectory - 20;
        if (locatorOffset < archiveOffset
            || ReadZipUInt32(bytes, locatorOffset)
            != Zip64EndOfCentralDirectoryLocatorSignature) {
            error = "The ZIP64 end-of-central-directory locator was not found.";
            return false;
        }
        if (ReadZipUInt32(bytes, locatorOffset + 4) != 0
            || ReadZipUInt32(bytes, locatorOffset + 16) != 1) {
            error = "Multi-disk ZIP64 packages are not supported.";
            return false;
        }

        ulong zip64Offset = ReadZipUInt64(bytes, locatorOffset + 8);
        if (archiveLength < 56
            || zip64Offset > (ulong)archiveLength - 56UL) {
            error = "The ZIP64 end-of-central-directory record exceeds the package bounds.";
            return false;
        }
        int recordOffset = checked(archiveOffset + (int)zip64Offset);
        if (ReadZipUInt32(bytes, recordOffset)
            != Zip64EndOfCentralDirectorySignature) {
            error = "The ZIP64 end-of-central-directory record was not found at its declared offset.";
            return false;
        }
        ulong recordSize = ReadZipUInt64(bytes, recordOffset + 4);
        if (recordSize < 44UL || recordSize > (ulong)archiveLength
            - zip64Offset - 12UL) {
            error = "The ZIP64 end-of-central-directory record has invalid bounds.";
            return false;
        }

        diskNumber = ReadZipUInt32(bytes, recordOffset + 16);
        centralDirectoryDisk = ReadZipUInt32(bytes, recordOffset + 20);
        entriesOnDisk = ReadZipUInt64(bytes, recordOffset + 24);
        totalEntries = ReadZipUInt64(bytes, recordOffset + 32);
        centralDirectorySize = ReadZipUInt64(bytes, recordOffset + 40);
        centralDirectoryOffset = ReadZipUInt64(bytes, recordOffset + 48);
        ulong locatorOffsetRelative = checked((ulong)(endOfCentralDirectory - archiveOffset - 20));
        ulong recordEnd = zip64Offset + 12UL + recordSize;
        if (recordEnd != locatorOffsetRelative
            || centralDirectoryOffset > zip64Offset
            || centralDirectorySize > zip64Offset - centralDirectoryOffset) {
            error = "The ZIP64 central-directory structure is inconsistent.";
            return false;
        }
        return true;
    }

    private static bool HasZip64Locator(byte[] bytes, int archiveOffset,
        int endOfCentralDirectory) {
        int locatorOffset = endOfCentralDirectory - 20;
        return locatorOffset >= archiveOffset
            && ReadZipUInt32(bytes, locatorOffset)
                == Zip64EndOfCentralDirectoryLocatorSignature;
    }

    private static bool HasZip64Locator(Stream source, long archiveOffset,
        long endOfCentralDirectory) {
        long locatorOffset = endOfCentralDirectory - 20;
        if (locatorOffset < 0) return false;
        var signature = new byte[4];
        ReadExactlyAt(source, archiveOffset + locatorOffset,
            signature, signature.Length);
        return ReadZipUInt32(signature, 0)
            == Zip64EndOfCentralDirectoryLocatorSignature;
    }

    private static bool TryReadZip64Directory(Stream source,
        long archiveOffset, long archiveLength, long endOfCentralDirectory,
        out uint diskNumber,
        out uint centralDirectoryDisk, out ulong entriesOnDisk,
        out ulong totalEntries, out ulong centralDirectorySize,
        out ulong centralDirectoryOffset, out string? error) {
        diskNumber = centralDirectoryDisk = 0;
        entriesOnDisk = totalEntries = centralDirectorySize = centralDirectoryOffset = 0;
        error = null;
        long locatorOffset = endOfCentralDirectory - 20;
        if (locatorOffset < 0) {
            error = "The ZIP64 end-of-central-directory locator was not found.";
            return false;
        }
        var locator = new byte[20];
        ReadExactlyAt(source, archiveOffset + locatorOffset, locator, locator.Length);
        if (ReadZipUInt32(locator, 0) != Zip64EndOfCentralDirectoryLocatorSignature) {
            error = "The ZIP64 end-of-central-directory locator was not found.";
            return false;
        }
        if (ReadZipUInt32(locator, 4) != 0 || ReadZipUInt32(locator, 16) != 1) {
            error = "Multi-disk ZIP64 packages are not supported.";
            return false;
        }

        ulong zip64Offset = ReadZipUInt64(locator, 8);
        if (archiveLength < 56 || zip64Offset > (ulong)archiveLength - 56UL) {
            error = "The ZIP64 end-of-central-directory record exceeds the package bounds.";
            return false;
        }
        var record = new byte[56];
        ReadExactlyAt(source, checked(archiveOffset + (long)zip64Offset), record, record.Length);
        if (ReadZipUInt32(record, 0) != Zip64EndOfCentralDirectorySignature) {
            error = "The ZIP64 end-of-central-directory record was not found at its declared offset.";
            return false;
        }
        ulong recordSize = ReadZipUInt64(record, 4);
        if (recordSize < 44UL || recordSize > (ulong)archiveLength - zip64Offset - 12UL) {
            error = "The ZIP64 end-of-central-directory record has invalid bounds.";
            return false;
        }

        diskNumber = ReadZipUInt32(record, 16);
        centralDirectoryDisk = ReadZipUInt32(record, 20);
        entriesOnDisk = ReadZipUInt64(record, 24);
        totalEntries = ReadZipUInt64(record, 32);
        centralDirectorySize = ReadZipUInt64(record, 40);
        centralDirectoryOffset = ReadZipUInt64(record, 48);
        ulong locatorOffsetRelative = checked((ulong)(endOfCentralDirectory - 20));
        ulong recordEnd = zip64Offset + 12UL + recordSize;
        if (recordEnd != locatorOffsetRelative
            || centralDirectoryOffset > zip64Offset
            || centralDirectorySize > zip64Offset - centralDirectoryOffset) {
            error = "The ZIP64 central-directory structure is inconsistent.";
            return false;
        }
        return true;
    }

    private static void ReadExactlyAt(Stream source, long position,
        byte[] buffer, int count) {
        source.Position = position;
        int offset = 0;
        while (offset < count) {
            int read = source.Read(buffer, offset, count - offset);
            if (read <= 0) throw new InvalidDataException(
                "The ZIP package ended before its declared metadata.");
            offset += read;
        }
    }

    private static ushort ReadZipUInt16(byte[] data, int offset) =>
        (ushort)(data[offset] | (data[offset + 1] << 8));

    private static uint ReadZipUInt32(byte[] data, int offset) =>
        (uint)(data[offset]
            | (data[offset + 1] << 8)
            | (data[offset + 2] << 16)
            | (data[offset + 3] << 24));

    private static ulong ReadZipUInt64(byte[] data, int offset) =>
        (ulong)ReadZipUInt32(data, offset)
        | (ulong)ReadZipUInt32(data, offset + 4) << 32;

    /// <summary>Fixed fields and raw name bytes of one ZIP central-directory file header.
    /// Size fields are the raw 32-bit values; <see cref="uint.MaxValue"/> defers to ZIP64 data.</summary>
    internal readonly struct ZipCentralDirectoryRecord {
        internal ZipCentralDirectoryRecord(ushort flags, ushort compressionMethod, uint crc32,
            uint compressedSize, uint uncompressedSize, byte[] nameBytes) {
            Flags = flags;
            CompressionMethod = compressionMethod;
            Crc32 = crc32;
            CompressedSize = compressedSize;
            UncompressedSize = uncompressedSize;
            NameBytes = nameBytes;
        }

        internal ushort Flags { get; }
        internal ushort CompressionMethod { get; }
        internal uint Crc32 { get; }
        internal uint CompressedSize { get; }
        internal uint UncompressedSize { get; }
        internal byte[] NameBytes { get; }
    }

    internal readonly struct ZipCentralDirectoryScanResult {
        private ZipCentralDirectoryScanResult(bool isValid,
            bool limitExceeded, long entryCount, string? error) {
            IsValid = isValid;
            LimitExceeded = limitExceeded;
            EntryCount = entryCount;
            Error = error;
        }

        internal bool IsValid { get; }
        internal bool LimitExceeded { get; }
        internal long EntryCount { get; }
        internal string? Error { get; }

        internal static ZipCentralDirectoryScanResult Valid(
            long entryCount) => new(true, false, entryCount, null);

        internal static ZipCentralDirectoryScanResult Exceeded(
            long entryCount) => new(true, true, entryCount, null);

        internal static ZipCentralDirectoryScanResult Invalid(
            string error) => new(false, false, 0, error);
    }

    private static bool ContainsNull(string value) {
        for (int i = 0; i < value.Length; i++) {
            if (value[i] == '\0') return true;
        }
        return false;
    }
}
