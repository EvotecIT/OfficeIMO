using System;
using System.IO;
using System.IO.Compression;

namespace OfficeIMO.Drawing.Internal;

/// <summary>
/// Centralizes archive-entry safety rules shared by ZIP-backed OfficeIMO format owners.
/// </summary>
internal static class OfficeArchiveSafety {
    private static readonly char[] PathSeparators = { '/' };
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

        string[] segments = fullName.Split(PathSeparators, StringSplitOptions.RemoveEmptyEntries);
        foreach (string segment in segments) {
            if (segment == "." || segment == "..") return true;
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
    internal static byte[] ReadEntryBytes(Stream source,
        long declaredLength, long maximumLength) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (declaredLength < 0 || declaredLength > maximumLength
            || maximumLength < 0 || declaredLength > int.MaxValue) {
            throw new InvalidDataException(
                "The archive entry has an invalid declared length.");
        }

        using var output = new MemoryStream(checked((int)declaredLength));
        var buffer = new byte[81920];
        long remaining = declaredLength;
        while (remaining > 0) {
            int requested = (int)Math.Min(buffer.Length, remaining);
            int read = source.Read(buffer, 0, requested);
            if (read <= 0 || read > requested) {
                throw new InvalidDataException(
                    "The archive entry is shorter than its declared length.");
            }
            output.Write(buffer, 0, read);
            remaining -= read;
        }

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

        ulong declaredEntries;
        ulong centralDirectorySize;
        ulong centralDirectoryOffset;
        bool zip64 = entriesOnDisk16 == ushort.MaxValue
            || totalEntries16 == ushort.MaxValue
            || centralDirectorySize32 == uint.MaxValue
            || centralDirectoryOffset32 == uint.MaxValue;
        if (zip64) {
            if (!TryReadZip64Directory(bytes, archiveOffset,
                    archiveLength, endOfCentralDirectory,
                    out uint zip64DiskNumber,
                    out uint zip64CentralDirectoryDisk,
                    out ulong zip64EntriesOnDisk,
                    out declaredEntries,
                    out centralDirectorySize,
                    out centralDirectoryOffset,
                    out string? zip64Error)) {
                return ZipCentralDirectoryScanResult.Invalid(zip64Error
                    ?? "The ZIP64 central directory is malformed.");
            }
            if (zip64DiskNumber != 0 || zip64CentralDirectoryDisk != 0
                || zip64EntriesOnDisk != declaredEntries) {
                return ZipCentralDirectoryScanResult.Invalid(
                    "Multi-disk ZIP packages are not supported.");
            }
        } else {
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
        return true;
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
