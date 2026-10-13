using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    /// <summary>Supports UTF-8 names and descriptors, plus level hints only for ordinary Deflate.</summary>
    internal static bool HasSupportedZipFlags(ushort flags, ushort method) =>
        (method == 0 || method == 8) && (flags & ~(method == 8 ? 0x080e : 0x0808)) == 0;

    private readonly struct ZipEntryCoordinates(ulong compressed, ulong expanded, ulong offset, uint disk, bool zip64) {
        internal ulong Compressed { get; } = compressed;
        internal ulong Expanded { get; } = expanded;
        internal ulong Offset { get; } = offset;
        internal uint Disk { get; } = disk;
        internal bool Zip64 { get; } = zip64;
    }

    private static bool TryReadZipEntryCoordinates(byte[] header, int headerOffset, byte[] extra, int extraOffset, int extraLength,
        out ZipEntryCoordinates coordinates) => TryResolveZip64Coordinates(
            ReadZipUInt32(header, headerOffset + 20), ReadZipUInt32(header, headerOffset + 24),
            ReadZipUInt32(header, headerOffset + 42), ReadZipUInt16(header, headerOffset + 34), extra, extraOffset, extraLength, out coordinates);

    private static bool TryResolveZip64Coordinates(ulong compressed, ulong expanded, ulong offset, uint disk,
        byte[] extra, int extraOffset, int extraLength, out ZipEntryCoordinates coordinates) {
        coordinates = default;
        if (!TryFindZip64Extra(extra, extraOffset, extraLength, out int zip64Offset, out int zip64Length)) return false;
        if (compressed == uint.MaxValue || expanded == uint.MaxValue || offset == uint.MaxValue || disk == ushort.MaxValue) {
            if (zip64Offset < 0) return false;
            int valueEnd = zip64Offset + zip64Length;
            if (expanded == uint.MaxValue) { if (zip64Offset > valueEnd - 8) return false; expanded = ReadZipUInt64(extra, zip64Offset); zip64Offset += 8; }
            if (compressed == uint.MaxValue) { if (zip64Offset > valueEnd - 8) return false; compressed = ReadZipUInt64(extra, zip64Offset); zip64Offset += 8; }
            if (offset == uint.MaxValue) { if (zip64Offset > valueEnd - 8) return false; offset = ReadZipUInt64(extra, zip64Offset); zip64Offset += 8; }
            if (disk == ushort.MaxValue) { if (zip64Offset > valueEnd - 4) return false; disk = ReadZipUInt32(extra, zip64Offset); }
        }
        if (compressed > long.MaxValue || expanded > long.MaxValue || offset > long.MaxValue) return false;
        coordinates = new ZipEntryCoordinates(compressed, expanded, offset, disk, zip64Offset >= 0);
        return true;
    }

    private static bool TryFindZip64Extra(byte[] extra, int offset, int length, out int zip64Offset, out int zip64Length) {
        zip64Offset = -1; zip64Length = 0;
        int end = offset + length;
        while (offset < end) {
            if (offset > end - 4) return false;
            ushort id = ReadZipUInt16(extra, offset), size = ReadZipUInt16(extra, offset + 2);
            offset += 4;
            if (size > end - offset) return false;
            if (id == 1) {
                if (zip64Offset >= 0) return false;
                zip64Offset = offset; zip64Length = size;
            }
            offset += size;
        }
        return true;
    }

    /// <summary>Checks the local header and bounded compressed range even when a member is excluded.</summary>
    internal static long ValidateEntryPayload(Stream source, ZipCentralDirectoryRecord record) {
        if (!HasSupportedZipFlags(record.Flags, record.CompressionMethod))
            throw new InvalidDataException("ZIP entry uses unsupported compression flags.");
        long offset = record.LocalHeaderOffset;
        long bound = record.PayloadUpperBound;
        if (offset < 0 || offset > bound - 30 || bound > source.Length)
            throw new InvalidDataException("ZIP local header is outside its package bounds.");
        var header = new byte[30];
        ReadExactlyAt(source, offset, header, header.Length);
        if (ReadZipUInt32(header, 0) != 0x04034b50U || ReadZipUInt16(header, 6) != record.Flags
            || ReadZipUInt16(header, 8) != record.CompressionMethod)
            throw new InvalidDataException("ZIP local and central compression metadata disagree.");
        int nameLength = ReadZipUInt16(header, 26);
        long payload = offset + 30L + nameLength + ReadZipUInt16(header, 28);
        if (payload > bound || record.CompressedSize > bound - payload || nameLength != record.NameBytes.Length)
            throw new InvalidDataException("ZIP compressed payload or name is outside its package bounds.");
        var name = new byte[nameLength];
        ReadExactlyAt(source, offset + 30, name, name.Length);
        for (int index = 0; index < name.Length; index++)
            if (name[index] != record.NameBytes[index]) throw new InvalidDataException("ZIP local and central names disagree.");
        ValidateLocalIntegrity(source, header, offset, payload, bound, record.Crc32, record.CompressedSize, record.UncompressedSize, record.UsesZip64);
        return payload;
    }

    private static void ValidateLocalIntegrity(Stream source, byte[] header, long offset, long payload, long bound,
        uint checksum, long compressedSize, long expandedSize, bool centralZip64) {
        var extra = new byte[ReadZipUInt16(header, 28)];
        ReadExactlyAt(source, offset + 30 + ReadZipUInt16(header, 26), extra, extra.Length);
        if (!TryFindZip64Extra(extra, 0, extra.Length, out int zip64Offset, out _))
            throw new InvalidDataException("ZIP local extra fields are malformed or repeat ZIP64 metadata.");
        if ((ReadZipUInt16(header, 6) & 8) == 0) {
            uint compressed = ReadZipUInt32(header, 18), expanded = ReadZipUInt32(header, 22);
            if (ReadZipUInt32(header, 14) != checksum
                || !TryResolveZip64Coordinates(compressed, expanded, 0, 0, extra, 0, extra.Length, out var local)
                || local.Compressed != (ulong)compressedSize || local.Expanded != (ulong)expandedSize)
                throw new InvalidDataException("ZIP local and central checksums or sizes disagree.");
        } else ValidateZipDataDescriptor(source, payload + compressedSize, bound, checksum, compressedSize, expandedSize, centralZip64 || zip64Offset >= 0);
    }

    /// <summary>Opens the bounded raw compressed range without a declared-size decoded-output clamp.</summary>
    internal static Stream OpenEntryPayload(Stream source, ZipCentralDirectoryRecord record, CancellationToken cancellationToken = default) {
        long payload = ValidateEntryPayload(source, record);
        var compressed = new SeekableReadWindowStream(source, payload, record.CompressedSize);
        if (record.CompressionMethod == 0) return compressed;
        if (record.CompressionMethod == 8) {
            try {
                if (!OfficeDeflateStreamValidator.TryValidateExact(compressed, record.UncompressedSize, cancellationToken))
                    throw new InvalidDataException("ZIP member has invalid or incomplete Deflate framing.");
                return new DeflateStream(compressed, CompressionMode.Decompress);
            } catch {
                compressed.Dispose();
                throw;
            }
        }
        compressed.Dispose();
        throw new NotSupportedException("ZIP members must be stored or Deflate-compressed.");
    }
}
