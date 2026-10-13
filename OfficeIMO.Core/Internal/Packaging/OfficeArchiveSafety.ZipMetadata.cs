using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    private readonly struct ZipEntryCoordinates(ulong compressed, ulong expanded, ulong offset, uint disk) {
        internal ulong Compressed { get; } = compressed;
        internal ulong Expanded { get; } = expanded;
        internal ulong Offset { get; } = offset;
        internal uint Disk { get; } = disk;
    }

    private static bool TryReadZipEntryCoordinates(byte[] header, int headerOffset, byte[] extra, int extraOffset, int extraLength,
        out ZipEntryCoordinates coordinates) {
        ulong compressed = ReadZipUInt32(header, headerOffset + 20), expanded = ReadZipUInt32(header, headerOffset + 24);
        ulong offset = ReadZipUInt32(header, headerOffset + 42);
        uint disk = ReadZipUInt16(header, headerOffset + 34);
        coordinates = default;
        if (compressed == uint.MaxValue || expanded == uint.MaxValue || offset == uint.MaxValue || disk == ushort.MaxValue) {
            int end = extraOffset + extraLength;
            bool found = false;
            while (extraOffset <= end - 4) {
                ushort id = ReadZipUInt16(extra, extraOffset), length = ReadZipUInt16(extra, extraOffset + 2);
                extraOffset += 4;
                if (length > end - extraOffset) return false;
                if (id != 1) { extraOffset += length; continue; }
                int valueEnd = extraOffset + length;
                if (expanded == uint.MaxValue) { if (extraOffset > valueEnd - 8) return false; expanded = ReadZipUInt64(extra, extraOffset); extraOffset += 8; }
                if (compressed == uint.MaxValue) { if (extraOffset > valueEnd - 8) return false; compressed = ReadZipUInt64(extra, extraOffset); extraOffset += 8; }
                if (offset == uint.MaxValue) { if (extraOffset > valueEnd - 8) return false; offset = ReadZipUInt64(extra, extraOffset); extraOffset += 8; }
                if (disk == ushort.MaxValue) { if (extraOffset > valueEnd - 4) return false; disk = ReadZipUInt32(extra, extraOffset); }
                found = true;
                break;
            }
            if (!found) return false;
        }
        if (compressed > long.MaxValue || expanded > long.MaxValue || offset > long.MaxValue) return false;
        coordinates = new ZipEntryCoordinates(compressed, expanded, offset, disk);
        return true;
    }

    /// <summary>Checks the local header and bounded compressed range even when a member is excluded.</summary>
    internal static long ValidateEntryPayload(Stream source, ZipCentralDirectoryRecord record) {
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
        return payload;
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
