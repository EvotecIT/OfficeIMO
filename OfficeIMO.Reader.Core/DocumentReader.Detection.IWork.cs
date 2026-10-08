using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    // Content detection may inspect an untrusted ZIP without a registered iWork handler.
    private const int MaximumNestedIWorkIndexProbeBytes = 16 * 1024 * 1024;

    private static bool TryInspectNestedIWorkIndex(Stream stream, long archiveStart,
        long localHeaderOffset, ushort compression, uint compressedSize, uint uncompressedSize,
        int maxEntries) {
        long returnPosition = stream.Position;
        try {
            if (!TryLocateNestedIndexPayload(stream, archiveStart, localHeaderOffset,
                    compression, compressedSize, uncompressedSize, out int length)) return false;
            var payload = new byte[length];
            return ReadExact(stream, payload, 0, length) &&
                   HasNestedIWorkDocument(payload, compression, uncompressedSize,
                       maxEntries, CancellationToken.None);
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } finally {
            stream.Position = returnPosition;
        }
    }

    private static async Task<bool> TryInspectNestedIWorkIndexAsync(Stream stream,
        long archiveStart, long localHeaderOffset, ushort compression,
        uint compressedSize, uint uncompressedSize, int maxEntries,
        CancellationToken cancellationToken) {
        long returnPosition = stream.Position;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (!await TryLocateNestedIndexPayloadAsync(stream, archiveStart,
                    localHeaderOffset, compression, compressedSize, uncompressedSize,
                    cancellationToken).ConfigureAwait(false)) return false;
            var payload = new byte[(int)compressedSize];
            return await ReadExactAsync(stream, payload, 0, payload.Length,
                       cancellationToken).ConfigureAwait(false) &&
                   HasNestedIWorkDocument(payload, compression, uncompressedSize,
                       maxEntries, cancellationToken);
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } finally {
            stream.Position = returnPosition;
        }
    }

    private static bool TryLocateNestedIndexPayload(Stream stream, long archiveStart,
        long localHeaderOffset, ushort compression, uint compressedSize,
        uint uncompressedSize, out int length) {
        length = 0;
        if (!NestedIndexSizeIsBounded(compression, compressedSize, uncompressedSize) ||
            localHeaderOffset < 0 || localHeaderOffset > stream.Length - archiveStart - 30) return false;
        stream.Position = archiveStart + localHeaderOffset;
        var header = new byte[30];
        if (!ReadExact(stream, header, 0, header.Length) ||
            !NestedIndexLocalHeaderMatches(header, compression)) return false;
        int nameLength = ReadUInt16(header, 26);
        int extraLength = ReadUInt16(header, 28);
        if (nameLength == 0 || nameLength > 4096 ||
            nameLength > stream.Length - stream.Position) return false;
        var name = new byte[nameLength];
        if (!ReadExact(stream, name, 0, name.Length) ||
            NormalizeZipEntryName(name) != "index.zip") return false;
        if (extraLength > stream.Length - stream.Position - compressedSize) return false;
        stream.Position += extraLength;
        length = (int)compressedSize;
        return true;
    }

    private static async Task<bool> TryLocateNestedIndexPayloadAsync(Stream stream,
        long archiveStart, long localHeaderOffset, ushort compression,
        uint compressedSize, uint uncompressedSize, CancellationToken cancellationToken) {
        if (!NestedIndexSizeIsBounded(compression, compressedSize, uncompressedSize) ||
            localHeaderOffset < 0 || localHeaderOffset > stream.Length - archiveStart - 30) return false;
        stream.Position = archiveStart + localHeaderOffset;
        var header = new byte[30];
        if (!await ReadExactAsync(stream, header, 0, header.Length,
                cancellationToken).ConfigureAwait(false) ||
            !NestedIndexLocalHeaderMatches(header, compression)) return false;
        int nameLength = ReadUInt16(header, 26);
        int extraLength = ReadUInt16(header, 28);
        if (nameLength == 0 || nameLength > 4096 ||
            nameLength > stream.Length - stream.Position) return false;
        var name = new byte[nameLength];
        if (!await ReadExactAsync(stream, name, 0, name.Length,
                cancellationToken).ConfigureAwait(false) ||
            NormalizeZipEntryName(name) != "index.zip") return false;
        if (extraLength > stream.Length - stream.Position - compressedSize) return false;
        stream.Position += extraLength;
        return true;
    }

    private static bool NestedIndexSizeIsBounded(ushort compression, uint compressedSize,
        uint uncompressedSize) => (compression == 0 || compression == 8) &&
        compressedSize > 0 && uncompressedSize > 0 &&
        compressedSize <= MaximumNestedIWorkIndexProbeBytes &&
        uncompressedSize <= MaximumNestedIWorkIndexProbeBytes;

    private static bool NestedIndexLocalHeaderMatches(byte[] header, ushort compression) =>
        ReadUInt32(header, 0) == ZipLocalHeaderSignature &&
        ReadUInt16(header, 8) == compression &&
        (ReadUInt16(header, 6) & 1) == 0;

    private static bool HasNestedIWorkDocument(byte[] payload, ushort compression,
        uint expectedLength, int maxEntries, CancellationToken cancellationToken) {
        byte[]? nestedBytes = InflateBoundedContainerEntry(payload, compression, expectedLength,
            cancellationToken);
        if (nestedBytes == null) return false;
        using var nested = new MemoryStream(nestedBytes, writable: false);
        if (!TryLocateZipCentralDirectory(nested, 0, out long directoryOffset,
                out int entryCount)) return false;
        nested.Position = directoryOffset;
        var header = new byte[ZipCentralDirectoryHeaderLength];
        for (int index = 0; index < Math.Min(entryCount, maxEntries); index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ReadExact(nested, header, 0, header.Length) ||
                ReadUInt32(header, 0) != ZipCentralDirectoryHeaderSignature) return false;
            int nameLength = ReadUInt16(header, 28);
            int extraLength = ReadUInt16(header, 30);
            int commentLength = ReadUInt16(header, 32);
            if (nameLength == 0 || nameLength > 4096 ||
                nameLength > nested.Length - nested.Position) return false;
            var name = new byte[nameLength];
            if (!ReadExact(nested, name, 0, name.Length)) return false;
            long next = nested.Position + extraLength + commentLength;
            if (next < nested.Position || next > nested.Length) return false;
            if (NormalizeZipEntryName(name) is "document.iwa" or "index/document.iwa" &&
                ReadUInt32(header, 24) > 0) return true;
            nested.Position = next;
        }
        return false;
    }

    private static byte[]? InflateBoundedContainerEntry(byte[] payload, ushort compression,
        uint expectedLength, CancellationToken cancellationToken) {
        if (compression == 0) return payload.Length == expectedLength ? payload : null;
        using var source = new MemoryStream(payload, writable: false);
        using var inflater = new DeflateStream(source, CompressionMode.Decompress);
        using var output = new MemoryStream();
        var buffer = new byte[8192];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = inflater.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (output.Length > expectedLength - read) return null;
            output.Write(buffer, 0, read);
        }
        return output.Length == expectedLength ? output.ToArray() : null;
    }
}
