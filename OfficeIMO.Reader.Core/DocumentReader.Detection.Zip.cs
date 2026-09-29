using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const uint ZipCentralDirectoryHeaderSignature = 0x02014B50U;
    private const uint ZipEndOfCentralDirectorySignature = 0x06054B50U;
    private const uint Zip64EndOfCentralDirectorySignature = 0x06064B50U;
    private const uint Zip64EndOfCentralDirectoryLocatorSignature = 0x07064B50U;
    private const uint ZipLocalHeaderSignature = 0x04034B50U;
    private const int ZipCentralDirectoryHeaderLength = 46;
    private const int ZipEndOfCentralDirectoryLength = 22;
    private const int ZipMaximumEndRecordLength = ZipEndOfCentralDirectoryLength + ushort.MaxValue;

    private static DetectionCandidate InspectZipContainer(Stream stream, long start, int maxEntries) {
        if (!TryLocateZipCentralDirectory(stream, start, out long centralDirectoryOffset, out int entryCount)) {
            return GenericZipCandidate();
        }

        stream.Position = start + centralDirectoryOffset;
        var header = new byte[ZipCentralDirectoryHeaderLength];
        int entriesToInspect = Math.Min(entryCount, maxEntries);
        bool hasIWorkIndex = false;
        (long Offset, ushort Compression, uint CompressedSize, uint UncompressedSize)? nestedIndex = null;
        for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
            if (!ReadExact(stream, header, 0, header.Length) ||
                ReadUInt32(header, 0) != ZipCentralDirectoryHeaderSignature) {
                break;
            }

            ushort compression = ReadUInt16(header, 10);
            uint compressedSize = ReadUInt32(header, 20);
            ushort nameLength = ReadUInt16(header, 28);
            ushort extraLength = ReadUInt16(header, 30);
            ushort commentLength = ReadUInt16(header, 32);
            uint localHeaderOffset = ReadUInt32(header, 42);
            if (nameLength == 0 || nameLength > 4096) break;

            var nameBytes = new byte[nameLength];
            if (!ReadExact(stream, nameBytes, 0, nameBytes.Length)) break;
            string name = NormalizeZipEntryName(nameBytes);
            long nextEntryOffset = stream.Position + extraLength + commentLength;
            if (nextEntryOffset < stream.Position || nextEntryOffset > stream.Length) break;

            uint uncompressedSize = ReadUInt32(header, 24);
            hasIWorkIndex |= IsDirectIWorkIndexEntry(name, uncompressedSize);
            if (name == "index.zip" && !nestedIndex.HasValue &&
                TryResolveNestedIndexEntry(stream, header, extraLength, out long resolvedOffset,
                    out uint resolvedCompressedSize, out uint resolvedUncompressedSize)) {
                nestedIndex = (resolvedOffset, compression, resolvedCompressedSize, resolvedUncompressedSize);
            }
            DetectionCandidate? match = MatchContainerEntry(name);
            if (match != null) return match;
            if (name == "mimetype") {
                DetectionCandidate? mimeType = TryReadContainerMimeType(
                    stream,
                    start,
                    localHeaderOffset,
                    compression,
                    compressedSize,
                    nextEntryOffset);
                if (mimeType != null) return mimeType;
            }

            stream.Position = nextEntryOffset;
        }

        if (hasIWorkIndex) return IWorkCandidate();
        if (nestedIndex.HasValue && TryInspectNestedIWorkIndex(stream, start,
                nestedIndex.Value.Offset, nestedIndex.Value.Compression,
                nestedIndex.Value.CompressedSize, nestedIndex.Value.UncompressedSize,
                maxEntries)) return IWorkCandidate();
        return GenericZipCandidate();
    }

    private static async Task<DetectionCandidate> InspectZipContainerAsync(
        Stream stream,
        long start,
        int maxEntries,
        CancellationToken cancellationToken) {
        (bool found, long centralDirectoryOffset, int entryCount) = await TryLocateZipCentralDirectoryAsync(
            stream,
            start,
            cancellationToken).ConfigureAwait(false);
        if (!found) return GenericZipCandidate();

        stream.Position = start + centralDirectoryOffset;
        var header = new byte[ZipCentralDirectoryHeaderLength];
        int entriesToInspect = Math.Min(entryCount, maxEntries);
        bool hasIWorkIndex = false;
        (long Offset, ushort Compression, uint CompressedSize, uint UncompressedSize)? nestedIndex = null;
        for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!await ReadExactAsync(stream, header, 0, header.Length, cancellationToken).ConfigureAwait(false) ||
                ReadUInt32(header, 0) != ZipCentralDirectoryHeaderSignature) {
                break;
            }

            ushort compression = ReadUInt16(header, 10);
            uint compressedSize = ReadUInt32(header, 20);
            ushort nameLength = ReadUInt16(header, 28);
            ushort extraLength = ReadUInt16(header, 30);
            ushort commentLength = ReadUInt16(header, 32);
            uint localHeaderOffset = ReadUInt32(header, 42);
            if (nameLength == 0 || nameLength > 4096) break;

            var nameBytes = new byte[nameLength];
            if (!await ReadExactAsync(stream, nameBytes, 0, nameBytes.Length, cancellationToken).ConfigureAwait(false)) break;
            string name = NormalizeZipEntryName(nameBytes);
            long nextEntryOffset = stream.Position + extraLength + commentLength;
            if (nextEntryOffset < stream.Position || nextEntryOffset > stream.Length) break;

            uint uncompressedSize = ReadUInt32(header, 24);
            hasIWorkIndex |= IsDirectIWorkIndexEntry(name, uncompressedSize);
            if (name == "index.zip" && !nestedIndex.HasValue &&
                await TryResolveNestedIndexEntryAsync(stream, header, extraLength,
                    cancellationToken).ConfigureAwait(false) is { } resolved) {
                nestedIndex = (resolved.Offset, compression, resolved.CompressedSize, resolved.UncompressedSize);
            }
            DetectionCandidate? match = MatchContainerEntry(name);
            if (match != null) return match;
            if (name == "mimetype") {
                DetectionCandidate? mimeType = await TryReadContainerMimeTypeAsync(
                        stream,
                        start,
                        localHeaderOffset,
                        compression,
                        compressedSize,
                        nextEntryOffset,
                        cancellationToken)
                    .ConfigureAwait(false);
                if (mimeType != null) return mimeType;
            }

            stream.Position = nextEntryOffset;
        }

        if (hasIWorkIndex) return IWorkCandidate();
        if (nestedIndex.HasValue && await TryInspectNestedIWorkIndexAsync(stream, start,
                nestedIndex.Value.Offset, nestedIndex.Value.Compression,
                nestedIndex.Value.CompressedSize, nestedIndex.Value.UncompressedSize,
                maxEntries, cancellationToken).ConfigureAwait(false)) return IWorkCandidate();
        return GenericZipCandidate();
    }

    private static bool TryLocateZipCentralDirectory(
        Stream stream,
        long start,
        out long centralDirectoryOffset,
        out int entryCount) {
        centralDirectoryOffset = 0;
        entryCount = 0;
        if (!TryGetZipWindow(stream, start, out long archiveLength, out int tailLength)) return false;

        var tail = new byte[tailLength];
        long tailOffset = archiveLength - tailLength;
        stream.Position = start + tailOffset;
        if (!ReadExact(stream, tail, 0, tail.Length)) return false;
        if (TryParseZipEndRecord(tail, tailOffset, archiveLength,
                out centralDirectoryOffset, out entryCount)) return true;
        if (!TryGetZip64LocatorOffset(tail, tailOffset, archiveLength, out long locatorOffset)) return false;
        return TryParseZip64EndRecord(stream, start, archiveLength, locatorOffset,
            out centralDirectoryOffset, out entryCount);
    }

    private static async Task<(bool Found, long CentralDirectoryOffset, int EntryCount)> TryLocateZipCentralDirectoryAsync(
        Stream stream,
        long start,
        CancellationToken cancellationToken) {
        if (!TryGetZipWindow(stream, start, out long archiveLength, out int tailLength)) {
            return (false, 0, 0);
        }

        var tail = new byte[tailLength];
        long tailOffset = archiveLength - tailLength;
        stream.Position = start + tailOffset;
        if (!await ReadExactAsync(stream, tail, 0, tail.Length, cancellationToken).ConfigureAwait(false)) {
            return (false, 0, 0);
        }

        bool found = TryParseZipEndRecord(
            tail,
            tailOffset,
            archiveLength,
            out long centralDirectoryOffset,
            out int entryCount);
        if (found) return (true, centralDirectoryOffset, entryCount);
        if (!TryGetZip64LocatorOffset(tail, tailOffset, archiveLength, out long locatorOffset)) {
            return (false, 0, 0);
        }
        return await TryParseZip64EndRecordAsync(stream, start, archiveLength,
            locatorOffset, cancellationToken).ConfigureAwait(false);
    }

    private static bool TryGetZipWindow(Stream stream, long start, out long archiveLength, out int tailLength) {
        archiveLength = 0;
        tailLength = 0;
        if (start < 0 || start > stream.Length) return false;

        archiveLength = stream.Length - start;
        if (archiveLength < ZipEndOfCentralDirectoryLength) return false;
        tailLength = (int)Math.Min(archiveLength, ZipMaximumEndRecordLength);
        return true;
    }

    private static bool TryParseZipEndRecord(
        byte[] tail,
        long tailOffset,
        long archiveLength,
        out long centralDirectoryOffset,
        out int entryCount) {
        centralDirectoryOffset = 0;
        entryCount = 0;
        for (int index = tail.Length - ZipEndOfCentralDirectoryLength; index >= 0; index--) {
            if (ReadUInt32(tail, index) != ZipEndOfCentralDirectorySignature) continue;

            ushort commentLength = ReadUInt16(tail, index + 20);
            if (index + ZipEndOfCentralDirectoryLength + commentLength != tail.Length) continue;
            if (ReadUInt16(tail, index + 4) != 0 || ReadUInt16(tail, index + 6) != 0) return false;

            ushort entriesOnDisk = ReadUInt16(tail, index + 8);
            ushort totalEntries = ReadUInt16(tail, index + 10);
            uint centralDirectorySize = ReadUInt32(tail, index + 12);
            uint offset = ReadUInt32(tail, index + 16);
            if (entriesOnDisk != totalEntries ||
                totalEntries == ushort.MaxValue ||
                centralDirectorySize == uint.MaxValue ||
                offset == uint.MaxValue) {
                return false;
            }

            long endRecordOffset = tailOffset + index;
            long centralDirectoryEnd = (long)offset + centralDirectorySize;
            if (centralDirectoryEnd < offset || centralDirectoryEnd > endRecordOffset || endRecordOffset > archiveLength) {
                return false;
            }

            centralDirectoryOffset = offset;
            entryCount = totalEntries;
            return true;
        }

        return false;
    }

    private static bool TryGetZip64LocatorOffset(byte[] tail, long tailOffset,
        long archiveLength, out long locatorOffset) {
        locatorOffset = 0;
        for (int index = tail.Length - ZipEndOfCentralDirectoryLength; index >= 0; index--) {
            if (ReadUInt32(tail, index) != ZipEndOfCentralDirectorySignature) continue;
            if (index + ZipEndOfCentralDirectoryLength + ReadUInt16(tail, index + 20) != tail.Length) continue;
            if (ReadUInt16(tail, index + 4) != 0 || ReadUInt16(tail, index + 6) != 0) return false;
            if (ReadUInt16(tail, index + 8) != ushort.MaxValue &&
                ReadUInt16(tail, index + 10) != ushort.MaxValue &&
                ReadUInt32(tail, index + 12) != uint.MaxValue &&
                ReadUInt32(tail, index + 16) != uint.MaxValue) return false;
            long endRecordOffset = tailOffset + index;
            if (endRecordOffset < 20 || endRecordOffset > archiveLength) return false;
            locatorOffset = endRecordOffset - 20;
            return true;
        }
        return false;
    }

    private static bool TryParseZip64EndRecord(Stream stream, long start, long archiveLength,
        long locatorOffset, out long centralDirectoryOffset, out int entryCount) {
        centralDirectoryOffset = 0;
        entryCount = 0;
        var locator = new byte[20];
        stream.Position = start + locatorOffset;
        if (!ReadExact(stream, locator, 0, locator.Length)) return false;
        if (!TryGetZip64RecordOffset(locator, locatorOffset, out long recordOffset)) return false;
        var record = new byte[56];
        stream.Position = start + recordOffset;
        return ReadExact(stream, record, 0, record.Length) &&
               TryParseZip64Record(record, recordOffset, locatorOffset, archiveLength,
                   out centralDirectoryOffset, out entryCount);
    }

    private static async Task<(bool Found, long CentralDirectoryOffset, int EntryCount)> TryParseZip64EndRecordAsync(
        Stream stream, long start, long archiveLength, long locatorOffset,
        CancellationToken cancellationToken) {
        var locator = new byte[20];
        stream.Position = start + locatorOffset;
        if (!await ReadExactAsync(stream, locator, 0, locator.Length, cancellationToken).ConfigureAwait(false) ||
            !TryGetZip64RecordOffset(locator, locatorOffset, out long recordOffset)) {
            return (false, 0, 0);
        }
        var record = new byte[56];
        stream.Position = start + recordOffset;
        if (!await ReadExactAsync(stream, record, 0, record.Length, cancellationToken).ConfigureAwait(false)) {
            return (false, 0, 0);
        }
        bool found = TryParseZip64Record(record, recordOffset, locatorOffset, archiveLength,
            out long centralDirectoryOffset, out int entryCount);
        return (found, centralDirectoryOffset, entryCount);
    }

    private static bool TryGetZip64RecordOffset(byte[] locator, long locatorOffset,
        out long recordOffset) {
        recordOffset = 0;
        if (ReadUInt32(locator, 0) != Zip64EndOfCentralDirectoryLocatorSignature ||
            ReadUInt32(locator, 4) != 0 || ReadUInt32(locator, 16) != 1 ||
            locatorOffset < 56) return false;
        ulong offset = ReadZipUInt64(locator, 8);
        if (offset > (ulong)(locatorOffset - 56)) return false;
        recordOffset = (long)offset;
        return true;
    }

    private static bool TryParseZip64Record(byte[] record, long recordOffset,
        long locatorOffset, long archiveLength, out long centralDirectoryOffset,
        out int entryCount) {
        centralDirectoryOffset = 0;
        entryCount = 0;
        if (ReadUInt32(record, 0) != Zip64EndOfCentralDirectorySignature ||
            ReadZipUInt64(record, 4) < 44 ||
            ReadZipUInt64(record, 4) > (ulong)(locatorOffset - recordOffset - 12) ||
            ReadUInt32(record, 16) != 0 || ReadUInt32(record, 20) != 0) return false;
        ulong onDisk = ReadZipUInt64(record, 24);
        ulong total = ReadZipUInt64(record, 32);
        ulong size = ReadZipUInt64(record, 40);
        ulong offset = ReadZipUInt64(record, 48);
        if (onDisk != total || total > int.MaxValue || offset > (ulong)recordOffset ||
            size > (ulong)recordOffset - offset || recordOffset > archiveLength) return false;
        centralDirectoryOffset = (long)offset;
        entryCount = (int)total;
        return true;
    }

    private static ulong ReadZipUInt64(byte[] bytes, int offset) =>
        ReadUInt32(bytes, offset) | ((ulong)ReadUInt32(bytes, offset + 4) << 32);

    private static bool TryResolveNestedIndexEntry(Stream stream, byte[] header,
        ushort extraLength, out long offset, out uint compressedSize, out uint uncompressedSize) {
        offset = 0;
        compressedSize = 0;
        uncompressedSize = 0;
        var extra = new byte[extraLength];
        return ReadExact(stream, extra, 0, extra.Length) &&
            TryResolveNestedIndexEntry(header, extra, out offset, out compressedSize, out uncompressedSize);
    }

    private static async Task<(long Offset, uint CompressedSize, uint UncompressedSize)?>
        TryResolveNestedIndexEntryAsync(Stream stream, byte[] header, ushort extraLength,
            CancellationToken cancellationToken) {
        var extra = new byte[extraLength];
        if (!await ReadExactAsync(stream, extra, 0, extra.Length,
                cancellationToken).ConfigureAwait(false)) return null;
        return TryResolveNestedIndexEntry(header, extra, out long offset,
            out uint compressedSize, out uint uncompressedSize)
            ? (offset, compressedSize, uncompressedSize) : null;
    }

    private static bool TryResolveNestedIndexEntry(byte[] header, byte[] extra,
        out long offset, out uint compressedSize, out uint uncompressedSize) {
        uint rawCompressedSize = ReadUInt32(header, 20);
        uint rawUncompressedSize = ReadUInt32(header, 24);
        uint rawOffset = ReadUInt32(header, 42);
        ulong resolvedCompressedSize = rawCompressedSize;
        ulong resolvedUncompressedSize = rawUncompressedSize;
        ulong resolvedOffset = rawOffset;
        if (rawCompressedSize == uint.MaxValue || rawUncompressedSize == uint.MaxValue ||
            rawOffset == uint.MaxValue) {
            bool found = false;
            for (int index = 0; index + 4 <= extra.Length;) {
                ushort id = ReadUInt16(extra, index);
                int length = ReadUInt16(extra, index + 2);
                index += 4;
                if (length > extra.Length - index) break;
                if (id == 0x0001) {
                    int end = index + length;
                    if (rawUncompressedSize == uint.MaxValue) {
                        if (index + 8 > end) break;
                        resolvedUncompressedSize = ReadZipUInt64(extra, index);
                        index += 8;
                    }
                    if (rawCompressedSize == uint.MaxValue) {
                        if (index + 8 > end) break;
                        resolvedCompressedSize = ReadZipUInt64(extra, index);
                        index += 8;
                    }
                    if (rawOffset == uint.MaxValue) {
                        if (index + 8 > end) break;
                        resolvedOffset = ReadZipUInt64(extra, index);
                    }
                    found = true;
                    break;
                }
                index += length;
            }
            if (!found) {
                offset = 0;
                compressedSize = uncompressedSize = 0;
                return false;
            }
        }
        if (resolvedOffset > long.MaxValue ||
            resolvedCompressedSize > MaximumNestedIWorkIndexProbeBytes ||
            resolvedUncompressedSize > MaximumNestedIWorkIndexProbeBytes) {
            offset = 0;
            compressedSize = uncompressedSize = 0;
            return false;
        }
        offset = (long)resolvedOffset;
        compressedSize = (uint)resolvedCompressedSize;
        uncompressedSize = (uint)resolvedUncompressedSize;
        return compressedSize > 0 && uncompressedSize > 0;
    }

    private static DetectionCandidate? TryReadContainerMimeType(
        Stream stream,
        long start,
        uint localHeaderOffset,
        ushort compression,
        uint compressedSize,
        long returnPosition) {
        if ((compression != 0 && compression != 8) || compressedSize == 0 || compressedSize > 128) return null;

        var header = new byte[30];
        long headerPosition = start + localHeaderOffset;
        if (headerPosition < start || headerPosition > stream.Length - header.Length) return null;
        stream.Position = headerPosition;
        if (!ReadExact(stream, header, 0, header.Length) || ReadUInt32(header, 0) != ZipLocalHeaderSignature) {
            stream.Position = returnPosition;
            return null;
        }

        ushort nameLength = ReadUInt16(header, 26);
        ushort extraLength = ReadUInt16(header, 28);
        long dataPosition = stream.Position + nameLength + extraLength;
        if (dataPosition < stream.Position || dataPosition > stream.Length - compressedSize) {
            stream.Position = returnPosition;
            return null;
        }

        stream.Position = dataPosition;
        var mimeBytes = new byte[(int)compressedSize];
        DetectionCandidate? candidate = ReadExact(stream, mimeBytes, 0, mimeBytes.Length)
            ? MatchContainerMimeType(mimeBytes, compression)
            : null;
        stream.Position = returnPosition;
        return candidate;
    }

    private static async Task<DetectionCandidate?> TryReadContainerMimeTypeAsync(
        Stream stream,
        long start,
        uint localHeaderOffset,
        ushort compression,
        uint compressedSize,
        long returnPosition,
        CancellationToken cancellationToken) {
        if ((compression != 0 && compression != 8) || compressedSize == 0 || compressedSize > 128) return null;

        var header = new byte[30];
        long headerPosition = start + localHeaderOffset;
        if (headerPosition < start || headerPosition > stream.Length - header.Length) return null;
        stream.Position = headerPosition;
        if (!await ReadExactAsync(stream, header, 0, header.Length, cancellationToken).ConfigureAwait(false) ||
            ReadUInt32(header, 0) != ZipLocalHeaderSignature) {
            stream.Position = returnPosition;
            return null;
        }

        ushort nameLength = ReadUInt16(header, 26);
        ushort extraLength = ReadUInt16(header, 28);
        long dataPosition = stream.Position + nameLength + extraLength;
        if (dataPosition < stream.Position || dataPosition > stream.Length - compressedSize) {
            stream.Position = returnPosition;
            return null;
        }

        stream.Position = dataPosition;
        var mimeBytes = new byte[(int)compressedSize];
        DetectionCandidate? candidate = await ReadExactAsync(stream, mimeBytes, 0, mimeBytes.Length, cancellationToken)
                .ConfigureAwait(false)
            ? MatchContainerMimeType(mimeBytes, compression)
            : null;
        stream.Position = returnPosition;
        return candidate;
    }

    private static string NormalizeZipEntryName(byte[] nameBytes) {
        return Encoding.UTF8.GetString(nameBytes).Replace('\\', '/').ToLowerInvariant();
    }

    private static DetectionCandidate? MatchContainerMimeType(byte[] mimeBytes, ushort compression) {
        if (compression == 8) {
            byte[]? inflated = InflateMimeType(mimeBytes);
            if (inflated == null) return null;
            mimeBytes = inflated;
        }
        string mediaType = Encoding.ASCII.GetString(mimeBytes).Trim();
        if (string.Equals(mediaType, "application/epub+zip", StringComparison.Ordinal)) {
            return EpubCandidate();
        }
        if (string.Equals(mediaType, "application/vnd.oasis.opendocument.text", StringComparison.Ordinal) ||
            string.Equals(mediaType, "application/vnd.oasis.opendocument.spreadsheet", StringComparison.Ordinal) ||
            string.Equals(mediaType, "application/vnd.oasis.opendocument.presentation", StringComparison.Ordinal)) {
            return OpenDocumentCandidate(mediaType);
        }
        return null;
    }

    private static byte[]? InflateMimeType(byte[] compressedBytes) {
        try {
            using (var input = new MemoryStream(compressedBytes, false))
            using (var inflater = new DeflateStream(input, CompressionMode.Decompress)) {
                var output = new byte[129];
                int total = 0;
                while (total < output.Length) {
                    int read = inflater.Read(output, total, output.Length - total);
                    if (read <= 0) break;
                    total += read;
                }
                if (total > 128) return null;
                var result = new byte[total];
                Buffer.BlockCopy(output, 0, result, 0, total);
                return result;
            }
        } catch (InvalidDataException) {
            return null;
        } catch (IOException) {
            return null;
        }
    }

    private static DetectionCandidate GenericZipCandidate() {
        return DetectionCandidate.High(ReaderInputKind.Zip, "application/zip", "container:zip-generic");
    }

    private static bool IsDirectIWorkIndexEntry(string name, uint length) =>
        length > 0 && name == "index/document.iwa";

    private static DetectionCandidate IWorkCandidate() => DetectionCandidate.High(
        ReaderInputKind.IWork, "application/octet-stream", "container:iwork-index");

    private static DetectionCandidate EpubCandidate() {
        return DetectionCandidate.High(
            ReaderInputKind.Epub,
            "application/epub+zip",
            "container:epub-mimetype",
            mediaTypeIsDeclared: true);
    }

    private static DetectionCandidate OpenDocumentCandidate(string mediaType) {
        return DetectionCandidate.High(
            ReaderInputKind.OpenDocument,
            mediaType,
            "container:opendocument-mimetype",
            mediaTypeIsDeclared: true);
    }
}
