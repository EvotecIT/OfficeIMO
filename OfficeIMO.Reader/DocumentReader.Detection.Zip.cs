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
        return ReadExact(stream, tail, 0, tail.Length) &&
               TryParseZipEndRecord(tail, tailOffset, archiveLength, out centralDirectoryOffset, out entryCount);
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
        return (found, centralDirectoryOffset, entryCount);
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
