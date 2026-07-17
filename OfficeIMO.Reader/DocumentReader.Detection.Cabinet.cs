using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private const int CabinetHeaderLength = 36;
    private const int CabinetFileHeaderLength = 16;
    private const int CabinetMaximumFileNameBytes = 4096;

    private static bool IsCabinet(byte[] prefix) {
        return prefix.Length >= 4 &&
               prefix[0] == (byte)'M' &&
               prefix[1] == (byte)'S' &&
               prefix[2] == (byte)'C' &&
               prefix[3] == (byte)'F';
    }

    private static DetectionCandidate InspectCabinetContainer(Stream stream, long start, int maxEntries) {
        if (!TryGetCabinetWindow(stream, start, out long archiveLength)) return GenericCabinetCandidate();

        var header = new byte[CabinetHeaderLength];
        stream.Position = start;
        if (!ReadExact(stream, header, 0, header.Length) ||
            !TryParseCabinetHeader(header, archiveLength, out uint declaredSize, out uint filesOffset, out int fileCount)) {
            return GenericCabinetCandidate();
        }

        long declaredEnd = start + declaredSize;
        stream.Position = start + filesOffset;
        var fileHeader = new byte[CabinetFileHeaderLength];
        int entriesToInspect = Math.Min(fileCount, maxEntries);
        for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
            if (stream.Position > declaredEnd - fileHeader.Length ||
                !ReadExact(stream, fileHeader, 0, fileHeader.Length) ||
                !TryReadCabinetFileName(stream, declaredEnd, out string name)) {
                break;
            }

            if (IsTopLevelOneNoteTableOfContents(name)) return OneNotePackageCandidate();
        }

        return GenericCabinetCandidate();
    }

    private static async Task<DetectionCandidate> InspectCabinetContainerAsync(
        Stream stream,
        long start,
        int maxEntries,
        CancellationToken cancellationToken) {
        if (!TryGetCabinetWindow(stream, start, out long archiveLength)) return GenericCabinetCandidate();

        var header = new byte[CabinetHeaderLength];
        stream.Position = start;
        if (!await ReadExactAsync(stream, header, 0, header.Length, cancellationToken).ConfigureAwait(false) ||
            !TryParseCabinetHeader(header, archiveLength, out uint declaredSize, out uint filesOffset, out int fileCount)) {
            return GenericCabinetCandidate();
        }

        long declaredEnd = start + declaredSize;
        stream.Position = start + filesOffset;
        var fileHeader = new byte[CabinetFileHeaderLength];
        int entriesToInspect = Math.Min(fileCount, maxEntries);
        for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (stream.Position > declaredEnd - fileHeader.Length ||
                !await ReadExactAsync(stream, fileHeader, 0, fileHeader.Length, cancellationToken).ConfigureAwait(false)) {
                break;
            }

            (bool found, string name) = await TryReadCabinetFileNameAsync(
                stream,
                declaredEnd,
                cancellationToken).ConfigureAwait(false);
            if (!found) break;
            if (IsTopLevelOneNoteTableOfContents(name)) return OneNotePackageCandidate();
        }

        return GenericCabinetCandidate();
    }

    private static DetectionCandidate InspectCabinetContainer(byte[] boundedPayload, int maxEntries) {
        if (boundedPayload.Length < CabinetHeaderLength ||
            !TryParseCabinetHeader(boundedPayload, null, out uint declaredSize, out uint filesOffset, out int fileCount)) {
            return GenericCabinetCandidate();
        }

        int declaredEnd = (int)Math.Min(declaredSize, (uint)boundedPayload.Length);
        if (filesOffset > (uint)declaredEnd) return GenericCabinetCandidate();
        int offset = (int)filesOffset;
        int entriesToInspect = Math.Min(fileCount, maxEntries);
        for (int entryIndex = 0; entryIndex < entriesToInspect; entryIndex++) {
            if (offset > declaredEnd - CabinetFileHeaderLength) break;
            int nameOffset = offset + CabinetFileHeaderLength;
            int nameEndLimit = Math.Min(declaredEnd, nameOffset + CabinetMaximumFileNameBytes + 1);
            int nameEnd = nameOffset;
            while (nameEnd < nameEndLimit && boundedPayload[nameEnd] != 0) nameEnd++;
            if (nameEnd == nameEndLimit) break;

            string name = Encoding.UTF8.GetString(boundedPayload, nameOffset, nameEnd - nameOffset);
            if (IsTopLevelOneNoteTableOfContents(name)) return OneNotePackageCandidate();
            offset = nameEnd + 1;
        }

        return GenericCabinetCandidate();
    }

    private static bool TryGetCabinetWindow(Stream stream, long start, out long archiveLength) {
        archiveLength = 0;
        if (start < 0 || start > stream.Length) return false;
        archiveLength = stream.Length - start;
        return archiveLength >= CabinetHeaderLength;
    }

    private static bool TryParseCabinetHeader(
        byte[] header,
        long? archiveLength,
        out uint declaredSize,
        out uint filesOffset,
        out int fileCount) {
        declaredSize = 0;
        filesOffset = 0;
        fileCount = 0;
        if (header.Length < CabinetHeaderLength || !IsCabinet(header)) return false;

        declaredSize = ReadUInt32(header, 8);
        filesOffset = ReadUInt32(header, 16);
        int folderCount = ReadUInt16(header, 26);
        fileCount = ReadUInt16(header, 28);
        if (declaredSize < CabinetHeaderLength ||
            archiveLength.HasValue && declaredSize > archiveLength.Value ||
            folderCount < 1 ||
            fileCount < 1 ||
            filesOffset < CabinetHeaderLength ||
            (long)filesOffset + CabinetFileHeaderLength > declaredSize) {
            return false;
        }

        return true;
    }

    private static bool TryReadCabinetFileName(Stream stream, long declaredEnd, out string name) {
        name = string.Empty;
        long remaining = declaredEnd - stream.Position;
        if (remaining < 1) return false;
        int bytesToRead = (int)Math.Min(remaining, CabinetMaximumFileNameBytes + 1L);
        var bytes = new byte[bytesToRead];
        long nameStart = stream.Position;
        if (!ReadExact(stream, bytes, 0, bytes.Length)) return false;

        int terminator = Array.IndexOf(bytes, (byte)0);
        if (terminator < 0) return false;
        stream.Position = nameStart + terminator + 1;
        name = Encoding.UTF8.GetString(bytes, 0, terminator);
        return true;
    }

    private static async Task<(bool Found, string Name)> TryReadCabinetFileNameAsync(
        Stream stream,
        long declaredEnd,
        CancellationToken cancellationToken) {
        long remaining = declaredEnd - stream.Position;
        if (remaining < 1) return (false, string.Empty);
        int bytesToRead = (int)Math.Min(remaining, CabinetMaximumFileNameBytes + 1L);
        var bytes = new byte[bytesToRead];
        long nameStart = stream.Position;
        if (!await ReadExactAsync(stream, bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false)) {
            return (false, string.Empty);
        }

        int terminator = Array.IndexOf(bytes, (byte)0);
        if (terminator < 0) return (false, string.Empty);
        stream.Position = nameStart + terminator + 1;
        return (true, Encoding.UTF8.GetString(bytes, 0, terminator));
    }

    private static bool IsTopLevelOneNoteTableOfContents(string name) {
        string normalized = name.Replace('\\', '/');
        return normalized.IndexOf('/') < 0 &&
               normalized.EndsWith(".onetoc2", StringComparison.OrdinalIgnoreCase);
    }

    private static DetectionCandidate GenericCabinetCandidate() {
        return DetectionCandidate.Unknown("container:cabinet-generic");
    }

    private static DetectionCandidate OneNotePackageCandidate() {
        return DetectionCandidate.High(ReaderInputKind.OneNote, "application/onenote", "container:onenote-package");
    }
}
