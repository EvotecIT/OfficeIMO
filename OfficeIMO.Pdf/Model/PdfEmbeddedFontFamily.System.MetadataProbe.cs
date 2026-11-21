namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxSystemFontTableCountToInspect = 4096;
    internal const int MaxSystemFontNameTableBytes = 4 * 1024 * 1024;
    internal const int MaxSystemFontNameMetadataBytes = 16 * 1024 * 1024;
    internal const int MaxSystemFontNameRecords = 4096;
    internal const int MaxSystemFontDecodedNameBytes = 128 * 1024;

    // Probing reads a few small ranges from hundreds of installed files. The file header and table
    // directory almost always sit in the first block, so one reusable buffer serves them. Name tables
    // are parsed from a second reusable buffer, as only the decoded names outlive the probe.
    private const int SystemFontProbePrefixBytes = 4096;
    private const int SystemFontProbeNameTableBufferBytes = 64 * 1024;
    [System.ThreadStatic] private static byte[]? _systemFontProbePrefix;
    [System.ThreadStatic] private static byte[]? _systemFontProbeNameTable;

    private static bool TryReadSystemFontNameMetadata(string path,
        out System.Collections.Generic.List<TrueTypeNameMetadata>? metadataFaces,
        out bool inspectedAllFaces) {
        metadataFaces = null;
        inspectedAllFaces = false;
        try {
            using var stream = new System.IO.FileStream(path, System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete,
                bufferSize: 1);
            if (stream.Length < 12 || stream.Length > MaxSystemFontFileBytes) {
                return false;
            }

            var reader = new SystemFontProbeReader(stream);
            int headerStart = reader.Read(0, 12, out byte[] header);
            var offsets = new System.Collections.Generic.List<long>();
            if (header[headerStart] == (byte)'t' && header[headerStart + 1] == (byte)'t' &&
                header[headerStart + 2] == (byte)'c' && header[headerStart + 3] == (byte)'f') {
                uint count = ReadUInt32(header, headerStart + 8);
                if (count == 0 || count > MaxTrueTypeCollectionFontsToInspect) {
                    return false;
                }

                int offsetStart = reader.Read(12, checked((int)count * 4), out byte[] offsetBytes);
                for (int index = 0; index < count; index++) {
                    offsets.Add(ReadUInt32(offsetBytes, offsetStart + index * 4));
                }
            } else {
                offsets.Add(0L);
            }

            var found = new System.Collections.Generic.List<TrueTypeNameMetadata>(offsets.Count);
            System.Collections.Generic.Dictionary<long, SystemFontNameTableCacheEntry>? nameTables = null;
            long nameTableBytes = 0L;
            bool complete = true;
            for (int index = 0; index < offsets.Count; index++) {
                if (TryReadSystemFontFaceNameMetadata(reader, offsets[index],
                        ref nameTables, ref nameTableBytes,
                        out TrueTypeNameMetadata? metadata) && metadata != null) {
                    found.Add(metadata);
                } else {
                    complete = false;
                }
            }

            metadataFaces = found;
            inspectedAllFaces = complete;
            return found.Count > 0;
        } catch (System.Exception exception) when (
            exception is System.IO.IOException ||
            exception is System.UnauthorizedAccessException ||
            exception is System.NotSupportedException ||
            exception is System.ArgumentException ||
            exception is System.ArithmeticException ||
            exception is System.FormatException ||
            exception is System.IndexOutOfRangeException ||
            exception is System.InvalidOperationException) {
            return false;
        }
    }

    private static bool TryReadSystemFontFaceNameMetadata(SystemFontProbeReader reader,
        long faceOffset,
        ref System.Collections.Generic.Dictionary<long, SystemFontNameTableCacheEntry>? nameTables,
        ref long nameTableBytes,
        out TrueTypeNameMetadata? metadata) {
        metadata = null;
        int headerStart = reader.Read(faceOffset, 12, out byte[] header);
        int tableCount = ReadUInt16(header, headerStart + 4);
        if (tableCount == 0 || tableCount > MaxSystemFontTableCountToInspect) {
            return false;
        }

        int directoryStart = reader.Read(faceOffset + 12,
            checked(tableCount * 16), out byte[] directory);
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = directoryStart + index * 16;
            if (directory[recordOffset] != (byte)'n' ||
                directory[recordOffset + 1] != (byte)'a' ||
                directory[recordOffset + 2] != (byte)'m' ||
                directory[recordOffset + 3] != (byte)'e') {
                continue;
            }

            uint tableOffset = ReadUInt32(directory, recordOffset + 8);
            uint tableLength = ReadUInt32(directory, recordOffset + 12);
            if (tableLength == 0 || tableLength > MaxSystemFontNameTableBytes ||
                tableOffset > long.MaxValue - tableLength) {
                return false;
            }

            int length = checked((int)tableLength);
            if (nameTables != null && nameTables.TryGetValue(tableOffset,
                    out SystemFontNameTableCacheEntry? cached)) {
                // Faces sharing a name table decode the same bytes to the same names.
                if (cached == null || cached.Length != length) {
                    return false;
                }
                metadata = cached.Metadata;
                return metadata != null;
            }

            if (nameTableBytes > MaxSystemFontNameMetadataBytes - length) {
                throw new System.NotSupportedException(
                    "TrueType collection name-table metadata exceeds supported limits.");
            }
            byte[] table = GetProbeNameTableBuffer(length);
            reader.ReadInto(tableOffset, length, table);
            bool parsed = TryReadTrueTypeNameTable(table, length, 0, length, out metadata);
            (nameTables ??= new System.Collections.Generic.Dictionary<long, SystemFontNameTableCacheEntry>())
                .Add(tableOffset, new SystemFontNameTableCacheEntry(length, parsed ? metadata : null));
            nameTableBytes += length;
            return parsed;
        }

        return false;
    }

    private static byte[] GetProbeNameTableBuffer(int length) =>
        length <= SystemFontProbeNameTableBufferBytes
            ? _systemFontProbeNameTable ??= new byte[SystemFontProbeNameTableBufferBytes]
            : new byte[length];

    /// <summary>
    /// Scores the only face of a non-collection font file from its name and OS/2 tables, giving the kind
    /// and score that reading and classifying the whole file would. Returns false when the file needs the
    /// whole read to decide: collections, duplicate or out-of-range tables, or unreadable names.
    /// </summary>
    private static bool TryScoreSingleSystemFontFace(string path, string normalizedMetadataFamily,
        out bool hasCandidate, out FontFaceKind kind, out int score) {
        hasCandidate = false;
        kind = FontFaceKind.Regular;
        score = 0;
        try {
            var fileInfo = new System.IO.FileInfo(path);
            if (!fileInfo.Exists ||
                (fileInfo.Attributes & System.IO.FileAttributes.ReparsePoint) != 0 ||
                fileInfo.Length > MaxSystemFontFileBytes) {
                return false;
            }

            using var stream = new System.IO.FileStream(path, System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete,
                bufferSize: 1);
            long fileLength = stream.Length;
            if (fileLength < 12 || fileLength > MaxSystemFontFileBytes) {
                return false;
            }

            var reader = new SystemFontProbeReader(stream);
            int headerStart = reader.Read(0, 12, out byte[] header);
            if (header[headerStart] == (byte)'t' && header[headerStart + 1] == (byte)'t' &&
                header[headerStart + 2] == (byte)'c' && header[headerStart + 3] == (byte)'f') {
                return false;
            }

            int tableCount = ReadUInt16(header, headerStart + 4);
            if (tableCount == 0 || tableCount > MaxSystemFontTableCountToInspect) {
                return false;
            }

            int directoryStart = reader.Read(12, checked(tableCount * 16), out byte[] directory);
            long nameOffset = 0L, os2Offset = 0L;
            int nameLength = 0, os2Length = 0, nameRecords = 0, os2Records = 0;
            for (int index = 0; index < tableCount; index++) {
                int record = directoryStart + index * 16;
                uint tableOffset = ReadUInt32(directory, record + 8);
                uint tableLength = ReadUInt32(directory, record + 12);
                // The whole-file directory read rejects these ranges; leave such files to it.
                if (tableOffset > int.MaxValue || tableLength > int.MaxValue ||
                    tableOffset > fileLength - tableLength) {
                    return false;
                }

                if (directory[record] == (byte)'n' && directory[record + 1] == (byte)'a' &&
                    directory[record + 2] == (byte)'m' && directory[record + 3] == (byte)'e') {
                    nameRecords++;
                    nameOffset = tableOffset;
                    nameLength = (int)tableLength;
                } else if (directory[record] == (byte)'O' && directory[record + 1] == (byte)'S' &&
                    directory[record + 2] == (byte)'/' && directory[record + 3] == (byte)'2') {
                    os2Records++;
                    os2Offset = tableOffset;
                    os2Length = (int)tableLength;
                }
            }

            if (nameRecords != 1 || os2Records > 1 || nameLength == 0 || nameLength > MaxSystemFontNameTableBytes) {
                return false;
            }

            byte[] table = GetProbeNameTableBuffer(nameLength);
            reader.ReadInto(nameOffset, nameLength, table);
            if (!TryReadTrueTypeNameTable(table, nameLength, 0, nameLength, out TrueTypeNameMetadata? metadata) || metadata == null) {
                return false;
            }

            int familyScore = GetMetadataFamilyMatchScore(metadata, normalizedMetadataFamily);
            if (familyScore < 0) {
                // The whole-file path rejects it on the same names without reading it.
                return true;
            }

            bool hasStyleFlags = false;
            int weight = 0;
            bool bold = false, italic = false, oblique = false;
            if (os2Records == 1 && os2Length >= 64) {
                int os2Start = reader.Read(os2Offset, 64, out byte[] os2);
                hasStyleFlags = TryReadSystemFontStyleFlags(os2, os2Start, out weight, out bold, out italic, out oblique);
            }

            kind = ClassifyMetadataFace(metadata, hasStyleFlags, weight, bold, italic, oblique, out int metadataScore);
            score = metadataScore + familyScore;
            hasCandidate = true;
            return true;
        } catch (System.Exception exception) when (
            exception is System.IO.IOException ||
            exception is System.UnauthorizedAccessException ||
            exception is System.NotSupportedException ||
            exception is System.ArgumentException ||
            exception is System.ArithmeticException ||
            exception is System.FormatException ||
            exception is System.IndexOutOfRangeException ||
            exception is System.InvalidOperationException) {
            return false;
        }
    }

    private static byte[] ReadFontFileRange(System.IO.FileStream stream,
        long offset, int length) {
        var buffer = new byte[length];
        ReadFontFileRange(stream, offset, length, buffer);
        return buffer;
    }

    private static void ReadFontFileRange(System.IO.FileStream stream,
        long offset, int length, byte[] buffer) {
        if (offset < 0 || length < 0 || offset > stream.Length - length) {
            throw new System.NotSupportedException(
                "TrueType font table data is truncated or invalid.");
        }

        stream.Position = offset;
        int read = 0;
        while (read < length) {
            int count = stream.Read(buffer, read, length - read);
            if (count == 0) {
                throw new System.NotSupportedException(
                    "TrueType font table data is truncated or invalid.");
            }

            read += count;
        }
    }

    /// <summary>Serves small ranges from the start of a font file out of one shared buffer.</summary>
    private sealed class SystemFontProbeReader {
        private readonly System.IO.FileStream _stream;
        private readonly byte[] _prefix;
        private readonly int _prefixLength;

        internal SystemFontProbeReader(System.IO.FileStream stream) {
            _stream = stream;
            _prefix = _systemFontProbePrefix ??= new byte[SystemFontProbePrefixBytes];
            _prefixLength = (int)System.Math.Min(SystemFontProbePrefixBytes, stream.Length);
            ReadFontFileRange(stream, 0, _prefixLength, _prefix);
        }

        /// <summary>Returns the buffer holding the range and the range's start within it.</summary>
        internal int Read(long offset, int length, out byte[] buffer) {
            if (offset < 0 || length < 0 || offset > _stream.Length - length) {
                throw new System.NotSupportedException(
                    "TrueType font table data is truncated or invalid.");
            }

            if (offset + length <= _prefixLength) {
                buffer = _prefix;
                return (int)offset;
            }

            buffer = ReadFontFileRange(_stream, offset, length);
            return 0;
        }

        internal void ReadInto(long offset, int length, byte[] buffer) {
            if (offset < 0 || length < 0 || offset > _stream.Length - length) {
                throw new System.NotSupportedException(
                    "TrueType font table data is truncated or invalid.");
            }

            if (offset + length <= _prefixLength) {
                System.Buffer.BlockCopy(_prefix, (int)offset, buffer, 0, length);
                return;
            }

            ReadFontFileRange(_stream, offset, length, buffer);
        }
    }

    private sealed class SystemFontNameTableCacheEntry {
        internal SystemFontNameTableCacheEntry(int length, TrueTypeNameMetadata? metadata) {
            Length = length;
            Metadata = metadata;
        }

        internal int Length { get; }
        internal TrueTypeNameMetadata? Metadata { get; }
    }
}
