namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxSystemFontTableCountToInspect = 4096;
    internal const int MaxSystemFontNameTableBytes = 4 * 1024 * 1024;
    internal const int MaxSystemFontNameMetadataBytes = 16 * 1024 * 1024;

    private static bool TryReadSystemFontNameMetadata(string path,
        out System.Collections.Generic.List<TrueTypeNameMetadata>? metadataFaces) {
        metadataFaces = null;
        try {
            using var stream = new System.IO.FileStream(path, System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete);
            if (stream.Length < 12 || stream.Length > MaxSystemFontFileBytes) {
                return false;
            }

            byte[] header = ReadFontFileRange(stream, 0, 12);
            var offsets = new System.Collections.Generic.List<long>();
            if (header[0] == (byte)'t' && header[1] == (byte)'t' &&
                header[2] == (byte)'c' && header[3] == (byte)'f') {
                uint count = ReadUInt32(header, 8);
                if (count == 0 || count > MaxTrueTypeCollectionFontsToInspect) {
                    return false;
                }

                byte[] offsetBytes = ReadFontFileRange(stream, 12,
                    checked((int)count * 4));
                for (int index = 0; index < count; index++) {
                    offsets.Add(ReadUInt32(offsetBytes, index * 4));
                }
            } else {
                offsets.Add(0L);
            }

            var found = new System.Collections.Generic.List<TrueTypeNameMetadata>(offsets.Count);
            var nameTables = new System.Collections.Generic.Dictionary<long,
                SystemFontNameTableCacheEntry>();
            long nameTableBytes = 0L;
            for (int index = 0; index < offsets.Count; index++) {
                if (TryReadSystemFontFaceNameMetadata(stream, offsets[index],
                        nameTables, ref nameTableBytes,
                        out TrueTypeNameMetadata? metadata) && metadata != null) {
                    found.Add(metadata);
                }
            }

            metadataFaces = found;
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

    private static bool TryReadSystemFontFaceNameMetadata(System.IO.FileStream stream,
        long faceOffset,
        System.Collections.Generic.IDictionary<long,
            SystemFontNameTableCacheEntry> nameTables,
        ref long nameTableBytes,
        out TrueTypeNameMetadata? metadata) {
        metadata = null;
        byte[] header = ReadFontFileRange(stream, faceOffset, 12);
        int tableCount = ReadUInt16(header, 4);
        if (tableCount == 0 || tableCount > MaxSystemFontTableCountToInspect) {
            return false;
        }

        byte[] directory = ReadFontFileRange(stream, faceOffset + 12,
            checked(tableCount * 16));
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = index * 16;
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
            byte[] table;
            if (nameTables.TryGetValue(tableOffset,
                    out SystemFontNameTableCacheEntry? cached)) {
                if (cached == null || cached.Length != length) {
                    return false;
                }
                table = cached.Data;
            } else {
                if (nameTableBytes > MaxSystemFontNameMetadataBytes - length) {
                    throw new System.NotSupportedException(
                        "TrueType collection name-table metadata exceeds supported limits.");
                }
                table = ReadFontFileRange(stream, tableOffset, length);
                nameTables.Add(tableOffset,
                    new SystemFontNameTableCacheEntry(length, table));
                nameTableBytes += length;
            }
            return TryReadTrueTypeNameTable(table, 0, table.Length,
                out metadata);
        }

        return false;
    }

    private static byte[] ReadFontFileRange(System.IO.FileStream stream,
        long offset, int length) {
        if (offset < 0 || length < 0 || offset > stream.Length - length) {
            throw new System.NotSupportedException(
                "TrueType font table data is truncated or invalid.");
        }

        stream.Position = offset;
        byte[] buffer = new byte[length];
        int read = 0;
        while (read < length) {
            int count = stream.Read(buffer, read, length - read);
            if (count == 0) {
                throw new System.NotSupportedException(
                    "TrueType font table data is truncated or invalid.");
            }

            read += count;
        }

        return buffer;
    }

    private sealed class SystemFontNameTableCacheEntry {
        internal SystemFontNameTableCacheEntry(int length, byte[] data) {
            Length = length;
            Data = data;
        }

        internal int Length { get; }
        internal byte[] Data { get; }
    }
}
