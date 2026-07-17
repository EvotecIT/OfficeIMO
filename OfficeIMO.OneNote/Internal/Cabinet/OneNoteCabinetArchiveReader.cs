namespace OfficeIMO.OneNote;

internal static class OneNoteCabinetArchiveReader {
    private const ushort CompressionNone = 0x0000;
    private const ushort CompressionMsZip = 0x0001;
    private const ushort CompressionLzx = 0x0003;
    private const ushort CompressionMask = 0x000F;
    private const ushort FlagPreviousCabinet = 0x0001;
    private const ushort FlagNextCabinet = 0x0002;
    private const ushort FlagReservePresent = 0x0004;
    private const ushort UtfNameAttribute = 0x0080;

    internal static IReadOnlyList<OneNoteCabinetEntry> Read(
        Stream stream,
        long? maxInputBytes,
        long maxExpandedBytes,
        long maxEntryBytes,
        int maxEntries) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The CAB stream must be readable.", nameof(stream));
        if (!stream.CanSeek) throw new ArgumentException("The CAB stream must be seekable.", nameof(stream));
        if (maxInputBytes.HasValue && stream.Length > maxInputBytes.Value) throw new IOException("The OneNote package exceeds MaxInputBytes.");
        if (stream.Length > int.MaxValue) throw new IOException("The OneNote package is too large to materialize.");
        long originalPosition = stream.Position;
        try {
            stream.Position = 0;
            byte[] data = ReadExactly(stream, (int)stream.Length);
            return Read(data, maxExpandedBytes, maxEntryBytes, maxEntries);
        } finally {
            stream.Position = originalPosition;
        }
    }

    internal static IReadOnlyList<OneNoteCabinetEntry> Read(byte[] data, long maxExpandedBytes, long maxEntryBytes, int maxEntries) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (maxExpandedBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxExpandedBytes));
        if (maxEntryBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxEntryBytes));
        if (maxEntries < 1) throw new ArgumentOutOfRangeException(nameof(maxEntries));
        if (data.Length < 36 || data[0] != (byte)'M' || data[1] != (byte)'S' || data[2] != (byte)'C' || data[3] != (byte)'F') {
            throw new OneNoteFormatException("ONENOTE_CAB_SIGNATURE", "The .onepkg file is not a Microsoft Cabinet archive.");
        }
        uint declaredSize = ReadUInt32(data, 8);
        uint filesOffset = ReadUInt32(data, 16);
        int folderCount = ReadUInt16(data, 26);
        int fileCount = ReadUInt16(data, 28);
        ushort flags = ReadUInt16(data, 30);
        if (declaredSize > data.Length || folderCount < 1 || fileCount > maxEntries) {
            throw new OneNoteFormatException("ONENOTE_CAB_HEADER", "The CAB header contains invalid sizes or entry counts.");
        }

        int offset = 36;
        int folderReserve = 0;
        int dataReserve = 0;
        if ((flags & FlagReservePresent) != 0) {
            Ensure(data, offset, 4, "CAB reserve header");
            int headerReserve = ReadUInt16(data, offset);
            folderReserve = data[offset + 2];
            dataReserve = data[offset + 3];
            offset = checked(offset + 4 + headerReserve);
            Ensure(data, offset, 0, "CAB reserve header");
        }
        if ((flags & FlagPreviousCabinet) != 0) {
            offset = SkipCString(data, offset);
            offset = SkipCString(data, offset);
        }
        if ((flags & FlagNextCabinet) != 0) {
            offset = SkipCString(data, offset);
            offset = SkipCString(data, offset);
        }

        var folders = new FolderHeader[folderCount];
        for (int index = 0; index < folders.Length; index++) {
            Ensure(data, offset, 8, "CFFOLDER");
            folders[index] = new FolderHeader(ReadUInt32(data, offset), ReadUInt16(data, offset + 4), ReadUInt16(data, offset + 6));
            offset = checked(offset + 8 + folderReserve);
        }

        if (filesOffset > int.MaxValue) throw new OneNoteFormatException("ONENOTE_CAB_FILE_TABLE", "The CAB file table offset is too large.");
        offset = (int)filesOffset;
        var files = new FileHeader[fileCount];
        for (int index = 0; index < files.Length; index++) {
            Ensure(data, offset, 16, "CFFILE");
            uint length = ReadUInt32(data, offset);
            uint folderOffset = ReadUInt32(data, offset + 4);
            ushort folderIndex = ReadUInt16(data, offset + 8);
            ushort attributes = ReadUInt16(data, offset + 14);
            ReadCString(data, offset + 16, (attributes & UtfNameAttribute) != 0, out string name, out int nextOffset);
            if (length > maxEntryBytes) throw new OneNoteFormatException("ONENOTE_CAB_ENTRY_LIMIT", "A .onepkg entry exceeds the configured size limit.");
            files[index] = new FileHeader(length, folderOffset, folderIndex, name);
            offset = nextOffset;
        }

        var folderData = new byte[folderCount][];
        long totalExpanded = 0;
        for (int index = 0; index < folders.Length; index++) {
            folderData[index] = DecompressFolder(data, folders[index], dataReserve, maxExpandedBytes - totalExpanded);
            totalExpanded += folderData[index].LongLength;
            if (totalExpanded > maxExpandedBytes) throw new OneNoteFormatException("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded .onepkg archive exceeds the configured size limit.");
        }

        var entries = new List<OneNoteCabinetEntry>(files.Length);
        long totalExtractedBytes = 0;
        foreach (FileHeader file in files) {
            if (file.FolderIndex >= folderData.Length) throw new OneNoteFormatException("ONENOTE_CAB_FOLDER", "A .onepkg entry references a missing or continued CAB folder.");
            byte[] source = folderData[file.FolderIndex];
            long end = (long)file.FolderOffset + file.Length;
            if (file.FolderOffset > source.Length || end > source.Length) throw new OneNoteFormatException("ONENOTE_CAB_ENTRY_RANGE", "A .onepkg entry extends past its CAB folder.");
            if (file.Length > maxExpandedBytes - totalExtractedBytes) {
                throw new OneNoteFormatException("ONENOTE_CAB_EXPANDED_LIMIT", "The extracted .onepkg entries exceed the configured size limit.");
            }
            var bytes = new byte[file.Length];
            if (bytes.Length > 0) Buffer.BlockCopy(source, (int)file.FolderOffset, bytes, 0, bytes.Length);
            totalExtractedBytes += file.Length;
            entries.Add(new OneNoteCabinetEntry(file.Name, bytes));
        }
        return entries.AsReadOnly();
    }

    private static byte[] DecompressFolder(byte[] cabinet, FolderHeader folder, int dataReserve, long maxExpandedBytes) {
        if (folder.DataOffset > int.MaxValue) throw new OneNoteFormatException("ONENOTE_CAB_DATA_OFFSET", "The CAB folder data offset is too large.");
        int offset = (int)folder.DataOffset;
        var blocks = new List<byte[]>(folder.BlockCount);
        var sizes = new List<int>(folder.BlockCount);
        long total = 0;
        for (int index = 0; index < folder.BlockCount; index++) {
            Ensure(cabinet, offset, 8, "CFDATA");
            uint declaredChecksum = ReadUInt32(cabinet, offset);
            int compressedLength = ReadUInt16(cabinet, offset + 4);
            int expandedLength = ReadUInt16(cabinet, offset + 6);
            int dataOffset = checked(offset + 8 + dataReserve);
            Ensure(cabinet, dataOffset, compressedLength, "CFDATA payload");
            if (declaredChecksum != 0) {
                uint actualChecksum = OneNoteCabinetChecksum.Compute(cabinet, offset + 4, checked(4 + dataReserve + compressedLength));
                if (actualChecksum != declaredChecksum) {
                    throw new OneNoteFormatException("ONENOTE_CAB_CHECKSUM", "A CAB data block has an invalid checksum.", offset);
                }
            }
            var block = new byte[compressedLength];
            Buffer.BlockCopy(cabinet, dataOffset, block, 0, block.Length);
            blocks.Add(block);
            sizes.Add(expandedLength);
            total += expandedLength;
            if (total > maxExpandedBytes) throw new OneNoteFormatException("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder exceeds the configured size limit.");
            if (total > int.MaxValue) throw new OneNoteFormatException("ONENOTE_CAB_EXPANDED_LIMIT", "The expanded CAB folder is too large to materialize.");
            offset = checked(dataOffset + compressedLength);
        }

        switch ((ushort)(folder.Compression & CompressionMask)) {
            case CompressionNone: {
                var output = new byte[(int)total];
                int outputOffset = 0;
                for (int index = 0; index < blocks.Count; index++) {
                    if (blocks[index].Length != sizes[index]) throw new OneNoteFormatException("ONENOTE_CAB_UNCOMPRESSED", "An uncompressed CAB block has inconsistent sizes.");
                    Buffer.BlockCopy(blocks[index], 0, output, outputOffset, blocks[index].Length);
                    outputOffset += blocks[index].Length;
                }
                return output;
            }
            case CompressionLzx:
                return OneNoteLzxDecoder.Decompress(blocks, sizes, (folder.Compression >> 8) & 0x1F, maxExpandedBytes);
            case CompressionMsZip:
                throw new OneNoteFormatException("ONENOTE_CAB_MSZIP", "MSZIP-compressed .onepkg archives are not supported yet.");
            default:
                throw new OneNoteFormatException("ONENOTE_CAB_COMPRESSION", "The .onepkg CAB uses an unsupported compression method.");
        }
    }

    private static byte[] ReadExactly(Stream stream, int length) {
        var data = new byte[length];
        int offset = 0;
        while (offset < data.Length) {
            int read = stream.Read(data, offset, data.Length - offset);
            if (read == 0) throw new EndOfStreamException("The .onepkg stream ended unexpectedly.");
            offset += read;
        }
        return data;
    }

    private static int SkipCString(byte[] data, int offset) {
        ReadCString(data, offset, false, out _, out int nextOffset);
        return nextOffset;
    }

    private static void ReadCString(byte[] data, int offset, bool utf8, out string value, out int nextOffset) {
        int end = offset;
        while (end < data.Length && data[end] != 0) end++;
        if (end >= data.Length) throw new OneNoteFormatException("ONENOTE_CAB_STRING", "A CAB string is not null-terminated.");
        if (utf8) value = System.Text.Encoding.UTF8.GetString(data, offset, end - offset);
        else {
            var characters = new char[end - offset];
            for (int index = 0; index < characters.Length; index++) characters[index] = (char)data[offset + index];
            value = new string(characters);
        }
        nextOffset = end + 1;
    }

    private static void Ensure(byte[] data, int offset, int length, string structure) {
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            throw new OneNoteFormatException("ONENOTE_CAB_TRUNCATED", "The .onepkg archive ended inside " + structure + ".", offset);
        }
    }

    private static ushort ReadUInt16(byte[] data, int offset) {
        Ensure(data, offset, 2, "integer");
        return (ushort)(data[offset] | (data[offset + 1] << 8));
    }

    private static uint ReadUInt32(byte[] data, int offset) {
        Ensure(data, offset, 4, "integer");
        return (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24));
    }

    private readonly struct FolderHeader {
        internal FolderHeader(uint dataOffset, ushort blockCount, ushort compression) {
            DataOffset = dataOffset;
            BlockCount = blockCount;
            Compression = compression;
        }
        internal uint DataOffset { get; }
        internal ushort BlockCount { get; }
        internal ushort Compression { get; }
    }

    private readonly struct FileHeader {
        internal FileHeader(uint length, uint folderOffset, ushort folderIndex, string name) {
            Length = length;
            FolderOffset = folderOffset;
            FolderIndex = folderIndex;
            Name = name;
        }
        internal uint Length { get; }
        internal uint FolderOffset { get; }
        internal ushort FolderIndex { get; }
        internal string Name { get; }
    }
}

internal sealed class OneNoteCabinetEntry {
    internal OneNoteCabinetEntry(string name, byte[] data) {
        Name = name;
        Data = data;
    }
    internal string Name { get; }
    internal byte[] Data { get; }
}
