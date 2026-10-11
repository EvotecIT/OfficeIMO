namespace OfficeIMO.Chm;

internal static class ChmArchiveReader {
    private sealed class DirectoryEntry {
        internal string Name = string.Empty;
        internal string Path = string.Empty;
        internal int Section;
        internal int Offset;
        internal int Length;
    }

    internal static ChmDocument Read(byte[] archive, ChmReadOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (archive.LongLength > options.MaxInputBytes) throw ChmBinary.Error("INPUT_LIMIT", "The CHM archive exceeds MaxInputBytes.");
        ChmBinary.Range(archive, 0, 88);
        if (!ChmBinary.Signature(archive, 0, "ITSF")) throw ChmBinary.Error("SIGNATURE", "The input is not an ITSF compiled HTML help archive.");
        uint version = ChmBinary.U32(archive, 4);
        if (version != 2 && version != 3) throw ChmBinary.Error("VERSION", "Only ITSF versions 2 and 3 are supported.");
        int headerLength = ChmBinary.Index(ChmBinary.U32(archive, 8));
        if (headerLength < (version == 3 ? 96 : 88)) throw ChmBinary.Error("HEADER", "The ITSF header is incomplete.");
        ChmBinary.Range(archive, 0, headerLength);
        int directoryOffset = ChmBinary.Index(ChmBinary.U64(archive, 72));
        int directoryLength = ChmBinary.Index(ChmBinary.U64(archive, 80));
        ChmBinary.Range(archive, directoryOffset, directoryLength);
        int dataOffset = version == 3 ? ChmBinary.Index(ChmBinary.U64(archive, 88)) : checked(directoryOffset + directoryLength);
        if (directoryOffset < headerLength || dataOffset < (long)directoryOffset + directoryLength)
            throw ChmBinary.Error("HEADER", "CHM header, directory, and content ranges overlap.");
        ChmBinary.Range(archive, dataOffset, 0);
        IReadOnlyList<DirectoryEntry> directory = ReadDirectory(archive, directoryOffset, directoryLength, options, token);
        var sectionZero = new Dictionary<string, ChmEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (DirectoryEntry entry in directory.Where(entry => entry.Section == 0)) {
            long absoluteOffset = (long)dataOffset + entry.Offset;
            ChmBinary.Range(archive, absoluteOffset, entry.Length);
            sectionZero.Add(entry.Path, new ChmEntry(entry.Name, entry.Path, 0, archive, (int)absoluteOffset, entry.Length));
        }
        byte[] expanded = Array.Empty<byte>();
        long logicalExpandedLength = 0;
        if (directory.Any(entry => entry.Section == 1))
            expanded = Decompress(sectionZero, options, token, out logicalExpandedLength);
        var entries = new List<ChmEntry>(directory.Count);
        foreach (DirectoryEntry entry in directory) {
            token.ThrowIfCancellationRequested();
            if (entry.Section == 0) { entries.Add(sectionZero[entry.Path]); continue; }
            if ((long)entry.Offset + entry.Length > logicalExpandedLength)
                throw ChmBinary.Error("ENTRY_BOUNDS", "An entry extends beyond the expanded CHM section: " + entry.Path);
            entries.Add(new ChmEntry(entry.Name, entry.Path, 1, expanded, entry.Offset, entry.Length));
        }
        return new ChmDocument(entries.AsReadOnly(), options, version, ChmBinary.U32(archive, 20), token);
    }

    private static IReadOnlyList<DirectoryEntry> ReadDirectory(byte[] archive, int offset, int length, ChmReadOptions options, CancellationToken token) {
        ChmBinary.Range(archive, offset, 84);
        if (length < 84 || !ChmBinary.Signature(archive, offset, "ITSP") || ChmBinary.U32(archive, offset + 4) != 1)
            throw ChmBinary.Error("DIRECTORY", "The CHM directory header is invalid or unsupported.");
        int headerLength = ChmBinary.Index(ChmBinary.U32(archive, offset + 8));
        int blockLength = ChmBinary.Index(ChmBinary.U32(archive, offset + 16));
        int count = ChmBinary.Index(ChmBinary.U32(archive, offset + 44));
        if (headerLength < 84 || headerLength > length || blockLength < 32 || blockLength > 1024 * 1024 ||
            count > (length - headerLength) / blockLength)
            throw ChmBinary.Error("DIRECTORY", "CHM directory chunk dimensions are invalid.");
        var result = new List<DirectoryEntry>();
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var utf8 = new UTF8Encoding(false, true);
        int leaves = 0;
        for (int block = 0; block < count; block++) {
            token.ThrowIfCancellationRequested();
            int start = checked(offset + headerLength + block * blockLength);
            if (ChmBinary.Signature(archive, start, "PMGI")) continue; // Listing chunks are read once; the lookup tree is unnecessary.
            if (!ChmBinary.Signature(archive, start, "PMGL")) throw ChmBinary.Error("DIRECTORY", "Unknown CHM directory chunk type.");
            leaves++;
            int free = ChmBinary.Index(ChmBinary.U32(archive, start + 4));
            if (free < 2 || free > blockLength - 20) throw ChmBinary.Error("DIRECTORY", "Invalid free-space length in a listing chunk.");
            int end = start + blockLength - free;
            int position = start + 20;
            int blockEntries = 0;
            while (position < end) {
                token.ThrowIfCancellationRequested();
                if (result.Count >= options.MaxEntries) throw ChmBinary.Error("ENTRY_LIMIT", "The CHM directory exceeds MaxEntries.");
                int nameLength = ChmBinary.Index(ChmBinary.EncInt(archive, ref position, end));
                if (nameLength < 1 || nameLength > options.MaxPathLength || nameLength > end - position)
                    throw ChmBinary.Error("DIRECTORY", "A CHM entry name is empty, truncated, or too long.");
                string name;
                try { name = utf8.GetString(archive, position, nameLength); }
                catch (DecoderFallbackException exception) { throw new ChmReadException("CHM_DIRECTORY", "An entry name is not valid UTF-8.", exception); }
                position += nameLength;
                int section = ChmBinary.Index(ChmBinary.EncInt(archive, ref position, end));
                int entryOffset = ChmBinary.Index(ChmBinary.EncInt(archive, ref position, end));
                int entryLength = ChmBinary.Index(ChmBinary.EncInt(archive, ref position, end));
                if (section != 0 && section != 1) throw ChmBinary.Error("SECTION_UNSUPPORTED", "The CHM uses an unsupported storage section.");
                if (entryLength > options.MaxEntryBytes) throw ChmBinary.Error("ENTRY_LIMIT", "An entry exceeds MaxEntryBytes: " + name);
                string path = ChmPaths.DirectoryPath(name);
                if (!paths.Add(path)) throw ChmBinary.Error("DUPLICATE_PATH", "Ambiguous case-insensitive CHM entry path: " + name);
                if (path.EndsWith("/", StringComparison.Ordinal) && entryLength != 0)
                    throw ChmBinary.Error("DIRECTORY", "A directory marker contains data.");
                result.Add(new DirectoryEntry { Name = name, Path = path, Section = section, Offset = entryOffset, Length = entryLength });
                blockEntries++;
            }
            if (blockEntries != ChmBinary.U16(archive, start + blockLength - 2))
                throw ChmBinary.Error("DIRECTORY", "A listing chunk's entry count disagrees with its records.");
        }
        if (leaves == 0) throw ChmBinary.Error("DIRECTORY", "The CHM has no listing chunks.");
        return result;
    }

    private static byte[] Decompress(Dictionary<string, ChmEntry> entries, ChmReadOptions options, CancellationToken token, out long logicalLength) {
        const string prefix = "::DataSpace/Storage/MSCompressed/";
        ChmEntry Require(string path) {
            if (!entries.TryGetValue(prefix + path, out ChmEntry? entry)) throw ChmBinary.Error("COMPRESSION_METADATA", "Required uncompressed LZX metadata is missing: " + path);
            return entry;
        }
        byte[] control = Require("ControlData").GetBytes();
        ChmBinary.Range(control, 0, 24);
        uint version = ChmBinary.U32(control, 8);
        if (!ChmBinary.Signature(control, 4, "LZXC") || (version != 1 && version != 2))
            throw ChmBinary.Error("COMPRESSION_UNSUPPORTED", "Only LZXC control versions 1 and 2 are supported.");
        ulong scale = version == 2 ? 32768UL : 1;
        ulong resetInterval = ChmBinary.U32(control, 12) * scale;
        ulong window = ChmBinary.U32(control, 16) * scale;
        uint windowsPerReset = ChmBinary.U32(control, 20);
        if (window < 32768 || window > 2097152 || (window & (window - 1)) != 0 || resetInterval == 0 ||
            resetInterval % (window / 2) != 0 || windowsPerReset == 0)
            throw ChmBinary.Error("COMPRESSION_METADATA", "Invalid LZX window or reset interval.");
        ulong resetWindows = resetInterval / (window / 2);
        if (resetWindows > (ulong)int.MaxValue / windowsPerReset) throw ChmBinary.Error("COMPRESSION_METADATA", "The LZX reset interval is out of range.");
        int resetFrames = (int)(resetWindows * windowsPerReset);
        if (resetFrames < 1) throw ChmBinary.Error("COMPRESSION_METADATA", "The LZX reset interval is empty.");
        int windowBits = 15;
        while ((1UL << windowBits) < window) windowBits++;
        byte[] table = Require("Transform/{7FC28940-9D31-11D0-9B27-00A0C91E9C7C}/InstanceData/ResetTable").GetBytes();
        ChmBinary.Range(table, 0, 40);
        if (ChmBinary.U32(table, 0) != 2 || ChmBinary.U32(table, 8) != 8 || ChmBinary.U64(table, 32) != 32768)
            throw ChmBinary.Error("COMPRESSION_UNSUPPORTED", "The CHM reset table uses an unsupported layout or frame size.");
        int frames = ChmBinary.Index(ChmBinary.U32(table, 4));
        int tableStart = ChmBinary.Index(ChmBinary.U32(table, 12));
        ulong declared = ChmBinary.U64(table, 16);
        ulong compressedLength = ChmBinary.U64(table, 24);
        if (declared > (ulong)options.MaxExpandedBytes)
            throw ChmBinary.Error("EXPANDED_LIMIT", "The LZX section exceeds MaxExpandedBytes.");
        ulong padded = ((declared + 32767) / 32768) * 32768;
        if (declared > (ulong)options.MaxExpandedBytes || padded > (ulong)options.MaxExpandedBytes)
            throw ChmBinary.Error("EXPANDED_LIMIT", "The LZX section exceeds MaxExpandedBytes.");
        if (tableStart < 40 || frames != (int)(padded / 32768)) throw ChmBinary.Error("COMPRESSION_METADATA", "The LZX frame count does not match its expanded length.");
        ChmBinary.Range(table, tableStart, (long)frames * 8);
        byte[] span = Require("SpanInfo").GetBytes();
        if (span.Length != 8 || ChmBinary.U64(span, 0) != declared) throw ChmBinary.Error("COMPRESSION_METADATA", "The LZX span and reset table disagree.");
        ChmEntry content = Require("Content");
        if (compressedLength > (ulong)content.Length) throw ChmBinary.Error("COMPRESSION_METADATA", "The compressed LZX stream is truncated.");
        var offsets = new int[frames + 1];
        offsets[frames] = ChmBinary.Index(compressedLength);
        for (int i = 0; i < frames; i++) {
            offsets[i] = ChmBinary.Index(ChmBinary.U64(table, checked(tableStart + i * 8)));
            if (offsets[i] >= offsets[frames] || (i == 0 ? offsets[i] != 0 : offsets[i] <= offsets[i - 1]))
                throw ChmBinary.Error("COMPRESSION_METADATA", "LZX frame offsets are invalid or non-increasing.");
        }
        var result = new byte[(int)padded];
        using Stream input = content.OpenRead();
        for (int first = 0; first < frames;) {
            token.ThrowIfCancellationRequested();
            int count = Math.Min(resetFrames, frames - first);
            // Bound each reset group's temporary output as well as the retained section.
            var chunks = new List<byte[]>(count);
            var sizes = new List<int>(count);
            for (int frame = first; frame < first + count; frame++) {
                int compressed = offsets[frame + 1] - offsets[frame];
                var chunk = new byte[compressed];
                input.Position = offsets[frame];
                int read = 0;
                while (read < compressed) { token.ThrowIfCancellationRequested(); int n = input.Read(chunk, read, compressed - read); if (n == 0) throw ChmBinary.Error("TRUNCATED", "A compressed frame is truncated."); read += n; }
                chunks.Add(chunk); sizes.Add(32768);
            }
            byte[] decoded;
            try { decoded = OfficeLzxDecoder.Decompress(chunks, sizes, windowBits, options.MaxExpandedBytes, token, requireCompleteBlocks: false); }
            catch (OfficeLzxException exception) { throw new ChmReadException("CHM_" + exception.Code, exception.Message, exception); }
            Buffer.BlockCopy(decoded, 0, result, first * 32768, decoded.Length);
            first += count;
        }
        logicalLength = (long)declared;
        return result;
    }
}
