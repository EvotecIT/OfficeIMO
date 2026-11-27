namespace OfficeIMO.Archives.SevenZip;

internal sealed partial class SevenZipArchive {
    private static Member[] ReadMembers(SevenZipInput input, SevenZipFolder[] folders, ArchiveLimits limits, CancellationToken token) {
        int count = input.Count(limits.Entries, nameof(ArchiveExtractionOptions.MaxEntryCount));
        if (count > input.Remaining / 2) throw SevenZipInput.Invalid("The 7z file name table is truncated.");
        limits.Operation.AdmitEntries(count);
        var properties = new Dictionary<byte, SevenZipInput>();
        while (true) {
            token.ThrowIfCancellationRequested();
            byte type = input.Byte();
            if (type == 0) break;
            SevenZipInput field = input.Property();
            switch (type) {
                case 0x18: throw new NotSupportedException("7z file start positions for split members are unsupported.");
                case 0x0e: case 0x0f: case 0x10: case 0x11: case 0x14: case 0x15:
                    if (properties.ContainsKey(type)) throw SevenZipInput.Invalid("The 7z file table contains duplicate metadata.");
                    properties.Add(type, field);
                    break;
            }
        }
        bool[]? emptyStreams = ReadFlags(properties, 0x0e, count);
        int emptyCount = emptyStreams?.Count(empty => empty) ?? 0;
        bool[]? emptyFiles = ReadFlags(properties, 0x0f, emptyCount);
        if (ReadFlags(properties, 0x10, emptyCount)?.Any(anti => anti) == true)
            throw new NotSupportedException("7z anti-items are unsupported.");
        if (!properties.TryGetValue(0x11, out SevenZipInput? names)) {
            if (count == 0) return Array.Empty<Member>();
            throw SevenZipInput.Invalid("The 7z file table has no names.");
        }
        names.RequireInline("name");
        if (count > names.Remaining / 2) throw SevenZipInput.Invalid("The 7z file name table is truncated.");
        DateTimeOffset?[]? stamps = properties.TryGetValue(0x14, out SevenZipInput? times)
            ? ReadTimes(times, count, token) : null;
        uint[]? attributes = properties.TryGetValue(0x15, out SevenZipInput? attrs)
            ? ReadAttributes(attrs, count, token) : null;
        long availableFiles = folders.Sum(folder => (long)folder.Extents.Length);
        if (availableFiles != count - emptyCount) throw SevenZipInput.Invalid("The 7z file table and unpack stream counts disagree.");

        var paths = new ArchivePaths(limits, limits.Metadata);
        var members = new Member[count];
        int emptyIndex = 0, folderIndex = 0, extentIndex = 0;
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            bool empty = emptyStreams?[i] == true;
            bool directory = empty && emptyFiles?[emptyIndex] != true;
            if (empty) emptyIndex++;
            string name = names.Name(limits.PathLength, token);
            ArchiveEntryKind kind = Kind(attributes?[i] ?? 0, directory);
            SevenZipExtent extent = default;
            SevenZipFolder? folder = null;
            if (!empty) {
                while (extentIndex >= folders[folderIndex].Extents.Length) {
                    folderIndex++;
                    extentIndex = 0;
                }
                folder = folders[folderIndex];
                extent = folder.Extents[extentIndex++];
                limits.CheckEntry(extent.Length);
            }
            string path;
            if (kind == ArchiveEntryKind.Directory) path = paths.AddDirectory(name);
            else path = paths.Add(name);
            var entry = new SevenZipArchiveEntry(i, name, path, kind, extent.Length, stamps?[i],
                empty ? null : folderIndex, folder?.Compression ?? ArchiveCompression.None);
            members[i] = new Member(entry, empty ? -1 : folderIndex, extent.Offset);
        }
        names.Finish();
        return members;
    }

    private static ArchiveEntryKind Kind(uint attributes, bool directory) {
        // The high word has a POSIX mode only when the 7z Unix extension flag is set.
        uint posixKind = (attributes & 0x8000) != 0 ? attributes >> 16 & 0xf000 : 0;
        if (posixKind == 0xa000) return ArchiveEntryKind.SymbolicLink;
        if ((attributes & 0x400) != 0 || posixKind is not (0 or 0x4000 or 0x8000))
            return ArchiveEntryKind.Special;
        if (!directory && ((attributes & 0x10) != 0 || posixKind == 0x4000))
            throw SevenZipInput.Invalid("A 7z file table record has contradictory directory flags.");
        return directory ? ArchiveEntryKind.Directory : ArchiveEntryKind.File;
    }

    private static bool[]? ReadFlags(Dictionary<byte, SevenZipInput> properties, byte type, int count) {
        if (!properties.TryGetValue(type, out SevenZipInput? field)) return null;
        bool[] flags = field.Bits(count);
        field.Finish();
        return flags;
    }

    private static uint[] ReadAttributes(SevenZipInput data, int count, CancellationToken token) {
        bool[]? present = data.Defined(count);
        data.RequireInline("attribute");
        var values = new uint[count];
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            if (present == null || present[i]) values[i] = data.U32();
        }
        data.Finish();
        return values;
    }

    private static DateTimeOffset?[] ReadTimes(SevenZipInput data, int count, CancellationToken token) {
        bool[]? present = data.Defined(count);
        data.RequireInline("modification time");
        var values = new DateTimeOffset?[count];
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            if (present != null && !present[i]) continue;
            long ticks = data.I64();
            // FILETIME spans a larger unsigned range than DateTimeOffset can represent.
            if (ticks is >= 0 and <= 2650467743999999999L) values[i] = new DateTimeOffset(DateTime.FromFileTimeUtc(ticks));
        }
        data.Finish();
        return values;
    }
}
