namespace OfficeIMO.Archives.SevenZip;

internal sealed partial class SevenZipArchive {
    private static SevenZipFolder[] ReadStreams(SevenZipInput input, long payloadEnd, int maximum, CancellationToken token) {
        byte type = input.Byte();
        if (type == 0) return Array.Empty<SevenZipFolder>();
        if (type != 6) throw SevenZipInput.Invalid("The 7z stream index is missing PackInfo.");
        long relativeOffset = input.Size();
        if (relativeOffset > payloadEnd - SignatureSize) throw SevenZipInput.Invalid("The 7z packed stream offset is outside the payload region.");
        int packedCount = input.Count(maximum, nameof(ArchiveExtractionOptions.MaxEntryCount));
        input.Expect(9);
        if (packedCount > input.Remaining) throw SevenZipInput.Invalid("The 7z packed size table is truncated.");
        var sizes = new long[packedCount];
        for (int i = 0; i < sizes.Length; i++) {
            token.ThrowIfCancellationRequested();
            sizes[i] = input.Size();
        }
        type = input.Byte();
        uint?[]? packedCrcs = null;
        if (type == 10) {
            packedCrcs = input.Checksums(packedCount);
            type = input.Byte();
        }
        if (type != 0) throw SevenZipInput.Invalid("The 7z PackInfo has an unexpected field.");

        input.Expect(7);
        input.Expect(11);
        int folderCount = input.Count(maximum, nameof(ArchiveExtractionOptions.MaxEntryCount));
        input.RequireInline("folder");
        if (folderCount > input.Remaining / 3) throw SevenZipInput.Invalid("The 7z folder table is truncated.");
        var folders = new SevenZipFolder[folderCount];
        long offset = SignatureSize + relativeOffset;
        var coders = new (int Method, byte[] Properties)[folderCount];
        for (int i = 0; i < folders.Length; i++) {
            token.ThrowIfCancellationRequested();
            coders[i] = SevenZipFolder.ReadCoder(input);
        }
        if (folderCount != packedCount) throw SevenZipInput.Invalid("The 7z packed stream count does not match its simple folders.");
        input.Expect(12);
        for (int i = 0; i < folders.Length; i++) {
            token.ThrowIfCancellationRequested();
            if (sizes[i] > payloadEnd - offset) throw SevenZipInput.Invalid("A 7z packed stream overlaps its header or extends beyond it.");
            long size = input.Size();
            folders[i] = new SevenZipFolder(coders[i].Method, coders[i].Properties, offset, sizes[i], packedCrcs?[i]) { Size = size };
            folders[i].ValidateLayout();
            offset += sizes[i];
        }
        type = input.Byte();
        if (type == 10) {
            uint?[] checksums = input.Checksums(folders.Length);
            for (int i = 0; i < folders.Length; i++) folders[i].Crc = checksums[i];
            type = input.Byte();
        }
        if (type != 0) throw SevenZipInput.Invalid("The 7z UnpackInfo has an unexpected field.");

        type = input.Byte();
        if (type == 8) {
            ReadSubStreams(input, folders, maximum, token);
            type = input.Byte();
        } else {
            foreach (SevenZipFolder folder in folders)
                folder.Extents = new[] { new SevenZipExtent(0, folder.Size, null) };
        }
        if (type != 0) throw SevenZipInput.Invalid("The 7z StreamsInfo has an unexpected field.");
        return folders;
    }

    private static void ReadSubStreams(SevenZipInput input, SevenZipFolder[] folders, int maximum, CancellationToken token) {
        byte type = input.Byte();
        var counts = new int[folders.Length];
        long totalCount = 0;
        for (int i = 0; i < counts.Length; i++) {
            token.ThrowIfCancellationRequested();
            counts[i] = type == 13 ? input.Count(maximum, nameof(ArchiveExtractionOptions.MaxEntryCount)) : 1;
            ArchiveLimits.Add(ref totalCount, counts[i], maximum, nameof(ArchiveExtractionOptions.MaxEntryCount));
        }
        if (type == 13) type = input.Byte();
        bool hasSizes = type == 9;
        int digestCount = 0;
        for (int i = 0; i < folders.Length; i++) {
            token.ThrowIfCancellationRequested();
            SevenZipFolder folder = folders[i];
            int count = counts[i];
            if (count == 0 && folder.Size != 0 || count > 1 && !hasSizes)
                throw SevenZipInput.Invalid("The 7z substream counts cannot cover the declared folder output.");
            if (count > 1 && count - 1 > input.Remaining)
                throw SevenZipInput.Invalid("The 7z substream size table is truncated.");
            var extents = new SevenZipExtent[count];
            long offset = 0;
            for (int j = 0; j < extents.Length; j++) {
                long length = j + 1 == count ? folder.Size - offset : input.Size();
                if (length > folder.Size - offset)
                    throw SevenZipInput.Invalid("The 7z substream sizes exceed their declared folder output.");
                extents[j] = new SevenZipExtent(offset, length, null);
                offset += length;
            }
            folder.Extents = extents;
            if (count != 1 || !folder.Crc.HasValue) digestCount += count;
        }
        if (hasSizes) type = input.Byte();
        if (type == 10) {
            uint?[] digests = input.Checksums(digestCount);
            int digest = 0;
            foreach (SevenZipFolder folder in folders) {
                if (folder.Extents.Length == 1 && folder.Crc.HasValue) continue;
                for (int i = 0; i < folder.Extents.Length; i++) {
                    var extent = folder.Extents[i];
                    folder.Extents[i] = new SevenZipExtent(extent.Offset, extent.Length, digests[digest++]);
                }
            }
            type = input.Byte();
        }
        if (type != 0) throw SevenZipInput.Invalid("The 7z SubStreamsInfo has an unexpected field.");
    }
}
