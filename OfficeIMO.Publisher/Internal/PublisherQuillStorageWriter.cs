namespace OfficeIMO.Publisher.Internal;

/// <summary>Publishes planned formatting pages, resource replacements and descriptor chains while retaining unrelated source bytes.</summary>
internal static class PublisherQuillStorageWriter {
    internal static byte[]? TryWrite(PublisherQuillSourceLayout layout, byte[] text,
        IReadOnlyList<byte[]> pc, IReadOnlyList<byte[]> pp, IReadOnlyList<int> storyLengths,
        IReadOnlyDictionary<string, byte[]> resources, long maximumBytes, CancellationToken token) {
        PublisherBinaryData data = layout.Data; byte[] source = data.Bytes;
        int count = checked(layout.Chunks.Count - layout.CharacterPages.Length - layout.ParagraphPages.Length + pc.Count + pp.Count);
        if (pc.Count == 0 || pp.Count == 0 || count > PublisherQuillSourceLayout.MaximumDescriptors
            || !KnownPages(layout.CharacterPages) || !KnownPages(layout.ParagraphPages)) return null;
        int shift = checked((Math.Max(0, text.Length - layout.Text.Length) + 511) / 512 * 512);
        long cursor = (long)source.Length + shift; CheckSize(cursor);
        int[] characterOffsets = PlanPages(layout.CharacterPages, pc.Count), paragraphOffsets = PlanPages(layout.ParagraphPages, pp.Count);
        PublisherQuillChunk btec = layout.Find("BTEC"), btep = layout.Find("BTEP");
        int characterIndexBytes = checked(16 + pc.Count * 8), paragraphIndexBytes = checked(16 + pp.Count * 8);
        int characterIndex = characterIndexBytes <= btec.Length ? btec.Offset + shift : Append(characterIndexBytes, false);
        int paragraphIndex = paragraphIndexBytes <= btep.Length ? btep.Offset + shift : Append(paragraphIndexBytes, false);
        var replacements = new Dictionary<PublisherQuillChunk, (int Offset, byte[] Bytes)>();
        foreach (var resource in resources) {
            token.ThrowIfCancellationRequested();
            if (resource.Key is not ("FONT" or "PL  ")) return null;
            PublisherQuillChunk original = layout.Find(resource.Key);
            replacements.Add(original, (Append(resource.Value.Length, false), resource.Value));
        }
        var rows = new List<PublisherQuillOutputDescriptor>();
        foreach (PublisherQuillChunk chunk in layout.Chunks) {
            token.ThrowIfCancellationRequested();
            if (chunk.Tag is "FDPC" or "FDPP") {
                if (chunk.Id != 0) continue;
                PublisherQuillChunk[] original = chunk.Tag == "FDPC" ? layout.CharacterPages : layout.ParagraphPages;
                int[] offsets = chunk.Tag == "FDPC" ? characterOffsets : paragraphOffsets;
                for (int i = 0; i < offsets.Length; i++) rows.Add(new PublisherQuillOutputDescriptor(original[Math.Min(i, original.Length - 1)], offsets[i], 512, i));
            } else {
                int offset = chunk == layout.Text ? layout.Text.Offset : chunk.Offset + shift;
                int length = chunk == layout.Text ? text.Length : chunk.Length;
                if (chunk == btec) { offset = characterIndex; length = characterIndexBytes; }
                if (chunk == btep) { offset = paragraphIndex; length = paragraphIndexBytes; }
                if (replacements.TryGetValue(chunk, out var resource)) { offset = resource.Offset; length = resource.Bytes.Length; }
                rows.Add(new PublisherQuillOutputDescriptor(chunk, offset, length, chunk.Id));
            }
        }
        if (rows.Count != count) return null;
        PublisherQuillDirectoryWriter directories = PublisherQuillDirectoryWriter.Plan(layout, count, shift, () => Append(512, true), token);
        token.ThrowIfCancellationRequested(); var output = new byte[(int)cursor];
        Buffer.BlockCopy(source, 0, output, 0, layout.Text.Offset);
        Buffer.BlockCopy(text, 0, output, layout.Text.Offset, text.Length);
        Buffer.BlockCopy(source, layout.Text.End, output, layout.Text.End + shift, source.Length - layout.Text.End);
        WritePages(layout.CharacterPages, pc, characterOffsets); WritePages(layout.ParagraphPages, pp, paragraphOffsets);
        WriteIndex(btec, pc, characterOffsets, characterIndex); WriteIndex(btep, pp, paragraphOffsets, paragraphIndex);
        foreach (var resource in replacements) {
            token.ThrowIfCancellationRequested(); Array.Clear(output, resource.Key.Offset + shift, resource.Key.Length);
            Buffer.BlockCopy(resource.Value.Bytes, 0, output, resource.Value.Offset, resource.Value.Bytes.Length);
        }
        directories.Publish(output, rows, token);
        for (int i = 0; i < storyLengths.Count; i++) Put(layout.StoryLengths.Offset + shift + 12 + i * 4, (uint)storyLengths[i]);
        token.ThrowIfCancellationRequested(); return output;

        bool KnownPages(PublisherQuillChunk[] pages) => pages.All(page => data.U32(layout.DescriptorOffsets[page] + 8) == 1
            && data.Tag(layout.DescriptorOffsets[page] + 12) == page.Tag);
        void CheckSize(long bytes) { if (bytes > maximumBytes || bytes > int.MaxValue) throw new InvalidDataException("Publisher native output byte limit exceeded."); }
        int Append(int bytes, bool align) {
            long start = align ? checked((cursor + 511) / 512 * 512) : cursor;
            cursor = checked(start + bytes); CheckSize(cursor); return (int)start;
        }
        int[] PlanPages(PublisherQuillChunk[] original, int pages) {
            var offsets = new int[pages];
            for (int i = 0; i < pages; i++) { token.ThrowIfCancellationRequested(); offsets[i] = i < original.Length ? original[i].Offset + shift : Append(512, true); }
            return offsets;
        }
        void WritePages(PublisherQuillChunk[] original, IReadOnlyList<byte[]> pages, int[] offsets) {
            foreach (PublisherQuillChunk page in original) Array.Clear(output, page.Offset + shift, page.Length);
            for (int i = 0; i < pages.Count; i++) {
                token.ThrowIfCancellationRequested(); Buffer.BlockCopy(pages[i], 0, output, offsets[i], 512);
                if (i != 0) Put(offsets[i] + 4, LastEnd(pages[i - 1]));
            }
        }
        void WriteIndex(PublisherQuillChunk original, IReadOnlyList<byte[]> pages, int[] offsets, int target) {
            Array.Clear(output, original.Offset + shift, original.Length);
            Put(target, (uint)pages.Count); Put(target + 4, 4); Put(target + 8, 0); Put(target + 12, 0);
            for (int i = 0; i < pages.Count; i++) { Put(target + 16 + i * 4, LastEnd(pages[i])); Put(target + 16 + pages.Count * 4 + i * 4, (uint)offsets[i]); }
        }
        void Put(int offset, uint value) { for (int i = 0; i < 4; i++) output[offset + i] = (byte)(value >> (i * 8)); }
    }

    private static uint LastEnd(byte[] page) => BitConverter.ToUInt32(page, 8 + (BitConverter.ToUInt16(page, 0) - 1) * 4);
}
