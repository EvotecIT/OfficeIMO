namespace OfficeIMO.Publisher.Internal;

/// <summary>Changes active formatting-page allocation within the admitted directory, preserving unrelated bytes and relocating enlarged indexes.</summary>
internal static class PublisherQuillPageAllocator {
    internal static byte[]? TryWrite(PublisherQuillSourceLayout layout, byte[] text,
        IReadOnlyList<byte[]> pc, IReadOnlyList<byte[]> pp, IReadOnlyList<int> storyLengths,
        long maximumBytes, CancellationToken token) {
        PublisherBinaryData data = layout.Data; byte[] source = data.Bytes;
        int descriptorCount = layout.Chunks.Count - layout.CharacterPages.Length - layout.ParagraphPages.Length + pc.Count + pp.Count;
        if (pc.Count == 0 || pp.Count == 0 || descriptorCount > PublisherQuillSourceLayout.MaximumDescriptors) return null;
        var descriptors = layout.Chunks.Select((chunk, index) => (chunk, offset: 32 + index * 24)).ToDictionary(item => item.chunk, item => item.offset);
        if (descriptorCount > layout.Chunks.Count)
            for (int i = 32 + layout.Chunks.Count * 24; i < 32 + descriptorCount * 24; i++) if (source[i] != 0) return null;
        if (!KnownPageMetadata(layout.CharacterPages) || !KnownPageMetadata(layout.ParagraphPages)) return null;
        int shift = checked((Math.Max(0, text.Length - layout.Text.Length) + 511) / 512 * 512);
        long cursor = (long)source.Length + shift;
        CheckSize(cursor);
        int[] characterOffsets = PlanPages(layout.CharacterPages, pc.Count), paragraphOffsets = PlanPages(layout.ParagraphPages, pp.Count);
        PublisherQuillChunk btec = layout.Find("BTEC"), btep = layout.Find("BTEP");
        int characterIndexBytes = checked(16 + pc.Count * 8), paragraphIndexBytes = checked(16 + pp.Count * 8);
        int characterIndex = characterIndexBytes <= btec.Length ? btec.Offset + shift : Append(characterIndexBytes, align: false);
        int paragraphIndex = paragraphIndexBytes <= btep.Length ? btep.Offset + shift : Append(paragraphIndexBytes, align: false);
        token.ThrowIfCancellationRequested(); var output = new byte[(int)cursor];
        Buffer.BlockCopy(source, 0, output, 0, layout.Text.Offset);
        Buffer.BlockCopy(text, 0, output, layout.Text.Offset, text.Length);
        Buffer.BlockCopy(source, layout.Text.End, output, layout.Text.End + shift, source.Length - layout.Text.End);
        WritePages(layout.CharacterPages, pc, characterOffsets);
        WritePages(layout.ParagraphPages, pp, paragraphOffsets);
        WriteIndex(btec, pc, characterOffsets, characterIndex);
        WriteIndex(btep, pp, paragraphOffsets, paragraphIndex);
        int position = 32;
        foreach (PublisherQuillChunk chunk in layout.Chunks) {
            token.ThrowIfCancellationRequested();
            if (chunk.Tag is "FDPC" or "FDPP") {
                if (chunk.Id != 0) continue;
                PublisherQuillChunk[] original = chunk.Tag == "FDPC" ? layout.CharacterPages : layout.ParagraphPages;
                int[] offsets = chunk.Tag == "FDPC" ? characterOffsets : paragraphOffsets;
                for (int i = 0; i < offsets.Length; i++) {
                    WriteDescriptor(original[Math.Min(i, original.Length - 1)], offsets[i], 512, i);
                }
            } else {
                int offset = chunk == layout.Text ? layout.Text.Offset : chunk.Offset + shift, length = chunk == layout.Text ? text.Length : chunk.Length;
                if (chunk == btec) { offset = characterIndex; length = characterIndexBytes; }
                if (chunk == btep) { offset = paragraphIndex; length = paragraphIndexBytes; }
                WriteDescriptor(chunk, offset, length, chunk.Id);
            }
        }
        if (position < 32 + layout.Chunks.Count * 24) Array.Clear(output, position, 32 + layout.Chunks.Count * 24 - position);
        Put16(12, descriptorCount); Put16(26, descriptorCount); Put(20, (uint)output.Length);
        for (int i = 0; i < storyLengths.Count; i++) Put(layout.StoryLengths.Offset + shift + 12 + i * 4, (uint)storyLengths[i]);
        token.ThrowIfCancellationRequested(); return output;

        bool KnownPageMetadata(PublisherQuillChunk[] pages) => pages.All(page => data.U32(descriptors[page] + 8) == 1 && data.Tag(descriptors[page] + 12) == page.Tag);
        void CheckSize(long length) { if (length > maximumBytes || length > int.MaxValue) throw new InvalidDataException("Publisher native output byte limit exceeded."); }
        int Append(int bytes, bool align) {
            long start = align ? checked((cursor + 511) / 512 * 512) : cursor;
            cursor = checked(start + bytes); CheckSize(cursor); return (int)start;
        }
        int[] PlanPages(PublisherQuillChunk[] original, int count) {
            var offsets = new int[count];
            for (int i = 0; i < count; i++) {
                token.ThrowIfCancellationRequested(); offsets[i] = i < original.Length ? original[i].Offset + shift : Append(512, align: true);
            }
            return offsets;
        }
        void WritePages(PublisherQuillChunk[] original, IReadOnlyList<byte[]> pages, int[] offsets) {
            for (int i = 0; i < original.Length; i++) Array.Clear(output, original[i].Offset + shift, original[i].Length);
            for (int i = 0; i < pages.Count; i++) {
                token.ThrowIfCancellationRequested(); Buffer.BlockCopy(pages[i], 0, output, offsets[i], 512);
                if (i != 0) Put(offsets[i] + 4, LastEnd(pages[i - 1]));
            }
        }
        void WriteIndex(PublisherQuillChunk original, IReadOnlyList<byte[]> pages, int[] offsets, int target) {
            Array.Clear(output, original.Offset + shift, original.Length);
            Put(target, (uint)pages.Count); Put(target + 4, 4); Put(target + 8, 0); Put(target + 12, 0);
            for (int i = 0; i < pages.Count; i++) {
                Put(target + 16 + i * 4, LastEnd(pages[i])); Put(target + 16 + pages.Count * 4 + i * 4, (uint)offsets[i]);
            }
        }
        void WriteDescriptor(PublisherQuillChunk original, int offset, int length, int id) {
            Buffer.BlockCopy(source, descriptors[original], output, position, 24);
            Put16(position + 6, id); Put(position + 16, (uint)offset); Put(position + 20, (uint)length); position += 24;
        }
        void Put(int offset, uint value) { for (int i = 0; i < 4; i++) output[offset + i] = (byte)(value >> (i * 8)); }
        void Put16(int offset, int value) { output[offset] = (byte)value; output[offset + 1] = (byte)(value >> 8); }
    }

    private static uint LastEnd(byte[] page) => BitConverter.ToUInt32(page, 8 + (BitConverter.ToUInt16(page, 0) - 1) * 4);

    internal static IReadOnlyList<byte[]>? PackMinimum(IReadOnlyList<PublisherQuillFormattingRange> ranges, int maximumPages, CancellationToken token) {
        var pages = new List<byte[]>(); int cursor = 0;
        while (cursor < ranges.Count) {
            if (pages.Count == maximumPages) return null;
            var current = new List<PublisherQuillFormattingRange>(); byte[]? accepted = null;
            while (cursor < ranges.Count) {
                token.ThrowIfCancellationRequested(); current.Add(ranges[cursor]);
                byte[]? candidate = PublisherQuillFormattingPage.TryWriteShared(current);
                if (candidate == null) break;
                accepted = candidate; cursor++;
            }
            if (accepted == null) return null;
            pages.Add(accepted);
        }
        return pages.Count == 0 ? null : pages;
    }
}
