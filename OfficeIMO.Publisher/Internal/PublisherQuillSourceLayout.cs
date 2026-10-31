namespace OfficeIMO.Publisher.Internal;

/// <summary>Admits known Quill directory chains and verifies all payload, directory and formatting-page ranges before relocation.</summary>
internal sealed class PublisherQuillSourceLayout {
    internal const int MaximumDescriptors = ushort.MaxValue;
    internal const int DescriptorsPerDirectory = 20;
    private PublisherQuillSourceLayout(PublisherBinaryData data, IReadOnlyList<PublisherQuillChunk> chunks,
        IReadOnlyList<PublisherQuillSourceDirectory> directories, IReadOnlyDictionary<PublisherQuillChunk, int> descriptors) {
        Data = data; Chunks = chunks;
        Directories = directories; DescriptorOffsets = descriptors;
        Text = Find("TEXT"); StoryIds = Find("SYID"); StoryLengths = Find("STRS");
        CharacterPages = chunks.Where(chunk => chunk.Tag == "FDPC").ToArray();
        ParagraphPages = chunks.Where(chunk => chunk.Tag == "FDPP").ToArray();
        IsMinimalProfile = chunks.Count(chunk => chunk.Tag == "STSH") == 2;
    }
    internal PublisherBinaryData Data { get; }
    internal IReadOnlyList<PublisherQuillChunk> Chunks { get; }
    internal IReadOnlyList<PublisherQuillSourceDirectory> Directories { get; }
    internal IReadOnlyDictionary<PublisherQuillChunk, int> DescriptorOffsets { get; }
    internal PublisherQuillChunk Text { get; }
    internal PublisherQuillChunk StoryIds { get; }
    internal PublisherQuillChunk StoryLengths { get; }
    internal PublisherQuillChunk[] CharacterPages { get; }
    internal PublisherQuillChunk[] ParagraphPages { get; }
    /// <summary>Known compact story profile without optional native auxiliary libraries.</summary>
    internal bool IsMinimalProfile { get; }
    internal PublisherQuillChunk Find(string tag) => Chunks.Single(chunk => chunk.Tag == tag);

    internal static PublisherQuillSourceLayout? TryRead(byte[] source, PublisherParseContext context) {
        var data = new PublisherBinaryData(source, "Quill/QuillSub/CONTENTS");
        IReadOnlyList<PublisherQuillChunk> chunks = PublisherQuillDirectory.Read(data, context);
        int pc = chunks.Count(chunk => chunk.Tag == "FDPC"), pp = chunks.Count(chunk => chunk.Tag == "FDPP");
        if (pc < 1 || pp < 1 || chunks.Count > MaximumDescriptors) return null;
        string[] expected = new[] { "TEXT", "STSH", "STSH", "STSH" }.Concat(Enumerable.Repeat("FDPP", pp))
            .Concat(Enumerable.Repeat("FDPC", pc)).Concat(new[] { "SYID", "SGP ", "INK ", "BTEP", "BTEC", "FONT" })
            .Concat(Enumerable.Repeat("TCD ", chunks.Count(chunk => chunk.Tag == "TCD "))).Concat(new[] { "STRS", "MCLD", "PL  " }).ToArray();
        string[] minimal = new[] { "TEXT", "SYID", "STRS", "STSH", "STSH", "INK ", "FONT" }
            .Concat(Enumerable.Repeat("FDPC", pc)).Concat(Enumerable.Repeat("FDPP", pp)).Concat(new[] { "BTEC", "BTEP" }).ToArray();
        bool compact = chunks.Select(chunk => chunk.Tag).SequenceEqual(minimal);
        if ((!compact && !chunks.Select(chunk => chunk.Tag).SequenceEqual(expected))
            || !chunks.Where(chunk => chunk.Tag == "STSH").Select(chunk => chunk.Id).SequenceEqual(compact ? new[] { 0, 1 } : new[] { 0, 1, 2 })
            || chunks.Where(chunk => chunk.Tag is not ("STSH" or "FDPC" or "FDPP" or "TCD ")).Any(chunk => chunk.Id != 0)
            || chunks.Where(chunk => chunk.Tag == "TCD ").Select(chunk => chunk.Id).Distinct().Count() != chunks.Count(chunk => chunk.Tag == "TCD ")
            || data.U16(8) != 4 || data.U16(10) != 7 || data.U16(12) != chunks.Count
            || data.U32(20) != source.Length)
            return null;
        PublisherQuillChunk[] physical = chunks.OrderBy(chunk => chunk.Offset).ToArray();
        if (physical[0] != chunks[0] || chunks[0].Offset != 512 || (chunks[0].Length & 1) != 0 || physical.Any(chunk => chunk.Length <= 0)) return null;
        for (int i = 1; i < physical.Length; i++) if (physical[i].Offset < physical[i - 1].End) return null;
        var directories = new List<PublisherQuillSourceDirectory>(); var descriptors = new Dictionary<PublisherQuillChunk, int>();
        uint next = 24; int ordinal = 0;
        while (next != uint.MaxValue) {
            context.Record(); int offset = data.Offset(next), page = offset == 24 ? 0 : offset;
            data.Range(page, 512);
            int count = data.U16(offset + 2);
            if (offset != 24 && offset % 512 != 0 || count is < 1 or > DescriptorsPerDirectory
                || data.U16(offset) != 504 || physical.Any(chunk => chunk.Offset < page + 512 && chunk.End > page)
                || directories.Any(directory => directory.PageOffset == page)) return null;
            directories.Add(new PublisherQuillSourceDirectory(offset, count));
            for (int i = 0; i < count; i++) descriptors.Add(chunks[ordinal++], offset + 8 + i * 24);
            next = data.U32(offset + 4);
        }
        if (ordinal != chunks.Count || data.U32(16) != directories.Count * 512L) return null;
        var layout = new PublisherQuillSourceLayout(data, chunks, directories, descriptors);
        return layout.CheckPages(layout.CharacterPages, "BTEC", context) && layout.CheckPages(layout.ParagraphPages, "BTEP", context) ? layout : null;
    }

    private bool CheckPages(PublisherQuillChunk[] pages, string indexTag, PublisherParseContext context) {
        PublisherQuillChunk index = Find(indexTag);
        if (index.Length != 16 + pages.Length * 8 || Data.U32(index.Offset) != pages.Length
            || Data.U32(index.Offset + 4) != 4 || Data.U32(index.Offset + 8) != 0 || Data.U32(index.Offset + 12) != 0) return false;
        int previous = Text.Offset;
        for (int p = 0; p < pages.Length; p++) {
            context.Record(); PublisherQuillChunk page = pages[p];
            if (page.Id != p || page.Length != 512 || page.Offset % 512 != 0) return false;
            int count = Data.U16(page.Offset);
            if (count is < 1 or > 84 || Data.U16(page.Offset + 2) != 1 || Data.U32(page.Offset + 4) != (p == 0 ? 0 : previous)) return false;
            var spans = new List<PublisherQuillSourceSpan>();
            for (int r = 0; r < count; r++) {
                context.Record(); int end = Data.Offset(Data.U32(page.Offset + 8 + r * 4));
                if (end <= previous || end > Text.End || ((end - Text.Offset) & 1) != 0) return false;
                previous = end;
                int offset = page.Offset + Data.U16(page.Offset + 8 + count * 4 + r * 2);
                if (offset < page.Offset + 8 + count * 6 || offset > page.End - 4) return false;
                int length = Data.Offset(Data.U32(offset));
                if (length < 4 || length > page.End - offset) return false;
                var span = new PublisherQuillSourceSpan(offset, length);
                if (spans.Any(other => other != span && offset < other.Offset + other.Length && other.Offset < offset + length)) return false;
                spans.Add(span);
            }
            if (Data.U32(index.Offset + 16 + p * 4) != previous
                || Data.U32(index.Offset + 16 + pages.Length * 4 + p * 4) != page.Offset) return false;
        }
        return previous == Text.End;
    }
}

internal readonly record struct PublisherQuillSourceDirectory(int Offset, int Count) {
    internal int PageOffset => Offset == 24 ? 0 : Offset;
}
