namespace OfficeIMO.Publisher.Internal;

/// <summary>Bounds formatting allocation and retains the established single-directory header-growth guard.</summary>
internal static class PublisherQuillPageAllocator {
    internal static byte[]? TryWrite(PublisherQuillSourceLayout layout, byte[] text,
        IReadOnlyList<byte[]> pc, IReadOnlyList<byte[]> pp, IReadOnlyList<int> storyLengths,
        long maximumBytes, CancellationToken token, IReadOnlyDictionary<string, byte[]>? resources = null) {
        int count = checked(layout.Chunks.Count - layout.CharacterPages.Length - layout.ParagraphPages.Length + pc.Count + pp.Count);
        if (pc.Count == 0 || pp.Count == 0 || count > PublisherQuillSourceLayout.MaximumDescriptors) return null;
        if (layout.Directories.Count == 1 && count <= PublisherQuillSourceLayout.DescriptorsPerDirectory && resources?.Count is not > 0)
            for (int i = 32 + layout.Chunks.Count * 24; i < 32 + count * 24; i++)
                if (layout.Data.Bytes[i] != 0) return null;
        return PublisherQuillStorageWriter.TryWrite(layout, text, pc, pp, storyLengths,
            resources ?? new Dictionary<string, byte[]>(), maximumBytes, token);
    }

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
