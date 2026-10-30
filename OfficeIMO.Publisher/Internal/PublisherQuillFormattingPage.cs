namespace OfficeIMO.Publisher.Internal;

/// <summary>One native 512-byte formatting page, shared by preserved-source editing and template-free feasibility encoding.</summary>
internal static class PublisherQuillFormattingPage {
    internal const int PageBytes = 512;

    internal static byte[] Write(IReadOnlyList<PublisherQuillFormattingRange> ranges) {
        int count = ranges.Count;
        if (count is < 1 or > 84) throw new ArgumentException("The formatting page requires between 1 and 84 ranges.", nameof(ranges));
        return Encode(ranges, shared: false) ?? throw new InvalidDataException("The generated style exceeds its formatting page.");
    }

    /// <summary>Uses the observed native sharing of byte-identical OPLs within one page; null means the page is full.</summary>
    internal static byte[]? TryWriteShared(IReadOnlyList<PublisherQuillFormattingRange> ranges) =>
        ranges.Count is < 1 or > 84 ? null : Encode(ranges, shared: true);

    private static byte[]? Encode(IReadOnlyList<PublisherQuillFormattingRange> ranges, bool shared) {
        int count = ranges.Count;
        var styles = new List<byte[]>();
        var indexes = new List<int>();
        foreach (PublisherQuillFormattingRange range in ranges) {
            int index = shared ? styles.FindIndex(style => style.SequenceEqual(range.Style)) : -1;
            if (index < 0) { index = styles.Count; styles.Add(range.Style); }
            indexes.Add(index);
        }
        int headerBytes = 8 + 6 * count;
        int styleOffset = PageBytes - styles.Sum(style => style.Length);
        if (styleOffset < headerBytes) return null;
        int previousEnd = 0;
        foreach (PublisherQuillFormattingRange range in ranges) {
            if (range.AbsoluteEnd <= previousEnd) throw new ArgumentException("Formatting endpoints must increase.", nameof(ranges));
            previousEnd = range.AbsoluteEnd;
        }
        using var stream = new MemoryStream(PageBytes);
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write((ushort)count); writer.Write((ushort)1); writer.Write(0U);
            foreach (PublisherQuillFormattingRange range in ranges) writer.Write(range.AbsoluteEnd);
            var offsets = new List<int>();
            int offset = styleOffset;
            foreach (byte[] style in styles) { offsets.Add(offset); offset += style.Length; }
            foreach (int index in indexes) writer.Write((ushort)offsets[index]);
            writer.Write(new byte[styleOffset - headerBytes]);
            foreach (byte[] style in styles) writer.Write(style);
        }
        return stream.ToArray();
    }
}

internal sealed record PublisherQuillFormattingRange(int AbsoluteEnd, byte[] Style);
