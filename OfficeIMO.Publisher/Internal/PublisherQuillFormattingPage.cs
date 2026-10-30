namespace OfficeIMO.Publisher.Internal;

/// <summary>One native 512-byte formatting page, shared by preserved-source editing and template-free feasibility encoding.</summary>
internal static class PublisherQuillFormattingPage {
    internal const int PageBytes = 512;

    internal static byte[] Write(IReadOnlyList<PublisherQuillFormattingRange> ranges) {
        int count = ranges.Count;
        if (count is < 1 or > 84) throw new ArgumentException("The formatting page requires between 1 and 84 ranges.", nameof(ranges));
        int headerBytes = 8 + 6 * count;
        int styleOffset = PageBytes - ranges.Sum(range => range.Style.Length);
        if (styleOffset < headerBytes) throw new InvalidDataException("The generated style exceeds its formatting page.");
        int previousEnd = 0;
        foreach (PublisherQuillFormattingRange range in ranges) {
            if (range.AbsoluteEnd <= previousEnd) throw new ArgumentException("Formatting endpoints must increase.", nameof(ranges));
            previousEnd = range.AbsoluteEnd;
        }
        using var stream = new MemoryStream(PageBytes);
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write((ushort)count); writer.Write((ushort)1); writer.Write(0U);
            foreach (PublisherQuillFormattingRange range in ranges) writer.Write(range.AbsoluteEnd);
            int offset = styleOffset;
            foreach (PublisherQuillFormattingRange range in ranges) {
                writer.Write((ushort)offset); offset += range.Style.Length;
            }
            writer.Write(new byte[styleOffset - headerBytes]);
            foreach (PublisherQuillFormattingRange range in ranges) writer.Write(range.Style);
        }
        return stream.ToArray();
    }
}

internal sealed record PublisherQuillFormattingRange(int AbsoluteEnd, byte[] Style);
