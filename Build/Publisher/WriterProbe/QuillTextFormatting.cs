using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Calculated formatting pages and their absolute Quill stream indexes.</summary>
internal static class QuillTextFormatting {
    internal const int PageBytes = 512;

    /// <summary>Creates one character page, one paragraph page and both required indexes.</summary>
    internal static IEnumerable<QuillRecord> CreateRecords(int absoluteTextEnd, int firstPageOffset, IReadOnlyList<QuillFormattingRange> characterRanges) {
        if (firstPageOffset % PageBytes != 0) throw new ArgumentException("The first formatting page must be aligned to 512 bytes.", nameof(firstPageOffset));
        if (characterRanges.Count == 0 || characterRanges[^1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Character formatting must cover the complete text section.", nameof(characterRanges));
        int characterPageOffset = firstPageOffset;
        int paragraphPageOffset = checked(characterPageOffset + PageBytes);
        yield return new QuillRecord("FDPC", 0, Page(characterRanges), PageBytes);
        yield return new QuillRecord("FDPP", 0, Page(new[] { new QuillFormattingRange(absoluteTextEnd, Chunk()) }), PageBytes);
        yield return new QuillRecord("BTEC", 0, Index(absoluteTextEnd, characterPageOffset));
        yield return new QuillRecord("BTEP", 0, Index(absoluteTextEnd, paragraphPageOffset));
    }

    private static byte[] Page(IReadOnlyList<QuillFormattingRange> ranges) {
        int rangeCount = ranges.Count;
        if (rangeCount is < 1 or > 84) throw new ArgumentException("The formatting page requires between 1 and 84 ranges.", nameof(ranges));
        int rangeHeaderBytes = 8 + 6 * rangeCount;
        int styleOffset = PageBytes - ranges.Sum(range => range.Style.Length);
        if (styleOffset < rangeHeaderBytes) throw new InvalidDataException("The generated style exceeds its formatting page.");
        int previousEnd = 0;
        foreach (QuillFormattingRange range in ranges) {
            if (range.AbsoluteEnd <= previousEnd) throw new ArgumentException("Formatting endpoints must increase.", nameof(ranges));
            previousEnd = range.AbsoluteEnd;
        }
        return Bytes(writer => {
            writer.Write((ushort)rangeCount);
            writer.Write((ushort)1); // The native validator requires this auxiliary value.
            writer.Write(0U);
            foreach (QuillFormattingRange range in ranges) writer.Write(range.AbsoluteEnd);
            int offset = styleOffset;
            foreach (QuillFormattingRange range in ranges) {
                writer.Write((ushort)offset); // Parallel page-relative style positions.
                offset += range.Style.Length;
            }
            writer.Write(new byte[styleOffset - rangeHeaderBytes]);
            foreach (QuillFormattingRange range in ranges) writer.Write(range.Style);
        });
    }

    private static byte[] Index(int absoluteTextEnd, int absolutePageOffset) => Bytes(writer => {
        writer.Write(1U); // Number of formatting-page references.
        writer.Write(4U); // Width of each page reference.
        writer.Write(0U); // Flags.
        writer.Write(0U);
        writer.Write(absoluteTextEnd); // Bounds use stream byte positions, unlike STRS story lengths.
        writer.Write(absolutePageOffset); // Absolute offset, rather than a page ordinal.
    });
}

internal sealed record QuillFormattingRange(int AbsoluteEnd, byte[] Style);
