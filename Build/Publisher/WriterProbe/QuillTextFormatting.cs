using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;
using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.Internal.PublisherQuillFormattingPage;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Calculated formatting pages and their absolute Quill stream indexes.</summary>
internal static class QuillTextFormatting {
    /// <summary>Creates character and paragraph pages from the same ordered story plan.</summary>
    internal static IEnumerable<QuillRecord> CreateRecords(int absoluteTextEnd, int firstPageOffset,
        IReadOnlyList<PublisherQuillFormattingRange> characterRanges, IReadOnlyList<PublisherQuillFormattingRange> paragraphRanges,
        int maximumPageCount) {
        if (firstPageOffset % PageBytes != 0) throw new ArgumentException("The first formatting page must be aligned to 512 bytes.", nameof(firstPageOffset));
        if (characterRanges.Count == 0 || characterRanges[^1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Character formatting must cover the complete text section.", nameof(characterRanges));
        if (paragraphRanges.Count == 0 || paragraphRanges[^1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Paragraph formatting must cover the complete text section.", nameof(paragraphRanges));
        IReadOnlyList<byte[]> pc = Pages(characterRanges, maximumPageCount - 1);
        IReadOnlyList<byte[]> pp = Pages(paragraphRanges, maximumPageCount - pc.Count);
        int[] pcOffsets = Enumerable.Range(0, pc.Count).Select(i => checked(firstPageOffset + i * PageBytes)).ToArray();
        int paragraphStart = checked(firstPageOffset + pc.Count * PageBytes);
        int[] ppOffsets = Enumerable.Range(0, pp.Count).Select(i => checked(paragraphStart + i * PageBytes)).ToArray();
        foreach (QuillRecord page in Records("FDPC", pc)) yield return page;
        foreach (QuillRecord page in Records("FDPP", pp)) yield return page;
        yield return new QuillRecord("BTEC", 0, Index(pc, pcOffsets));
        yield return new QuillRecord("BTEP", 0, Index(pp, ppOffsets));
    }

    private static IReadOnlyList<byte[]> Pages(IReadOnlyList<PublisherQuillFormattingRange> ranges, int maximumPages) {
        if (maximumPages < 1) throw new NotSupportedException("The formatting pages exceed the specimen's single directory.");
        // Preserve earlier unshared encodings. Table cells can exceed that
        // envelope while repeating the same few native style definitions.
        if (8 + ranges.Count * 6 + ranges.Sum(range => range.Style.Length) <= PageBytes)
            return new[] { PublisherQuillFormattingPage.Write(ranges) };
        return PublisherQuillPageAllocator.PackMinimum(ranges, maximumPages, CancellationToken.None)
            ?? throw new NotSupportedException("The formatting ranges exceed the specimen's single-directory capacity.");
    }

    private static IEnumerable<QuillRecord> Records(string tag, IReadOnlyList<byte[]> pages) {
        for (int i = 0; i < pages.Count; i++) {
            if (i != 0) SetPreviousEnd(pages[i], 0, LastEnd(pages[i - 1]));
            yield return new QuillRecord(tag, checked((ushort)i), pages[i], PageBytes);
        }
    }

    private static byte[] Index(IReadOnlyList<byte[]> pages, IReadOnlyList<int> offsets) {
        var bytes = new byte[checked(16 + pages.Count * 8)];
        WriteIndex(bytes, 0, pages, offsets);
        return bytes;
    }
}
