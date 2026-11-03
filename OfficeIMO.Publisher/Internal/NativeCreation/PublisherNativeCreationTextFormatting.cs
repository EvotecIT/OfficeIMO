using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.Internal.PublisherQuillFormattingPage;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Calculated formatting pages and their absolute Quill stream indexes.</summary>
internal static class PublisherNativeCreationTextFormatting {
    /// <summary>Creates character and paragraph pages from the same ordered story plan.</summary>
    internal static IEnumerable<PublisherNativeCreationQuillRecord> CreateRecords(int absoluteTextEnd, int firstPageOffset,
        IReadOnlyList<PublisherQuillFormattingRange> characterRanges, IReadOnlyList<PublisherQuillFormattingRange> paragraphRanges,
        int maximumPageCount, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        if (firstPageOffset % PageBytes != 0) throw new ArgumentException("The first formatting page must be aligned to 512 bytes.", nameof(firstPageOffset));
        if (characterRanges.Count == 0 || characterRanges[characterRanges.Count - 1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Character formatting must cover the complete text section.", nameof(characterRanges));
        if (paragraphRanges.Count == 0 || paragraphRanges[paragraphRanges.Count - 1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Paragraph formatting must cover the complete text section.", nameof(paragraphRanges));
        IReadOnlyList<byte[]> pc = Pages(characterRanges, maximumPageCount - 1, context);
        IReadOnlyList<byte[]> pp = Pages(paragraphRanges, maximumPageCount - pc.Count, context);
        int[] pcOffsets = Offsets(firstPageOffset, pc.Count, context);
        int paragraphStart = checked(firstPageOffset + pc.Count * PageBytes);
        int[] ppOffsets = Offsets(paragraphStart, pp.Count, context);
        foreach (PublisherNativeCreationQuillRecord page in Records("FDPC", pc, context)) yield return page;
        foreach (PublisherNativeCreationQuillRecord page in Records("FDPP", pp, context)) yield return page;
        yield return new PublisherNativeCreationQuillRecord("BTEC", 0, Index(pc, pcOffsets, context));
        yield return new PublisherNativeCreationQuillRecord("BTEP", 0, Index(pp, ppOffsets, context));
        context.Token.ThrowIfCancellationRequested();
    }

    private static IReadOnlyList<byte[]> Pages(IReadOnlyList<PublisherQuillFormattingRange> ranges, int maximumPages,
        PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        if (maximumPages < 1) throw new NotSupportedException("The formatting pages exceed the Quill descriptor limit.");
        // Preserve earlier unshared encodings. Table cells can exceed that
        // envelope while repeating the same few native style definitions.
        int styleBytes = 0;
        foreach (PublisherQuillFormattingRange range in ranges) {
            context.Token.ThrowIfCancellationRequested();
            styleBytes = checked(styleBytes + range.Style.Length);
        }
        if (8 + ranges.Count * 6 + styleBytes <= PageBytes) {
            byte[] page = PublisherQuillFormattingPage.Write(ranges);
            context.Token.ThrowIfCancellationRequested();
            return new[] { page };
        }
        IReadOnlyList<byte[]> pages = PublisherQuillPageAllocator.PackMinimum(ranges, maximumPages, context.Token)
            ?? throw new NotSupportedException("The formatting ranges exceed the Quill descriptor limit.");
        context.Token.ThrowIfCancellationRequested();
        return pages;
    }

    private static int[] Offsets(int start, int count, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        var offsets = new int[count];
        for (int index = 0; index < count; index++) {
            context.Token.ThrowIfCancellationRequested();
            offsets[index] = checked(start + index * PageBytes);
        }
        context.Token.ThrowIfCancellationRequested();
        return offsets;
    }

    private static IEnumerable<PublisherNativeCreationQuillRecord> Records(string tag, IReadOnlyList<byte[]> pages,
        PublisherParseContext context) {
        for (int i = 0; i < pages.Count; i++) {
            context.Token.ThrowIfCancellationRequested();
            if (i != 0) SetPreviousEnd(pages[i], 0, LastEnd(pages[i - 1]));
            yield return new PublisherNativeCreationQuillRecord(tag, checked((ushort)i), pages[i], PageBytes);
        }
        context.Token.ThrowIfCancellationRequested();
    }

    private static byte[] Index(IReadOnlyList<byte[]> pages, IReadOnlyList<int> offsets, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        var bytes = new byte[checked(16 + pages.Count * 8)];
        WriteIndex(bytes, 0, pages, offsets);
        context.Token.ThrowIfCancellationRequested();
        return bytes;
    }
}
