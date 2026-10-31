using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;
using OfficeIMO.Publisher.Internal;
using static OfficeIMO.Publisher.Internal.PublisherQuillFormattingPage;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Calculated formatting pages and their absolute Quill stream indexes.</summary>
internal static class QuillTextFormatting {
    /// <summary>Creates character and paragraph pages from the same ordered story plan.</summary>
    internal static IEnumerable<QuillRecord> CreateRecords(int absoluteTextEnd, int firstPageOffset,
        IReadOnlyList<PublisherQuillFormattingRange> characterRanges, IReadOnlyList<PublisherQuillFormattingRange> paragraphRanges) {
        if (firstPageOffset % PageBytes != 0) throw new ArgumentException("The first formatting page must be aligned to 512 bytes.", nameof(firstPageOffset));
        if (characterRanges.Count == 0 || characterRanges[^1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Character formatting must cover the complete text section.", nameof(characterRanges));
        if (paragraphRanges.Count == 0 || paragraphRanges[^1].AbsoluteEnd != absoluteTextEnd)
            throw new ArgumentException("Paragraph formatting must cover the complete text section.", nameof(paragraphRanges));
        int characterPageOffset = firstPageOffset;
        int paragraphPageOffset = checked(characterPageOffset + PageBytes);
        yield return new QuillRecord("FDPC", 0, Page(characterRanges), PageBytes);
        yield return new QuillRecord("FDPP", 0, Page(paragraphRanges), PageBytes);
        yield return new QuillRecord("BTEC", 0, Index(absoluteTextEnd, characterPageOffset));
        yield return new QuillRecord("BTEP", 0, Index(absoluteTextEnd, paragraphPageOffset));
    }

    private static byte[] Page(IReadOnlyList<PublisherQuillFormattingRange> ranges) {
        // Preserve earlier unshared encodings. Table cells can exceed that
        // envelope while repeating the same few native style definitions.
        if (8 + ranges.Count * 6 + ranges.Sum(range => range.Style.Length) <= PageBytes)
            return PublisherQuillFormattingPage.Write(ranges);
        return PublisherQuillFormattingPage.TryWriteShared(ranges)
            ?? throw new NotSupportedException("The specimen exceeds one native formatting page even with shared styles.");
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
