using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Calculated formatting pages and their absolute Quill stream indexes.</summary>
internal static class QuillTextFormatting {
    internal const int PageBytes = 512;

    /// <summary>Creates one character page, one paragraph page and both required indexes.</summary>
    internal static IEnumerable<QuillRecord> CreateRecords(int absoluteTextEnd, int firstPageOffset) {
        if (firstPageOffset % PageBytes != 0) throw new ArgumentException("The first formatting page must be aligned to 512 bytes.", nameof(firstPageOffset));
        int characterPageOffset = firstPageOffset;
        int paragraphPageOffset = checked(characterPageOffset + PageBytes);
        yield return new QuillRecord("FDPC", 0, Page(absoluteTextEnd, Chunk(Scalar(0x0C, 0x22, 18 * 12700))), PageBytes);
        yield return new QuillRecord("FDPP", 0, Page(absoluteTextEnd, Chunk()), PageBytes);
        yield return new QuillRecord("BTEC", 0, Index(absoluteTextEnd, characterPageOffset));
        yield return new QuillRecord("BTEP", 0, Index(absoluteTextEnd, paragraphPageOffset));
    }

    private static byte[] Page(int absoluteTextEnd, byte[] style) {
        const int rangeCount = 1;
        const int rangeHeaderBytes = 8 + 6 * rangeCount;
        int styleOffset = PageBytes - style.Length;
        if (styleOffset < rangeHeaderBytes) throw new InvalidDataException("The generated style exceeds its formatting page.");
        return Bytes(writer => {
            writer.Write((ushort)rangeCount);
            writer.Write((ushort)1); // The native validator requires this auxiliary value.
            writer.Write(0U);
            writer.Write(absoluteTextEnd); // Absolute byte position in the Quill stream.
            writer.Write((ushort)styleOffset); // Style position relative to this page.
            writer.Write(new byte[styleOffset - rangeHeaderBytes]);
            writer.Write(style);
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
