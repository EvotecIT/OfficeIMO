using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillCharacterStyleWriter {
    /// <summary>Language is an explicit native request; an unspecified replacement retains source inheritance.</summary>
    internal static bool HasLanguageChange(OfficeRichTextRun before, OfficeRichTextRun after) =>
        after.LanguageId.HasValue && before.LanguageId != after.LanguageId;

    /// <summary>Applies an explicit language to an already selected native typing definition without changing its opaque fields.</summary>
    internal static byte[]? RewriteLanguage(byte[] style, uint language, PublisherParseContext context) {
        if (language > ushort.MaxValue) return null;
        var source = new PublisherBinaryData(style, "Quill/FDPC");
        var fields = new PublisherBlockReader(source, context).Chunk(0, style.Length);
        if (!LanguageFieldIsQualified(fields)) return null;
        return PublisherOplWriter.Rewrite(source, 0, style.Length, fields,
            new SortedDictionary<byte, byte[]> { [0x12] = PublisherOplWriter.UInt32(0x12, language) }, context);
    }

    private static bool LanguageFieldIsQualified(IReadOnlyList<PublisherBlock> fields) {
        PublisherBlock[] language = fields.Where(field => field.Id == 0x12).ToArray();
        return language.Length <= 1 && language.All(field => field.Type == 0x22 && field.Value <= ushort.MaxValue);
    }
}
