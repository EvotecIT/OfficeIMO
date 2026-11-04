using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillCharacterStyleWriter {
    /// <summary>Language is an explicit native request; an unspecified replacement retains source inheritance.</summary>
    internal static bool HasLanguageChange(OfficeRichTextRun before, OfficeRichTextRun after) =>
        after.LanguageId.HasValue && before.LanguageId != after.LanguageId;

    private static bool LanguageFieldIsQualified(IReadOnlyList<PublisherBlock> fields) {
        PublisherBlock[] language = fields.Where(field => field.Id == 0x12).ToArray();
        return language.Length <= 1 && language.All(field => field.Type == 0x22 && field.Value <= ushort.MaxValue);
    }
}
