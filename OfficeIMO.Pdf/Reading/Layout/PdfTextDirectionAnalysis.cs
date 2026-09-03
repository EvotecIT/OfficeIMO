using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfTextDirectionAnalysis {
    internal static PdfReadingDirection Resolve(
        PdfReadingDirection requested,
        IEnumerable<string> textInSourceOrder) {
        if (requested != PdfReadingDirection.Auto) return requested;
        foreach (string text in textInSourceOrder) {
            PdfReadingDirection? direction = FindFirstStrongDirection(text);
            if (direction.HasValue) return direction.Value;
        }
        return PdfReadingDirection.LeftToRight;
    }

    private static PdfReadingDirection? FindFirstStrongDirection(string text) {
        for (int index = 0; index < text.Length;) {
            int scalar = char.ConvertToUtf32(text, index);
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(text, index);
            if (IsExplicitRightToLeftScalar(scalar) ||
                (IsRightToLeftScalar(scalar) && IsLetter(category))) {
                return PdfReadingDirection.RightToLeft;
            }
            if (IsExplicitLeftToRightScalar(scalar) || IsLetter(category)) {
                return PdfReadingDirection.LeftToRight;
            }
            index += scalar > 0xFFFF ? 2 : 1;
        }
        return null;
    }

    private static bool IsExplicitLeftToRightScalar(int scalar) =>
        scalar == 0x200E;

    private static bool IsExplicitRightToLeftScalar(int scalar) =>
        scalar is 0x200F or 0x061C;

    private static bool IsLetter(UnicodeCategory category) => category is
        UnicodeCategory.UppercaseLetter or
        UnicodeCategory.LowercaseLetter or
        UnicodeCategory.TitlecaseLetter or
        UnicodeCategory.ModifierLetter or
        UnicodeCategory.OtherLetter;

    private static bool IsRightToLeftScalar(int scalar) =>
        scalar is >= 0x0590 and <= 0x08FF ||
        scalar is >= 0xFB1D and <= 0xFDFF ||
        scalar is >= 0xFE70 and <= 0xFEFF ||
        scalar is >= 0x10800 and <= 0x10FFF ||
        scalar is >= 0x1E800 and <= 0x1EEFF;
}
