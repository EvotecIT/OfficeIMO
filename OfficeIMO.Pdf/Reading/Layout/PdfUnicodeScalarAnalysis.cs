using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfUnicodeScalarAnalysis {
    internal static int CountScalars(string value) {
        int count = 0;
        for (int index = 0; index < value.Length; index += char.IsSurrogatePair(value, index) ? 2 : 1) count++;
        return count;
    }

    internal static int CountDecimalDigits(string value) {
        int count = 0;
        for (int index = 0; index < value.Length;) {
            if (CharUnicodeInfo.GetDecimalDigitValue(value, index) >= 0) count++;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return count;
    }

    internal static bool ContainsDecimalDigit(string value) => CountDecimalDigits(value) > 0;

    internal static bool IsFirstLetterOrDigit(string value) =>
        value.Length > 0 && IsLetterOrDigit(CharUnicodeInfo.GetUnicodeCategory(value, 0));

    internal static bool IsLastLetterOrDigit(string value) {
        if (value.Length == 0) return false;
        int index = value.Length - 1;
        if (index > 0 && char.IsLowSurrogate(value[index]) && char.IsHighSurrogate(value[index - 1])) index--;
        return IsLetterOrDigit(CharUnicodeInfo.GetUnicodeCategory(value, index));
    }

    internal static bool ContainsLetter(string value) {
        for (int index = 0; index < value.Length;) {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            if (category is UnicodeCategory.UppercaseLetter or
                UnicodeCategory.LowercaseLetter or
                UnicodeCategory.TitlecaseLetter or
                UnicodeCategory.ModifierLetter or
                UnicodeCategory.OtherLetter) return true;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return false;
    }

    internal static bool IsAllDecimalDigits(string value) {
        if (value.Length == 0) return false;
        for (int index = 0; index < value.Length;) {
            if (CharUnicodeInfo.GetDecimalDigitValue(value, index) < 0) return false;
            index += char.IsSurrogatePair(value, index) ? 2 : 1;
        }
        return true;
    }

    internal static bool IsAllWordish(string value) {
        if (value.Length == 0) return false;
        for (int index = 0; index < value.Length;) {
            int scalar = char.ConvertToUtf32(value, index);
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            if (!IsLetter(category) &&
                category is not (UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark) &&
                scalar is not (0x27 or 0x2D or 0x2F or 0x2019)) return false;
            index += scalar > 0xFFFF ? 2 : 1;
        }
        return true;
    }

    private static bool IsLetterOrDigit(UnicodeCategory category) =>
        IsLetter(category) || category == UnicodeCategory.DecimalDigitNumber;

    private static bool IsLetter(UnicodeCategory category) => category is
        UnicodeCategory.UppercaseLetter or
        UnicodeCategory.LowercaseLetter or
        UnicodeCategory.TitlecaseLetter or
        UnicodeCategory.ModifierLetter or
        UnicodeCategory.OtherLetter;
}
