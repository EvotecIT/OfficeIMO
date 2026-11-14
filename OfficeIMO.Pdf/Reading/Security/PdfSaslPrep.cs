using System.Globalization;
using System.Text;

namespace OfficeIMO.Pdf;

/// <summary>
/// SASLprep (RFC 4013), the stringprep profile (RFC 3454) that ISO 32000-2 applies to revision 5 and later
/// Standard security passwords before UTF-8 encoding.
/// </summary>
/// <remarks>
/// Unassigned code points are allowed (RFC 3454 "query" semantics): the check against Unicode 3.2 would reject
/// passwords whose characters were assigned later and that other conforming readers accept. The bidirectional rule
/// uses table D.1 exactly and treats letters and spacing marks outside D.1 as left-to-right (table D.2).
/// </remarks>
internal static class PdfSaslPrep {
    internal static bool TryPrepare(string input, out string prepared) {
        prepared = string.Empty;
        var mapped = new StringBuilder(input.Length);
        for (int index = 0; index < input.Length; index++) {
            char current = input[index];
            if (IsNonAsciiSpace(current)) {
                mapped.Append(' ');
            } else if (!IsMappedToNothing(current)) {
                mapped.Append(current);
            }
        }

        string normalized;
        try {
            normalized = mapped.ToString().Normalize(NormalizationForm.FormKC);
        } catch (ArgumentException) {
            // Unpaired surrogates (prohibited table C.5) cannot be normalized.
            return false;
        }

        bool hasRandAL = false;
        bool hasL = false;
        int firstCodePoint = -1;
        int lastCodePoint = -1;
        for (int index = 0; index < normalized.Length; index++) {
            int codePoint;
            if (char.IsHighSurrogate(normalized[index])) {
                if (index + 1 >= normalized.Length || !char.IsLowSurrogate(normalized[index + 1])) return false;
                codePoint = char.ConvertToUtf32(normalized[index], normalized[index + 1]);
                index++;
            } else if (char.IsLowSurrogate(normalized[index])) {
                return false;
            } else {
                codePoint = normalized[index];
            }

            if (IsProhibited(codePoint)) return false;
            if (IsRandAL(codePoint)) {
                hasRandAL = true;
            } else if (IsLeftToRight(codePoint, normalized, index)) {
                hasL = true;
            }

            if (firstCodePoint < 0) firstCodePoint = codePoint;
            lastCodePoint = codePoint;
        }

        // RFC 3454 section 6: a string with right-to-left characters has no left-to-right characters and
        // starts and ends with a right-to-left character.
        if (hasRandAL && (hasL || !IsRandAL(firstCodePoint) || !IsRandAL(lastCodePoint))) return false;
        prepared = normalized;
        return true;
    }

    // RFC 3454 table B.1.
    private static bool IsMappedToNothing(char value) =>
        value == '\u00AD' || value == '\u034F' || value == '\u1806' ||
        (value >= '\u180B' && value <= '\u180D') ||
        (value >= '\u200B' && value <= '\u200D') ||
        value == '\u2060' ||
        (value >= '\uFE00' && value <= '\uFE0F') ||
        value == '\uFEFF';

    // RFC 3454 table C.1.2. RFC 4013 maps these to SPACE before B.1 removes the remaining zero-width characters.
    private static bool IsNonAsciiSpace(char value) =>
        value == '\u00A0' || value == '\u1680' ||
        (value >= '\u2000' && value <= '\u200B') ||
        value == '\u202F' || value == '\u205F' || value == '\u3000';

    // RFC 4013 section 2.3: tables C.1.2, C.2.1, C.2.2, C.3, C.4, C.5, C.6, C.7, C.8 and C.9.
    private static bool IsProhibited(int c) {
        if (c < 0x20 || c == 0x7F || (c >= 0x80 && c <= 0x9F)) return true;
        if (c == 0xA0 || c == 0x1680 || (c >= 0x2000 && c <= 0x200F) || c == 0x202F || c == 0x205F || c == 0x3000) return true;
        if (c == 0x0340 || c == 0x0341 || c == 0x06DD || c == 0x070F || c == 0x180E) return true;
        if (c == 0x2028 || c == 0x2029 || (c >= 0x202A && c <= 0x202E) || (c >= 0x2060 && c <= 0x2063) || (c >= 0x206A && c <= 0x206F)) return true;
        if (c >= 0x2FF0 && c <= 0x2FFB) return true;
        if (c >= 0xD800 && c <= 0xDFFF) return true;
        if (c >= 0xE000 && c <= 0xF8FF) return true;
        if (c >= 0xFDD0 && c <= 0xFDEF) return true;
        if (c == 0xFEFF || (c >= 0xFFF9 && c <= 0xFFFD)) return true;
        if ((c & 0xFFFE) == 0xFFFE) return true;
        if (c >= 0x1D173 && c <= 0x1D17A) return true;
        if (c == 0xE0001 || (c >= 0xE0020 && c <= 0xE007F)) return true;
        return (c >= 0xF0000 && c <= 0xFFFFD) || (c >= 0x100000 && c <= 0x10FFFD);
    }

    // RFC 3454 table D.1.
    private static readonly int[] RandALRanges = {
        0x05BE, 0x05BE, 0x05C0, 0x05C0, 0x05C3, 0x05C3, 0x05D0, 0x05EA, 0x05F0, 0x05F4,
        0x061B, 0x061B, 0x061F, 0x061F, 0x0621, 0x063A, 0x0640, 0x064A, 0x066D, 0x066F,
        0x0671, 0x06D5, 0x06DD, 0x06DD, 0x06E5, 0x06E6, 0x06FA, 0x06FE, 0x0700, 0x070D,
        0x0710, 0x0710, 0x0712, 0x072C, 0x0780, 0x07A5, 0x07B1, 0x07B1, 0x200F, 0x200F,
        0xFB1D, 0xFB1D, 0xFB1F, 0xFB28, 0xFB2A, 0xFB36, 0xFB38, 0xFB3C, 0xFB3E, 0xFB3E,
        0xFB40, 0xFB41, 0xFB43, 0xFB44, 0xFB46, 0xFBB1, 0xFBD3, 0xFD3D, 0xFD50, 0xFD8F,
        0xFD92, 0xFDC7, 0xFDF0, 0xFDFC, 0xFE70, 0xFE74, 0xFE76, 0xFEFC
    };

    private static bool IsRandAL(int codePoint) {
        for (int index = 0; index < RandALRanges.Length; index += 2) {
            if (codePoint < RandALRanges[index]) return false;
            if (codePoint <= RandALRanges[index + 1]) return true;
        }

        return false;
    }

    private static bool IsLeftToRight(int codePoint, string text, int index) {
        // Letters assigned after Unicode 3.2 in right-to-left blocks are not left-to-right.
        if ((codePoint >= 0x0590 && codePoint <= 0x08FF) || (codePoint >= 0xFB1D && codePoint <= 0xFDFF) ||
            (codePoint >= 0xFE70 && codePoint <= 0xFEFF) || (codePoint >= 0x10800 && codePoint <= 0x10FFF) ||
            (codePoint >= 0x1E800 && codePoint <= 0x1EFFF)) return false;
        UnicodeCategory category = codePoint > 0xFFFF
            ? CharUnicodeInfo.GetUnicodeCategory(text, index - 1)
            : CharUnicodeInfo.GetUnicodeCategory((char)codePoint);
        return category is UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or
            UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter or
            UnicodeCategory.SpacingCombiningMark;
    }
}
