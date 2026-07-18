using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>Dependency-free Unicode line-break opportunities shared by OfficeIMO renderers.</summary>
/// <remarks>
/// The implementation deliberately covers deterministic, high-value boundaries rather than
/// claiming the complete Unicode Line Breaking Algorithm. It never splits a surrogate pair or
/// grapheme cluster and applies common CJK punctuation constraints.
/// </remarks>
public static class OfficeTextLineBreaks {
    /// <summary>
    /// Returns safe UTF-16 indexes where an unspaced token can wrap without inserting text.
    /// </summary>
    public static IReadOnlyList<int> GetBreakPositions(string? text) {
        if (text == null || text.Length == 0) {
            return Array.Empty<int>();
        }

        int[] elementStarts = StringInfo.ParseCombiningCharacters(text);
        if (elementStarts.Length < 2) {
            return Array.Empty<int>();
        }

        var positions = new List<int>();
        for (int elementIndex = 1; elementIndex < elementStarts.Length; elementIndex++) {
            int boundary = elementStarts[elementIndex];
            int left = ReadLastScalar(text, boundary);
            int right = ReadFirstScalar(text, boundary);
            if (CanBreakBetween(left, right)) {
                positions.Add(boundary);
            }
        }

        return positions.ToArray();
    }

    /// <summary>Returns true when an index is a non-edge grapheme-cluster boundary.</summary>
    public static bool IsValidBreakPosition(string? text, int position) {
        if (string.IsNullOrEmpty(text) || position <= 0 || position >= text!.Length) {
            return false;
        }

        int[] elementStarts = StringInfo.ParseCombiningCharacters(text);
        return Array.BinarySearch(elementStarts, position) >= 0;
    }

    private static bool CanBreakBetween(int left, int right) {
        if (IsNonStarter(right) || IsOpeningPunctuation(left) || IsClosingPunctuation(right)) {
            return false;
        }

        return IsCjkScalar(left) ||
            IsCjkScalar(right) ||
            IsBreakAfterScalar(left);
    }

    private static bool IsBreakAfterScalar(int scalar) => scalar is
        0x002D or // hyphen-minus
        0x002F or // solidus
        0x058A or // Armenian hyphen
        0x05BE or // Hebrew punctuation maqaf
        0x1400 or // Canadian syllabics hyphen
        0x1806 or // Mongolian todo soft hyphen
        0x200B or // zero-width space
        0x2010 or // hyphen
        0x2012 or // figure dash
        0x2013 or // en dash
        0x2027 or // hyphenation point
        0x30A0;   // katakana-hiragana double hyphen

    private static bool IsNonStarter(int scalar) {
        if (scalar == 0x200D ||
            scalar >= 0xFE00 && scalar <= 0xFE0F ||
            scalar >= 0xE0100 && scalar <= 0xE01EF) {
            return true;
        }

        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(char.ConvertFromUtf32(scalar), 0);
        return category == UnicodeCategory.NonSpacingMark ||
            category == UnicodeCategory.SpacingCombiningMark ||
            category == UnicodeCategory.EnclosingMark;
    }

    private static bool IsCjkScalar(int scalar) =>
        scalar >= 0x3040 && scalar <= 0x309F ||
        scalar >= 0x30A0 && scalar <= 0x30FF ||
        scalar >= 0x31F0 && scalar <= 0x31FF ||
        scalar >= 0x3400 && scalar <= 0x4DBF ||
        scalar >= 0x4E00 && scalar <= 0x9FFF ||
        scalar >= 0xAC00 && scalar <= 0xD7AF ||
        scalar >= 0xF900 && scalar <= 0xFAFF ||
        scalar >= 0x20000 && scalar <= 0x2FA1F;

    private static bool IsOpeningPunctuation(int scalar) => scalar switch {
        '(' or '[' or '{' or 0x201C or 0x2018 or 0x3008 or 0x300A or 0x300C or
        0x300E or 0x3010 or 0x3014 or 0x3016 or 0x3018 or 0x301A or 0xFF08 or
        0xFF3B or 0xFF5B => true,
        _ => false
    };

    private static bool IsClosingPunctuation(int scalar) => scalar switch {
        ')' or ']' or '}' or ',' or '.' or ':' or ';' or '!' or '?' or 0x2019 or
        0x201D or 0x2026 or 0x3001 or 0x3002 or 0x3009 or 0x300B or 0x300D or
        0x300F or 0x3011 or 0x3015 or 0x3017 or 0x3019 or 0x301B or 0xFF01 or
        0xFF09 or 0xFF0C or 0xFF0E or 0xFF1A or 0xFF1B or 0xFF1F or 0xFF3D or
        0xFF5D => true,
        _ => false
    };

    private static int ReadFirstScalar(string text, int index) {
        char first = text[index];
        return char.IsHighSurrogate(first) &&
            index + 1 < text.Length &&
            char.IsLowSurrogate(text[index + 1])
            ? char.ConvertToUtf32(first, text[index + 1])
            : first;
    }

    private static int ReadLastScalar(string text, int boundary) {
        char last = text[boundary - 1];
        return char.IsLowSurrogate(last) &&
            boundary > 1 &&
            char.IsHighSurrogate(text[boundary - 2])
            ? char.ConvertToUtf32(text[boundary - 2], last)
            : last;
    }
}
