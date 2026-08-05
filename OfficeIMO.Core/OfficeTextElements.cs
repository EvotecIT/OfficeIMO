using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Unicode text-element helpers shared by OfficeIMO.Drawing text layout and raster rendering.
/// </summary>
public static class OfficeTextElements {
    /// <summary>Enumerates Unicode grapheme clusters without splitting surrogate pairs or combining sequences.</summary>
    public static IEnumerable<string> Enumerate(string? value) {
        if (string.IsNullOrEmpty(value)) {
            yield break;
        }

        TextElementEnumerator enumerator = StringInfo.GetTextElementEnumerator(value);
        while (enumerator.MoveNext()) {
            yield return enumerator.GetTextElement();
        }
    }

    /// <summary>Splits text into Unicode grapheme clusters.</summary>
    public static IReadOnlyList<string> Split(string? value, bool includeEmptyElement = false) {
        var elements = new List<string>();
        foreach (string element in Enumerate(value)) {
            elements.Add(element);
        }

        return elements.Count == 0 && includeEmptyElement
            ? new[] { string.Empty }
            : elements;
    }

    /// <summary>Removes the last Unicode grapheme cluster.</summary>
    public static string RemoveLast(string value) {
        int[] indexes = StringInfo.ParseCombiningCharacters(value);
        return indexes.Length <= 1 ? string.Empty : value.Substring(0, indexes[indexes.Length - 1]);
    }

    /// <summary>Removes the first Unicode grapheme cluster.</summary>
    public static string RemoveFirst(string value) {
        int[] indexes = StringInfo.ParseCombiningCharacters(value);
        return indexes.Length <= 1 ? string.Empty : value.Substring(indexes[1]);
    }

    /// <summary>Determines whether text contains a scalar from a right-to-left script range.</summary>
    public static bool ContainsRightToLeft(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        for (int index = 0; index < value!.Length; index++) {
            int scalar = value[index];
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                scalar = char.ConvertToUtf32(value[index], value[++index]);
            }

            if (IsRightToLeftScalar(scalar)) return true;
        }

        return false;
    }

    /// <summary>Determines whether text contains a script that requires joining or contextual shaping.</summary>
    public static bool ContainsJoiningScript(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        for (int index = 0; index < value!.Length; index++) {
            int scalar = value[index];
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                scalar = char.ConvertToUtf32(value[index], value[++index]);
            }
            if (IsJoiningScriptScalar(scalar)) return true;
        }
        return false;
    }

    /// <summary>
    /// Determines whether text contains an Indic, Southeast Asian, or related script whose glyph
    /// selection or mark positioning requires a full shaping engine.
    /// </summary>
    public static bool ContainsShapingRequiredScript(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        for (int index = 0; index < value!.Length; index++) {
            int scalar = value[index];
            if (char.IsHighSurrogate(value[index]) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1])) {
                scalar = char.ConvertToUtf32(value[index], value[++index]);
            }
            if (IsShapingRequiredScriptScalar(scalar)) return true;
        }
        return false;
    }

    /// <summary>Determines whether text contains explicit Unicode bidi embedding, override, or isolate controls.</summary>
    public static bool ContainsBidiControl(string? value) {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (char character in value!) {
            if (character == '\u061C' || character == '\u200E' || character == '\u200F'
                || character >= '\u202A' && character <= '\u202E'
                || character >= '\u2066' && character <= '\u2069') return true;
        }
        return false;
    }

    /// <summary>Resolves base direction from the first strong Unicode character.</summary>
    public static OfficeTextDirection ResolveBaseDirection(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return OfficeTextDirection.Auto;
        }

        for (int index = 0; index < value!.Length;) {
            int scalarIndex = index;
            int scalar = ReadScalar(value, ref index);
            if (scalar == 0x061C || scalar == 0x200F) {
                return OfficeTextDirection.RightToLeft;
            }

            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(value, scalarIndex);
            if (IsRightToLeftScalar(scalar) &&
                (IsLetterCategory(category) || category == UnicodeCategory.OtherNotAssigned)) {
                return OfficeTextDirection.RightToLeft;
            }

            if (scalar == 0x200E) {
                return OfficeTextDirection.LeftToRight;
            }

            if (IsStrongLeftToRightCategory(category)) {
                return OfficeTextDirection.LeftToRight;
            }
        }

        return OfficeTextDirection.Auto;
    }

    /// <summary>Determines whether a Unicode scalar belongs to a right-to-left script range.</summary>
    public static bool IsRightToLeftScalar(int scalar) =>
        IsInRange(scalar, 0x0590, 0x05FF) ||
        IsInRange(scalar, 0x0600, 0x06FF) ||
        IsInRange(scalar, 0x0700, 0x074F) ||
        IsInRange(scalar, 0x0750, 0x077F) ||
        IsInRange(scalar, 0x0780, 0x07BF) ||
        IsInRange(scalar, 0x07C0, 0x07FF) ||
        IsInRange(scalar, 0x0800, 0x083F) ||
        IsInRange(scalar, 0x0840, 0x085F) ||
        IsInRange(scalar, 0x0860, 0x089F) ||
        IsInRange(scalar, 0x08A0, 0x08FF) ||
        IsInRange(scalar, 0xFB1D, 0xFDFF) ||
        IsInRange(scalar, 0xFE70, 0xFEFF) ||
        IsInRange(scalar, 0x10800, 0x10FFF) ||
        IsInRange(scalar, 0x1E800, 0x1E8DF) ||
        IsInRange(scalar, 0x1E900, 0x1E95F) ||
        IsInRange(scalar, 0x1EE00, 0x1EEFF);

    private static bool IsJoiningScriptScalar(int scalar) =>
        IsInRange(scalar, 0x0600, 0x08FF)
        || IsInRange(scalar, 0xFB50, 0xFDFF)
        || IsInRange(scalar, 0xFE70, 0xFEFF)
        || IsInRange(scalar, 0x1EE00, 0x1EEFF);

    private static bool IsShapingRequiredScriptScalar(int scalar) =>
        IsInRange(scalar, 0x0900, 0x0DFF) // Indic scripts
        || IsInRange(scalar, 0x0E00, 0x0FFF) // Thai, Lao, Tibetan
        || IsInRange(scalar, 0x1000, 0x109F) // Myanmar
        || IsInRange(scalar, 0x1780, 0x17FF) // Khmer
        || IsInRange(scalar, 0x1800, 0x18AF) // Mongolian
        || IsInRange(scalar, 0x1900, 0x197F) // Limbu and Tai Le
        || IsInRange(scalar, 0x1980, 0x19DF) // New Tai Lue
        || IsInRange(scalar, 0x1A00, 0x1A1F) // Buginese
        || IsInRange(scalar, 0x1A20, 0x1AAF) // Tai Tham
        || IsInRange(scalar, 0x1B00, 0x1C4F) // Balinese through Lepcha
        || IsInRange(scalar, 0xA800, 0xA87F) // Syloti Nagri and Phags-pa
        || IsInRange(scalar, 0xA880, 0xA8FF) // Saurashtra and Devanagari Extended
        || IsInRange(scalar, 0xA900, 0xA95F) // Kayah Li and Rejang
        || IsInRange(scalar, 0xA980, 0xAA7F) // Javanese, Myanmar extensions, Cham
        || IsInRange(scalar, 0xAA80, 0xAADF) // Tai Viet
        || IsInRange(scalar, 0xABC0, 0xABFF) // Meetei Mayek
        || IsInRange(scalar, 0x11000, 0x11FFF); // Supplementary Brahmic scripts

    private static bool IsStrongLeftToRightCategory(UnicodeCategory category) =>
        IsLetterCategory(category);

    private static bool IsLetterCategory(UnicodeCategory category) =>
        category == UnicodeCategory.UppercaseLetter ||
        category == UnicodeCategory.LowercaseLetter ||
        category == UnicodeCategory.TitlecaseLetter ||
        category == UnicodeCategory.ModifierLetter ||
        category == UnicodeCategory.OtherLetter;

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first) &&
            index < text.Length &&
            char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }

    private static bool IsInRange(int value, int minimum, int maximum) => value >= minimum && value <= maximum;
}
