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

    /// <summary>Determines whether a Unicode scalar belongs to a right-to-left script range.</summary>
    public static bool IsRightToLeftScalar(int scalar) =>
        IsInRange(scalar, 0x0590, 0x05FF) ||
        IsInRange(scalar, 0x0600, 0x06FF) ||
        IsInRange(scalar, 0x0700, 0x074F) ||
        IsInRange(scalar, 0x0750, 0x077F) ||
        IsInRange(scalar, 0x0780, 0x07BF) ||
        IsInRange(scalar, 0x07C0, 0x07FF) ||
        IsInRange(scalar, 0x0840, 0x085F) ||
        IsInRange(scalar, 0x08A0, 0x08FF) ||
        IsInRange(scalar, 0xFB1D, 0xFDFF) ||
        IsInRange(scalar, 0xFE70, 0xFEFF) ||
        IsInRange(scalar, 0x1E900, 0x1E95F) ||
        IsInRange(scalar, 0x1EE00, 0x1EEFF);

    private static bool IsInRange(int value, int minimum, int maximum) => value >= minimum && value <= maximum;
}
