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
}
