using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Defines culture-aware transformations that change stored text casing.
/// </summary>
public enum OfficeTextCase {
    /// <summary>Preserves the text exactly as supplied.</summary>
    None,

    /// <summary>Converts all cased characters to uppercase.</summary>
    Uppercase,

    /// <summary>Converts all cased characters to lowercase.</summary>
    Lowercase,

    /// <summary>Capitalizes words using the selected culture.</summary>
    TitleCase,

    /// <summary>Lowercases text and capitalizes the first cased character of each sentence.</summary>
    SentenceCase,

    /// <summary>Converts lowercase characters to uppercase and uppercase characters to lowercase.</summary>
    ToggleCase
}

/// <summary>
/// Applies reusable, culture-aware text casing without depending on a document format.
/// </summary>
public static class OfficeTextCaseTransformer {
    /// <summary>
    /// Applies the requested casing transformation.
    /// </summary>
    /// <param name="text">Text to transform.</param>
    /// <param name="textCase">Transformation to apply.</param>
    /// <param name="culture">Culture used for casing. The current culture is used when omitted.</param>
    /// <returns>The transformed text.</returns>
    public static string Apply(string text, OfficeTextCase textCase, CultureInfo? culture = null) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        CultureInfo selectedCulture = culture ?? CultureInfo.CurrentCulture;
        switch (textCase) {
            case OfficeTextCase.None:
                return text;
            case OfficeTextCase.Uppercase:
                return text.ToUpper(selectedCulture);
            case OfficeTextCase.Lowercase:
                return text.ToLower(selectedCulture);
            case OfficeTextCase.TitleCase:
                return selectedCulture.TextInfo.ToTitleCase(text.ToLower(selectedCulture));
            case OfficeTextCase.SentenceCase:
                return ToSentenceCase(text, selectedCulture);
            case OfficeTextCase.ToggleCase:
                return Toggle(text, selectedCulture);
            default:
                throw new ArgumentOutOfRangeException(nameof(textCase), textCase, "Unsupported text casing transformation.");
        }
    }

    private static string ToSentenceCase(string text, CultureInfo culture) {
        if (text.Length == 0) return text;
        string normalized = text.ToLower(culture);
        var result = new StringBuilder(normalized.Length);
        bool capitalizeNext = true;
        for (int index = 0; index < normalized.Length; index++) {
            char character = normalized[index];
            if (capitalizeNext && char.IsLetter(character)) {
                result.Append(char.ToUpper(character, culture));
                capitalizeNext = false;
            } else {
                result.Append(character);
            }

            if (character == '.' || character == '!' || character == '?' || character == '\r' || character == '\n') {
                capitalizeNext = true;
            }
        }
        return result.ToString();
    }

    private static string Toggle(string text, CultureInfo culture) {
        if (text.Length == 0) return text;
        var result = new StringBuilder(text.Length);
        for (int index = 0; index < text.Length; index++) {
            char character = text[index];
            if (char.IsUpper(character)) result.Append(char.ToLower(character, culture));
            else if (char.IsLower(character)) result.Append(char.ToUpper(character, culture));
            else result.Append(character);
        }
        return result.ToString();
    }
}
