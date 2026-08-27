using System;
using System.Collections.Generic;
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
    ToggleCase,

    /// <summary>Uppercases the first cased character of each word while preserving the remaining characters.</summary>
    Capitalize
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
            case OfficeTextCase.Capitalize:
                return CapitalizeWords(text, selectedCulture);
            default:
                throw new ArgumentOutOfRangeException(nameof(textCase), textCase, "Unsupported text casing transformation.");
        }
    }

    /// <summary>
    /// Applies one casing transformation across adjacent text segments, then redistributes the result without
    /// resetting sentence or word context at formatting boundaries.
    /// </summary>
    public static IReadOnlyList<string> ApplySegments(
        IReadOnlyList<string> segments,
        OfficeTextCase textCase,
        CultureInfo? culture = null) {
        if (segments == null) throw new ArgumentNullException(nameof(segments));
        if (segments.Count == 0) return Array.Empty<string>();

        var source = new StringBuilder();
        foreach (string segment in segments) source.Append(segment ?? string.Empty);
        string transformed = Apply(source.ToString(), textCase, culture);
        var result = new string[segments.Count];
        int sourceOffset = 0;
        int transformedOffset = 0;
        for (int index = 0; index < segments.Count; index++) {
            sourceOffset += (segments[index] ?? string.Empty).Length;
            int transformedEnd = index == segments.Count - 1
                ? transformed.Length
                : Apply(source.ToString(0, sourceOffset), textCase, culture).Length;
            transformedEnd = Math.Max(transformedOffset, Math.Min(transformed.Length, transformedEnd));
            result[index] = transformed.Substring(transformedOffset, transformedEnd - transformedOffset);
            transformedOffset = transformedEnd;
        }
        return result;
    }

    private static string ToSentenceCase(string text, CultureInfo culture) {
        if (text.Length == 0) return text;
        string normalized = text.ToLower(culture);
        var result = new StringBuilder(normalized.Length);
        bool capitalizeNext = true;
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(normalized);
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            if (capitalizeNext && IsLetter(element)) {
                result.Append(element.ToUpper(culture));
                capitalizeNext = false;
            } else {
                result.Append(element);
            }

            if (element.IndexOf('.') >= 0 || element.IndexOf('!') >= 0 || element.IndexOf('?') >= 0 ||
                element.IndexOf('\r') >= 0 || element.IndexOf('\n') >= 0) {
                capitalizeNext = true;
            }
        }
        return result.ToString();
    }

    private static string Toggle(string text, CultureInfo culture) {
        if (text.Length == 0) return text;
        var result = new StringBuilder(text.Length);
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(text);
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            string upper = element.ToUpper(culture);
            string lower = element.ToLower(culture);
            if (string.Equals(element, upper, StringComparison.Ordinal) && !string.Equals(element, lower, StringComparison.Ordinal)) {
                result.Append(lower);
            } else if (string.Equals(element, lower, StringComparison.Ordinal) && !string.Equals(element, upper, StringComparison.Ordinal)) {
                result.Append(upper);
            } else {
                result.Append(element);
            }
        }
        return result.ToString();
    }

    private static string CapitalizeWords(string text, CultureInfo culture) {
        if (text.Length == 0) return text;
        var result = new StringBuilder(text.Length);
        bool capitalizeNext = true;
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(text);
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            if (capitalizeNext && IsLetter(element)) {
                result.Append(element.ToUpper(culture));
                capitalizeNext = false;
            } else {
                result.Append(element);
                if (IsLetter(element) || IsCombiningMark(element)) capitalizeNext = false;
            }

            if (IsWordSeparator(element)) capitalizeNext = true;
        }
        return result.ToString();
    }

    private static bool IsWordSeparator(string textElement) {
        if (textElement == "'" || textElement == "’") return false;
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.SpaceSeparator ||
               category == UnicodeCategory.LineSeparator ||
               category == UnicodeCategory.ParagraphSeparator ||
               category == UnicodeCategory.Control ||
               category == UnicodeCategory.DashPunctuation ||
               category == UnicodeCategory.OpenPunctuation ||
               category == UnicodeCategory.ClosePunctuation ||
               category == UnicodeCategory.InitialQuotePunctuation ||
               category == UnicodeCategory.FinalQuotePunctuation ||
               category == UnicodeCategory.OtherPunctuation;
    }

    private static bool IsCombiningMark(string textElement) {
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.NonSpacingMark ||
               category == UnicodeCategory.SpacingCombiningMark ||
               category == UnicodeCategory.EnclosingMark;
    }

    private static bool IsLetter(string textElement) {
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.UppercaseLetter ||
               category == UnicodeCategory.LowercaseLetter ||
               category == UnicodeCategory.TitlecaseLetter ||
               category == UnicodeCategory.ModifierLetter ||
               category == UnicodeCategory.OtherLetter;
    }
}
