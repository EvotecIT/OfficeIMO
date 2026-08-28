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
        return TransformSegments(new[] { text }, textCase, culture ?? CultureInfo.CurrentCulture)[0];
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
        return TransformSegments(segments, textCase, culture ?? CultureInfo.CurrentCulture);
    }

    private static string[] TransformSegments(IReadOnlyList<string> segments, OfficeTextCase textCase, CultureInfo culture) {
        if (!Enum.IsDefined(typeof(OfficeTextCase), textCase)) {
            throw new ArgumentOutOfRangeException(nameof(textCase), textCase, "Unsupported text casing transformation.");
        }

        var source = new StringBuilder();
        var boundaries = new int[segments.Count];
        var result = new StringBuilder[segments.Count];
        for (int index = 0; index < segments.Count; index++) {
            string value = segments[index] ?? string.Empty;
            source.Append(value);
            boundaries[index] = source.Length;
            result[index] = new StringBuilder(value.Length);
        }

        if (textCase == OfficeTextCase.None) {
            for (int index = 0; index < segments.Count; index++) result[index].Append(segments[index] ?? string.Empty);
            return ToStrings(result);
        }

        bool sentenceStart = true;
        bool capitalizeWordStart = true;
        bool titleWordStart = true;
        bool titleDutchJ = false;
        bool titlecasesDutchIJ = string.Equals(culture.TwoLetterISOLanguageName, "nl", StringComparison.OrdinalIgnoreCase) &&
            string.Equals(culture.TextInfo.ToTitleCase("ij"), "IJ", StringComparison.Ordinal);
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(source.ToString());
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            string lower = element.ToLower(culture);
            string transformed;
            switch (textCase) {
                case OfficeTextCase.Uppercase:
                    transformed = element.ToUpper(culture);
                    break;
                case OfficeTextCase.Lowercase:
                    transformed = lower;
                    break;
                case OfficeTextCase.TitleCase:
                    transformed = TransformTitleCaseElement(element, lower, culture, titlecasesDutchIJ, ref titleWordStart, ref titleDutchJ);
                    break;
                case OfficeTextCase.SentenceCase:
                    transformed = sentenceStart && IsLetter(element) ? lower.ToUpper(culture) : lower;
                    if (IsLetter(element)) sentenceStart = false;
                    if (EndsSentence(element)) sentenceStart = true;
                    break;
                case OfficeTextCase.ToggleCase:
                    string upper = element.ToUpper(culture);
                    transformed = string.Equals(element, upper, StringComparison.Ordinal) && !string.Equals(element, lower, StringComparison.Ordinal)
                        ? lower
                        : string.Equals(element, lower, StringComparison.Ordinal) && !string.Equals(element, upper, StringComparison.Ordinal)
                            ? upper
                            : element;
                    break;
                case OfficeTextCase.Capitalize:
                    transformed = capitalizeWordStart && IsLetter(element) ? element.ToUpper(culture) : element;
                    if (IsLetter(element) || IsCombiningMark(element)) capitalizeWordStart = false;
                    if (IsWordSeparator(element)) capitalizeWordStart = true;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(textCase), textCase, "Unsupported text casing transformation.");
            }
            AppendTransformedElement(result, boundaries, elements.ElementIndex, element, transformed);
        }

        return ToStrings(result);
    }

    private static void AppendTransformedElement(
        StringBuilder[] result,
        int[] boundaries,
        int sourceStart,
        string source,
        string transformed) {
        if (source.Length == transformed.Length) {
            AppendLengthPreservingSlice(result, boundaries, sourceStart, transformed, 0, transformed.Length);
            return;
        }

        int commonPrefix = 0;
        int commonLimit = Math.Min(source.Length, transformed.Length);
        while (commonPrefix < commonLimit && source[commonPrefix] == transformed[commonPrefix]) commonPrefix++;

        int commonSuffix = 0;
        while (commonSuffix < source.Length - commonPrefix &&
               commonSuffix < transformed.Length - commonPrefix &&
               source[source.Length - commonSuffix - 1] == transformed[transformed.Length - commonSuffix - 1]) {
            commonSuffix++;
        }

        if (commonPrefix > 0) {
            AppendLengthPreservingSlice(result, boundaries, sourceStart, transformed, 0, commonPrefix);
        }

        int transformedMiddleLength = transformed.Length - commonPrefix - commonSuffix;
        if (transformedMiddleLength > 0) {
            int changedSourceOffset = sourceStart + commonPrefix;
            int target = FindSegmentIndex(boundaries, changedSourceOffset);
            result[target].Append(transformed, commonPrefix, transformedMiddleLength);
        }

        if (commonSuffix > 0) {
            AppendLengthPreservingSlice(
                result,
                boundaries,
                sourceStart + source.Length - commonSuffix,
                transformed,
                transformed.Length - commonSuffix,
                commonSuffix);
        }
    }

    private static void AppendLengthPreservingSlice(
        StringBuilder[] result,
        int[] boundaries,
        int sourceStart,
        string transformed,
        int transformedStart,
        int length) {
        int sourceOffset = sourceStart;
        int transformedOffset = transformedStart;
        int remaining = length;
        while (remaining > 0) {
            int target = FindSegmentIndex(boundaries, sourceOffset);
            int available = Math.Min(remaining, boundaries[target] - sourceOffset);
            if (available <= 0) {
                target = Math.Min(target + 1, boundaries.Length - 1);
                available = Math.Min(remaining, boundaries[target] - sourceOffset);
            }
            result[target].Append(transformed, transformedOffset, available);
            sourceOffset += available;
            transformedOffset += available;
            remaining -= available;
        }
    }

    private static int FindSegmentIndex(int[] boundaries, int sourceOffset) {
        int low = 0;
        int high = boundaries.Length;
        while (low < high) {
            int middle = low + ((high - low) / 2);
            if (sourceOffset < boundaries[middle]) high = middle;
            else low = middle + 1;
        }
        return Math.Min(low, boundaries.Length - 1);
    }

    private static string TransformTitleCaseElement(
        string element,
        string lower,
        CultureInfo culture,
        bool titlecasesDutchIJ,
        ref bool wordStart,
        ref bool capitalizeDutchJ) {
        if (IsLetter(element)) {
            if (capitalizeDutchJ && string.Equals(lower, "j", StringComparison.Ordinal)) {
                capitalizeDutchJ = false;
                wordStart = false;
                return element.ToUpper(culture);
            }

            capitalizeDutchJ = false;
            if (wordStart) {
                wordStart = false;
                capitalizeDutchJ = titlecasesDutchIJ && string.Equals(lower, "i", StringComparison.Ordinal);
                return culture.TextInfo.ToTitleCase(lower);
            }
            return lower;
        }

        capitalizeDutchJ = false;
        if (IsTitleWordSeparator(element)) wordStart = true;
        return lower;
    }

    private static bool EndsSentence(string element) =>
        element.IndexOf('.') >= 0 || element.IndexOf('!') >= 0 || element.IndexOf('?') >= 0 ||
        element.IndexOf('\r') >= 0 || element.IndexOf('\n') >= 0;

    private static bool IsTitleWordSeparator(string textElement) {
        if (textElement == "'" || textElement == "’") return false;
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.SpaceSeparator ||
               category == UnicodeCategory.LineSeparator ||
               category == UnicodeCategory.ParagraphSeparator ||
               category == UnicodeCategory.Control ||
               category == UnicodeCategory.ConnectorPunctuation ||
               category == UnicodeCategory.DashPunctuation ||
               category == UnicodeCategory.OpenPunctuation ||
               category == UnicodeCategory.ClosePunctuation ||
               category == UnicodeCategory.InitialQuotePunctuation ||
               category == UnicodeCategory.FinalQuotePunctuation ||
               category == UnicodeCategory.OtherPunctuation ||
               category == UnicodeCategory.MathSymbol ||
               category == UnicodeCategory.CurrencySymbol ||
               category == UnicodeCategory.ModifierSymbol ||
               category == UnicodeCategory.OtherSymbol;
    }

    private static string[] ToStrings(StringBuilder[] builders) {
        var result = new string[builders.Length];
        for (int index = 0; index < builders.Length; index++) result[index] = builders[index].ToString();
        return result;
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
