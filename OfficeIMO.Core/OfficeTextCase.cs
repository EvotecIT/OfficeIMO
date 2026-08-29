using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
    private static readonly CultureInfo UnicodeTitlecaseFallbackCulture = CultureInfo.GetCultureInfo("en-US");
    // Unicode 17.0 PropList.txt Sentence_Terminal ranges. Keep this table independent of the
    // target runtime so net472, net8.0, and net10.0 apply the same sentence boundaries.
    private static readonly int[] SentenceTerminalRangeBounds = {
        0x0021, 0x0021, 0x002E, 0x002E, 0x003F, 0x003F, 0x0589, 0x0589,
        0x061D, 0x061F, 0x06D4, 0x06D4, 0x0700, 0x0702, 0x07F9, 0x07F9,
        0x0837, 0x0837, 0x0839, 0x0839, 0x083D, 0x083E, 0x0964, 0x0965,
        0x104A, 0x104B, 0x1362, 0x1362, 0x1367, 0x1368, 0x166E, 0x166E,
        0x1735, 0x1736, 0x17D4, 0x17D5, 0x1803, 0x1803, 0x1809, 0x1809,
        0x1944, 0x1945, 0x1AA8, 0x1AAB, 0x1B4E, 0x1B4F, 0x1B5A, 0x1B5B,
        0x1B5E, 0x1B5F, 0x1B7D, 0x1B7F, 0x1C3B, 0x1C3C, 0x1C7E, 0x1C7F,
        0x2024, 0x2024, 0x203C, 0x203D, 0x2047, 0x2049, 0x2CF9, 0x2CFB,
        0x2E2E, 0x2E2E, 0x2E3C, 0x2E3C, 0x2E53, 0x2E54, 0x3002, 0x3002,
        0xA4FF, 0xA4FF, 0xA60E, 0xA60F, 0xA6F3, 0xA6F3, 0xA6F7, 0xA6F7,
        0xA876, 0xA877, 0xA8CE, 0xA8CF, 0xA92F, 0xA92F, 0xA9C8, 0xA9C9,
        0xAA5D, 0xAA5F, 0xAAF0, 0xAAF1, 0xABEB, 0xABEB, 0xFE12, 0xFE12,
        0xFE15, 0xFE16, 0xFE52, 0xFE52, 0xFE56, 0xFE57, 0xFF01, 0xFF01,
        0xFF0E, 0xFF0E, 0xFF1F, 0xFF1F, 0xFF61, 0xFF61, 0x10A56, 0x10A57,
        0x10F55, 0x10F59, 0x10F86, 0x10F89, 0x11047, 0x11048, 0x110BE, 0x110C1,
        0x11141, 0x11143, 0x111C5, 0x111C6, 0x111CD, 0x111CD, 0x111DE, 0x111DF,
        0x11238, 0x11239, 0x1123B, 0x1123C, 0x112A9, 0x112A9, 0x113D4, 0x113D5,
        0x1144B, 0x1144C, 0x115C2, 0x115C3, 0x115C9, 0x115D7, 0x11641, 0x11642,
        0x1173C, 0x1173E, 0x11944, 0x11944, 0x11946, 0x11946, 0x11A42, 0x11A43,
        0x11A9B, 0x11A9C, 0x11C41, 0x11C42, 0x11EF7, 0x11EF8, 0x11F43, 0x11F44,
        0x16A6E, 0x16A6F, 0x16AF5, 0x16AF5, 0x16B37, 0x16B38, 0x16B44, 0x16B44,
        0x16D6E, 0x16D6F, 0x16E98, 0x16E98, 0x1BC9F, 0x1BC9F, 0x1DA88, 0x1DA88
    };

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
        string sourceText = source.ToString();
        string[] contextualLower = GetContextualLowerElements(sourceText, culture);
        int elementIndex = 0;
        TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(sourceText);
        while (elements.MoveNext()) {
            string element = elements.GetTextElement();
            string lower = contextualLower[elementIndex++];
            UnicodeCategory elementCategory = CharUnicodeInfo.GetUnicodeCategory(element, 0);
            if (elementCategory == UnicodeCategory.TitlecaseLetter && string.Equals(element, lower, StringComparison.Ordinal)) {
                // .NET Framework's invariant tables leave Unicode titlecase letters unchanged. Use a stable
                // Unicode-aware fallback so all supported target frameworks produce the same lowercase pair.
                lower = element.ToLower(UnicodeTitlecaseFallbackCulture);
            }
            bool isCasedLetter = IsCasedLetter(element, lower, culture);
            string transformed;
            switch (textCase) {
                case OfficeTextCase.Uppercase:
                    transformed = element.ToUpper(culture);
                    break;
                case OfficeTextCase.Lowercase:
                    transformed = lower;
                    break;
                case OfficeTextCase.TitleCase:
                    transformed = TransformTitleCaseElement(element, lower, culture, titlecasesDutchIJ, isCasedLetter, ref titleWordStart, ref titleDutchJ);
                    break;
                case OfficeTextCase.SentenceCase:
                    transformed = sentenceStart && isCasedLetter ? lower.ToUpper(culture) : lower;
                    if (isCasedLetter) sentenceStart = false;
                    if (EndsSentence(element)) sentenceStart = true;
                    break;
                case OfficeTextCase.ToggleCase:
                    string upper = element.ToUpper(culture);
                    transformed = elementCategory == UnicodeCategory.UppercaseLetter || elementCategory == UnicodeCategory.TitlecaseLetter
                        ? lower
                        : elementCategory == UnicodeCategory.LowercaseLetter
                            ? upper
                            : string.Equals(element, upper, StringComparison.Ordinal) && !string.Equals(element, lower, StringComparison.Ordinal)
                                ? lower
                                : string.Equals(element, lower, StringComparison.Ordinal) && !string.Equals(element, upper, StringComparison.Ordinal)
                                    ? upper
                                    : element;
                    break;
                case OfficeTextCase.Capitalize:
                    transformed = capitalizeWordStart && isCasedLetter ? element.ToUpper(culture) : element;
                    if (isCasedLetter || IsCombiningMark(element)) capitalizeWordStart = false;
                    if (IsWordSeparator(element)) capitalizeWordStart = true;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(textCase), textCase, "Unsupported text casing transformation.");
            }
            AppendTransformedElement(result, boundaries, elements.ElementIndex, element, transformed);
        }

        return ToStrings(result);
    }

    private static string[] GetContextualLowerElements(string source, CultureInfo culture) {
        var sourceElements = new List<string>();
        TextElementEnumerator sourceEnumerator = StringInfo.GetTextElementEnumerator(source);
        while (sourceEnumerator.MoveNext()) sourceElements.Add(sourceEnumerator.GetTextElement());

        var lowerElements = new List<string>();
        TextElementEnumerator lowerEnumerator = StringInfo.GetTextElementEnumerator(source.ToLower(culture));
        while (lowerEnumerator.MoveNext()) lowerElements.Add(lowerEnumerator.GetTextElement());
        if (lowerElements.Count != sourceElements.Count) {
            lowerElements.Clear();
            lowerElements.AddRange(sourceElements.Select(element => element.ToLower(culture)));
        }
        for (int index = 0; index < sourceElements.Count; index++) {
            if (string.Equals(sourceElements[index], "Σ", StringComparison.Ordinal) &&
                HasCasedLetterBefore(sourceElements, index, culture) &&
                !HasCasedLetterAfter(sourceElements, index, culture)) {
                lowerElements[index] = "ς";
            }
            if (CharUnicodeInfo.GetUnicodeCategory(sourceElements[index], 0) == UnicodeCategory.TitlecaseLetter &&
                string.Equals(sourceElements[index], lowerElements[index], StringComparison.Ordinal)) {
                lowerElements[index] = sourceElements[index].ToLower(UnicodeTitlecaseFallbackCulture);
            }
        }
        return lowerElements.ToArray();
    }

    private static bool HasCasedLetterBefore(IReadOnlyList<string> elements, int index, CultureInfo culture) {
        for (int candidate = index - 1; candidate >= 0; candidate--) {
            string element = elements[candidate];
            if (IsCaseIgnorable(element)) continue;
            return IsCasedLetter(element, element.ToLower(culture), culture);
        }
        return false;
    }

    private static bool HasCasedLetterAfter(IReadOnlyList<string> elements, int index, CultureInfo culture) {
        for (int candidate = index + 1; candidate < elements.Count; candidate++) {
            string element = elements[candidate];
            if (IsCaseIgnorable(element)) continue;
            return IsCasedLetter(element, element.ToLower(culture), culture);
        }
        return false;
    }

    private static bool IsCaseIgnorable(string textElement) {
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(textElement, 0);
        return category == UnicodeCategory.NonSpacingMark ||
               category == UnicodeCategory.EnclosingMark ||
               category == UnicodeCategory.Format ||
               category == UnicodeCategory.ModifierLetter ||
               category == UnicodeCategory.ModifierSymbol ||
               textElement == "'" ||
               textElement == "’";
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
        bool isCasedLetter,
        ref bool wordStart,
        ref bool capitalizeDutchJ) {
        if (isCasedLetter) {
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

    private static bool IsCasedLetter(string element, string lower, CultureInfo culture) {
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(element, 0);
        if (category is UnicodeCategory.UppercaseLetter or UnicodeCategory.LowercaseLetter or UnicodeCategory.TitlecaseLetter) {
            return true;
        }
        return !string.Equals(element.ToUpper(culture), lower, StringComparison.Ordinal);
    }

    private static bool EndsSentence(string element) {
        for (int index = 0; index < element.Length; index++) {
            int codePoint = char.ConvertToUtf32(element, index);
            if (codePoint == '\r' || codePoint == '\n' || IsSentenceTerminal(codePoint)) return true;
            if (codePoint > char.MaxValue) index++;
        }
        return false;
    }

    private static bool IsSentenceTerminal(int codePoint) {
        int low = 0;
        int high = SentenceTerminalRangeBounds.Length / 2 - 1;
        while (low <= high) {
            int middle = low + (high - low) / 2;
            int start = SentenceTerminalRangeBounds[middle * 2];
            int end = SentenceTerminalRangeBounds[middle * 2 + 1];
            if (codePoint < start) high = middle - 1;
            else if (codePoint > end) low = middle + 1;
            else return true;
        }
        return false;
    }

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
