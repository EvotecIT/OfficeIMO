using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Resolves UTF-16 break positions using versioned US English and reformed German patterns.</summary>
/// <remarks>
/// Resources are embedded and loaded once per language. Unknown languages, non-word tokens and
/// tokens longer than 512 UTF-16 code units return no automatic breaks. Leading/trailing punctuation
/// is retained; normalization never changes the caller's text or splits a grapheme.
/// </remarks>
public static partial class OfficeTextHyphenationPatterns {
    private static readonly Lazy<PatternDictionary> English = new Lazy<PatternDictionary>(
        () => PatternDictionary.Load("en-us", 2, 3));
    private static readonly Lazy<PatternDictionary> German = new Lazy<PatternDictionary>(
        () => PatternDictionary.Load("de-1996", 2, 2));

    /// <summary>Canonical language tags for the embedded resources.</summary>
    public static IReadOnlyList<string> SupportedLanguages { get; } = Array.AsReadOnly(new[] { "en-US", "de-DE" });

    /// <summary>Tests whether the tag selects an embedded pattern resource.</summary>
    /// <remarks><c>en</c> selects US English; <c>de</c>, <c>de-1996</c> and <c>de-DE-1996</c> select reformed German. Other regional or spelling tags are unsupported.</remarks>
    public static bool SupportsLanguage(string? language) => SelectLanguage(language) != null;

    /// <summary>Returns optional breaks in the original token, preserving case, punctuation and Unicode normalization.</summary>
    /// <param name="token">One token; embedded whitespace, digits and internal punctuation suppress automatic breaks.</param>
    /// <param name="language">A supported language tag. An empty or unsupported tag produces no breaks.</param>
    public static IReadOnlyList<int> GetBreakpoints(string token, string? language) {
        if (token == null) throw new ArgumentNullException(nameof(token));
        Lazy<PatternDictionary>? selected = SelectLanguage(language);
        if (selected == null || token.Length == 0 || token.Length > 512) return Array.Empty<int>();

        int first = 0, last = token.Length;
        while (first < last && char.IsPunctuation(token[first])) first++;
        while (last > first && char.IsPunctuation(token[last - 1])) last--;
        if (first == last) return Array.Empty<int>();
        string word = token.Substring(first, last - first);
        int[] starts = StringInfo.ParseCombiningCharacters(word);
        var normalized = new StringBuilder(word.Length);
        var sourceBoundaries = new List<int> { first };
        var characterCounts = new List<int> { 0 };
        for (int index = 0; index < starts.Length; index++) {
            int end = index + 1 < starts.Length ? starts[index + 1] : word.Length;
            string element = word.Substring(starts[index], end - starts[index]);
            for (int offset = 0; offset < element.Length; offset++) {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(element, offset);
                if (!char.IsLetter(element[offset]) && category != UnicodeCategory.NonSpacingMark
                    && category != UnicodeCategory.SpacingCombiningMark) return Array.Empty<int>();
            }
            element = element.Normalize(NormalizationForm.FormC).ToLowerInvariant();
            normalized.Append(element);
            for (int offset = 0; offset < element.Length; offset++) {
                bool boundary = offset + 1 == element.Length;
                sourceBoundaries.Add(boundary ? first + end : -1);
                characterCounts.Add(boundary ? index + 1 : -1);
            }
        }
        PatternDictionary dictionary = selected.Value;
        var result = new List<int>();
        foreach (int point in dictionary.GetBreakpoints(normalized.ToString())) {
            if (sourceBoundaries[point] >= 0 && characterCounts[point] >= dictionary.LeftMinimum
                && starts.Length - characterCounts[point] >= dictionary.RightMinimum) {
                result.Add(sourceBoundaries[point]);
            }
        }
        return result.ToArray();
    }

    private static Lazy<PatternDictionary>? SelectLanguage(string? language) {
        switch (language?.Trim().ToLowerInvariant()) {
            case "en":
            case "en-us": return English;
            case "de":
            case "de-de":
            case "de-1996":
            case "de-de-1996": return German;
            default: return null;
        }
    }
}
