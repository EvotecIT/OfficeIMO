using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO;

/// <summary>
/// Comparison settings shared by the document <c>Find</c> and <c>ReplaceAll</c> operations in Word, PowerPoint and PDF.
/// </summary>
/// <remarks>
/// The defaults are the same in every format: matching is ordinal and case-insensitive, and a match may start or end
/// inside a word. Each format applies these settings to its own text model; see the format's <c>Find</c> and
/// <c>ReplaceAll</c> documentation for the scope a single match can span.
/// </remarks>
public class OfficeTextSearchOptions {
    /// <summary>Uses ordinal case-sensitive matching when <c>true</c>. The default, <c>false</c>, matches ordinally ignoring case.</summary>
    public bool MatchCase { get; set; }

    /// <summary>
    /// Requires a match to start and end at a word boundary when <c>true</c>. A word character is a letter, digit,
    /// combining mark or underscore. The default is <c>false</c>.
    /// </summary>
    public bool WholeWords { get; set; }
}

/// <summary>Literal matching helpers that apply <see cref="OfficeTextSearchOptions"/> to plain strings.</summary>
internal static class OfficeTextSearch {
    internal static StringComparison GetComparison(OfficeTextSearchOptions? options) =>
        options?.MatchCase == true ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

    internal static void ValidateSearchText(string text, string parameterName) {
        if (text == null) throw new ArgumentNullException(parameterName);
        if (text.Length == 0) throw new ArgumentException("Search text cannot be empty.", parameterName);
    }

    /// <summary>Counts matches of <paramref name="text"/> in <paramref name="value"/>.</summary>
    /// <param name="value">Text to search.</param>
    /// <param name="text">Literal text to find.</param>
    /// <param name="comparison">Comparison used for matching.</param>
    /// <param name="wholeWords">Requires word boundaries around each match.</param>
    /// <param name="precedingText">Text immediately before <paramref name="value"/>, used for the leading word boundary.</param>
    /// <param name="followingText">Text immediately after <paramref name="value"/>, used for the trailing word boundary.</param>
    internal static int Count(string value, string text, StringComparison comparison, bool wholeWords, string? precedingText = null, string? followingText = null) {
        int count = 0;
        int start = 0;
        while (start <= value.Length - text.Length) {
            int found = value.IndexOf(text, start, comparison);
            if (found < 0) break;
            if (wholeWords && !HasWordBoundaries(value, found, text.Length, precedingText, followingText)) {
                start = found + 1;
                continue;
            }
            count++;
            start = found + text.Length;
        }
        return count;
    }

    /// <summary>Replaces every match of <paramref name="text"/> in <paramref name="value"/>.</summary>
    /// <returns><paramref name="value"/> itself when nothing matched; otherwise the replaced string.</returns>
    internal static string Replace(string value, string text, string? replacement, StringComparison comparison, bool wholeWords, ref int count, string? precedingText = null, string? followingText = null) {
        StringBuilder? builder = null;
        int copied = 0;
        int start = 0;
        while (start <= value.Length - text.Length) {
            int found = value.IndexOf(text, start, comparison);
            if (found < 0) break;
            if (wholeWords && !HasWordBoundaries(value, found, text.Length, precedingText, followingText)) {
                start = found + 1;
                continue;
            }
            builder ??= new StringBuilder(value.Length);
            builder.Append(value, copied, found - copied);
            builder.Append(replacement);
            count++;
            copied = start = found + text.Length;
        }
        if (builder == null) return value;
        builder.Append(value, copied, value.Length - copied);
        return builder.ToString();
    }

    internal static bool HasWordBoundaries(string value, int start, int length, string? precedingText = null, string? followingText = null) {
        bool left = start > 0
            ? !IsWordCharacter(value, start - 1)
            : string.IsNullOrEmpty(precedingText) || !IsWordCharacter(precedingText!, precedingText!.Length - 1);
        if (!left) return false;
        int end = start + length;
        return end < value.Length
            ? !IsWordCharacter(value, end)
            : string.IsNullOrEmpty(followingText) || !IsWordCharacter(followingText!, 0);
    }

    internal static bool IsWordCharacter(string text, int index) {
        if (index > 0 && char.IsLowSurrogate(text[index]) && char.IsHighSurrogate(text[index - 1])) index--;
        if (text[index] == '_' || char.IsLetterOrDigit(text, index)) return true;
        UnicodeCategory category = char.GetUnicodeCategory(text, index);
        return category is UnicodeCategory.NonSpacingMark or
            UnicodeCategory.SpacingCombiningMark or
            UnicodeCategory.EnclosingMark;
    }
}
