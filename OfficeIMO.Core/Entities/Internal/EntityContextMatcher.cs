using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Entities.Internal;

internal sealed class EntityKeyword {
    internal EntityKeyword(string schemeId, string language, string value) {
        SchemeId = schemeId;
        Language = language;
        Value = value;
        Folded = EntityContextMatcher.Fold(value).TrimEnd(' ', '.', ':', '#');
    }

    internal string SchemeId { get; }
    internal string Language { get; }
    internal string Value { get; }
    internal string Folded { get; }
}

internal static class EntityContextMatcher {
    private const string Connectors = " nr no number numer id is to of the est de du del den der n o nummer numero n° nº ";

    internal static EntityKeyword? Find(List<EntityKeyword>? keywords, HashSet<string>? languages, string? folded) {
        if (keywords == null || folded == null) return null;
        EntityKeyword? best = null;
        int bestEnd = -1;
        foreach (EntityKeyword keyword in keywords) {
            if (keyword.Folded.Length == 0 || languages != null && !languages.Contains(keyword.Language)) continue;
            int end = FindKeywordEnd(folded, keyword.Folded);
            if (end < 0 || end < bestEnd || end == bestEnd && best != null && best.Folded.Length >= keyword.Folded.Length) continue;
            best = keyword;
            bestEnd = end;
        }
        return best;
    }

    private static int FindKeywordEnd(string value, string keyword) {
        int searchEnd = value.Length;
        while (searchEnd >= keyword.Length) {
            int start = value.LastIndexOf(keyword, searchEnd - 1, StringComparison.Ordinal);
            if (start < 0) break;
            int end = start + keyword.Length;
            if ((start == 0 || !IsWord(value[start - 1])) && (end == value.Length || !IsWord(value[end])) &&
                HasConnectingWordsOnly(value, end)) return end;

            // Move the matching window one position to the left of the occurrence.
            searchEnd = end - 1;
        }
        return -1;
    }

    internal static string Fold(string value) {
        string decomposed;
        try {
            decomposed = value.Normalize(NormalizationForm.FormD);
        } catch (ArgumentException) {
            // Invalid UTF-16 in a host's text must not prevent other entities from being found.
            decomposed = value;
        }
        var folded = new StringBuilder(decomposed.Length);
        foreach (char character in decomposed) {
            if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.NonSpacingMark) continue;
            char lower = char.ToLowerInvariant(character);
            folded.Append(char.IsWhiteSpace(lower) ? ' ' : lower == 'ł' ? 'l' : lower);
        }
        return folded.ToString();
    }

    private static bool HasConnectingWordsOnly(string value, int start) {
        int words = 0;
        for (int index = start; index < value.Length;) {
            if (!IsWord(value[index])) {
                index++;
                continue;
            }
            int wordStart = index++;
            while (index < value.Length && IsWord(value[index])) index++;
            if (++words > 4 || index - wordStart > 6 ||
                Connectors.IndexOf(" " + value.Substring(wordStart, index - wordStart) + " ", StringComparison.Ordinal) < 0) {
                return false;
            }
        }
        return true;
    }

    internal static bool IsWord(char value) => char.IsLetterOrDigit(value) || CharUnicodeInfo.GetUnicodeCategory(value) is
        UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.ConnectorPunctuation || char.IsSurrogate(value);
}
