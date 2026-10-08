using System.Text.Json;

namespace OfficeIMO.Bibliography;

/// <summary>Local, versioned CSL stop words and phrases for English title casing.</summary>
internal static class CslTitleCaseWords {
    private static readonly IReadOnlyDictionary<string, string[]> Words = ReadWords();

    internal static int MatchLength(string source, int start, string firstWord) {
        if (!Words.TryGetValue(firstWord, out string[]? phrases)) return 0;
        foreach (string phrase in phrases) {
            int end = start + phrase.Length;
            if (end > source.Length || string.Compare(source, start, phrase, 0, phrase.Length, StringComparison.OrdinalIgnoreCase) != 0) continue;
            if (end < source.Length && (char.IsLetterOrDigit(source[end]) || char.GetUnicodeCategory(source[end]) == UnicodeCategory.NonSpacingMark)) continue;
            return phrase.Length;
        }
        return 0;
    }

    private static IReadOnlyDictionary<string, string[]> ReadWords() {
        using Stream stream = typeof(CslTitleCaseWords).Assembly.GetManifestResourceStream("OfficeIMO.Bibliography.Citations.Data.stop-words.json")!;
        using JsonDocument json = JsonDocument.Parse(stream);
        return json.RootElement.GetProperty("stop-words").EnumerateArray().Select(word => word.GetString()!)
            .GroupBy(word => word.Split(new[] { ' ', '-' })[0], StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.OrderByDescending(word => word.Length).ToArray(), StringComparer.OrdinalIgnoreCase);
    }
}
