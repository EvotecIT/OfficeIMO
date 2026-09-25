using System.Linq;

namespace OfficeIMO.Pdf;

/// <summary>
/// Shared literal-search matching for text search, located edits, and redaction search. Whitespace runs, including line
/// breaks, match one query space, and a letter-hyphen-line-break-letter junction matches both the joined word
/// ("hyphenation") and the hyphenated compound without the break ("well-known").
/// </summary>
internal static class PdfTextSearchNormalization {
    /// <summary>Collapses whitespace and retains a line-end hyphen in the query.</summary>
    internal static string NormalizeQuery(string text) => NormalizeQueries(text)[0];

    /// <summary>Also accepts the joined spelling when the query itself contains a hyphenated line break.</summary>
    internal static string[] NormalizeQueries(string text) {
        int[] lineBreaks = GetLineBreakCounts(text);
        PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
        if (!hasHyphenJunction) return new[] { joined.Text };
        string dehyphenated = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: true, out _).Text;
        return new[] { joined.Text, dehyphenated };
    }

    /// <summary>Returns true when <paramref name="text"/> contains <paramref name="value"/> under line-break and hyphenation-tolerant matching.</summary>
    internal static bool Contains(string text, string value, StringComparison comparison) {
        if (ContainsExact(text, value, comparison)) return true;
        string[] queries = NormalizeQueries(value);
        if (queries[0].Length == 0) return false;
        bool hasWhitespace = false;
        for (int index = 0; index < text.Length && !hasWhitespace; index++) hasWhitespace = char.IsWhiteSpace(text[index]);
        if (!hasWhitespace) return queries.Any(query => ContainsExact(text, query, comparison));
        int[] lineBreaks = GetLineBreakCounts(text);
        PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
        if (queries.Any(query => ContainsExact(joined.Text, query, comparison))) return true;
        return hasHyphenJunction && queries.Any(query =>
            ContainsExact(PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: true, out _).Text, query, comparison));
    }

    /// <summary>Enumerates source ranges without collecting dense occurrences before a caller can stop.</summary>
    internal static IEnumerable<(int Start, int End)> FindSourceRanges(string text, string value, StringComparison comparison) {
        string[] queries = NormalizeQueries(value);
        if (queries[0].Length == 0 || text.Length == 0) yield break;
        int[] lineBreaks = GetLineBreakCounts(text);
        PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
        foreach (string query in queries)
            foreach ((int Start, int End) range in EnumerateSourceRanges(joined, query, comparison)) yield return range;
        if (hasHyphenJunction) {
            PdfNormalizedSearchText dehyphenated = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: true, out _);
            foreach (string query in queries)
                foreach ((int Start, int End) range in EnumerateSourceRanges(dehyphenated, query, comparison)) yield return range;
        }
    }

    private static IEnumerable<(int Start, int End)> EnumerateSourceRanges(PdfNormalizedSearchText normalized, string query, StringComparison comparison) {
        int start = 0;
        while (start <= normalized.Text.Length - query.Length) {
            int found = normalized.Text.IndexOf(query, start, comparison);
            if (found < 0) break;
            yield return (normalized.GetSourceStart(found), normalized.GetSourceEnd(found + query.Length - 1));
            start = found + 1;
        }
    }

    private static int[] GetLineBreakCounts(string text) {
        var lineBreaks = new int[text.Length];
        for (int index = 0; index < text.Length; index++) {
            if (text[index] == '\r' && index + 1 < text.Length && text[index + 1] == '\n') continue;
            if (text[index] is '\r' or '\n' or '\u2028' or '\u2029') lineBreaks[index] = 1;
        }
        return lineBreaks;
    }

    /// <summary>Plain substring containment without whitespace or hyphenation normalization.</summary>
    internal static bool ContainsExact(string text, string value, StringComparison comparison) {
#if NET6_0_OR_GREATER
        return text.Contains(value, comparison);
#else
        return text.IndexOf(value, comparison) >= 0;
#endif
    }
}

/// <summary>Whitespace-collapsed search text that records the source range each normalized character covers.</summary>
internal sealed class PdfNormalizedSearchText {
    private readonly int[] _sourceStarts;
    private readonly int[] _sourceEnds;

    private PdfNormalizedSearchText(string text, int[] sourceStarts, int[] sourceEnds) {
        Text = text; _sourceStarts = sourceStarts; _sourceEnds = sourceEnds;
    }

    internal string Text { get; }

    internal int GetSourceStart(int index) => _sourceStarts[index];

    internal int GetSourceEnd(int index) => _sourceEnds[index];

    /// <summary>
    /// Collapses whitespace runs to one space. A run that contains a line break and follows a letter-hyphen pair before
    /// another letter is dropped when there is exactly one logical line break; with <paramref name="removeLineEndHyphens"/> the hyphen is dropped too.
    /// </summary>
    internal static PdfNormalizedSearchText Create(string source, int[] lineBreaks, bool removeLineEndHyphens, out bool hasHyphenJunction) {
        hasHyphenJunction = false;
        var text = new System.Text.StringBuilder(source.Length);
        var starts = new List<int>(source.Length);
        var ends = new List<int>(source.Length);
        int index = 0;
        while (index < source.Length) {
            if (!char.IsWhiteSpace(source[index])) {
                text.Append(source[index]);
                starts.Add(index);
                ends.Add(index + 1);
                index++;
                continue;
            }
            int runEnd = index;
            int lineBreakCount = 0;
            while (runEnd < source.Length && char.IsWhiteSpace(source[runEnd])) {
                lineBreakCount += lineBreaks[runEnd];
                runEnd++;
            }
            if (lineBreakCount == 1 && IsLineEndHyphenJunction(source, index, runEnd)) {
                hasHyphenJunction = true;
                if (removeLineEndHyphens) {
                    text.Length--;
                    starts.RemoveAt(starts.Count - 1);
                    ends.RemoveAt(ends.Count - 1);
                } else {
                    // Keep the source range, but match every supported line-end hyphen against
                    // the ordinary hyphen callers use in copied or typed queries.
                    text[text.Length - 1] = '-';
                }
            } else {
                text.Append(' ');
                starts.Add(index);
                ends.Add(runEnd);
            }
            index = runEnd;
        }
        return new PdfNormalizedSearchText(text.ToString(), starts.ToArray(), ends.ToArray());
    }

    private static bool IsLineEndHyphenJunction(string source, int runStart, int runEnd) {
        if (runStart < 2 || runEnd >= source.Length || !IsLineEndHyphen(source[runStart - 1]) ||
            !char.IsLetter(source, runEnd)) return false;
        int letterIndex = runStart - 2;
        while (letterIndex >= 0 && IsCombiningMark(source, letterIndex)) letterIndex--;
        if (letterIndex > 0 && char.IsLowSurrogate(source[letterIndex]) && char.IsHighSurrogate(source[letterIndex - 1]))
            letterIndex--;
        return letterIndex >= 0 && char.IsLetter(source, letterIndex);
    }

    private static bool IsCombiningMark(string source, int index) => char.GetUnicodeCategory(source, index) is
        System.Globalization.UnicodeCategory.NonSpacingMark or
        System.Globalization.UnicodeCategory.SpacingCombiningMark or
        System.Globalization.UnicodeCategory.EnclosingMark;

    private static bool IsLineEndHyphen(char value) => (int)value is 0x002D or 0x00AD or 0x2010 or 0x2011;
}
