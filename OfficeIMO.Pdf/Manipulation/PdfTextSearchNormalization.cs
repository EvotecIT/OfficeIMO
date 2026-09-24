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
        bool[] lineBreaks = GetLineBreaks(text);
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
        bool[] lineBreaks = GetLineBreaks(text);
        PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
        if (queries.Any(query => ContainsExact(joined.Text, query, comparison))) return true;
        return hasHyphenJunction && queries.Any(query =>
            ContainsExact(PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: true, out _).Text, query, comparison));
    }

    /// <summary>Returns the source ranges of every line-break and hyphenation-tolerant occurrence of <paramref name="value"/>.</summary>
    internal static List<(int Start, int End)> FindSourceRanges(string text, string value, StringComparison comparison) {
        var ranges = new List<(int Start, int End)>();
        string[] queries = NormalizeQueries(value);
        if (queries[0].Length == 0 || text.Length == 0) return ranges;
        bool[] lineBreaks = GetLineBreaks(text);
        PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
        foreach (string query in queries) AddSourceRanges(joined, query, comparison, ranges);
        if (hasHyphenJunction) {
            PdfNormalizedSearchText dehyphenated = PdfNormalizedSearchText.Create(text, lineBreaks, removeLineEndHyphens: true, out _);
            foreach (string query in queries) AddSourceRanges(dehyphenated, query, comparison, ranges);
        }
        return ranges;
    }

    private static void AddSourceRanges(PdfNormalizedSearchText normalized, string query, StringComparison comparison, List<(int Start, int End)> ranges) {
        int start = 0;
        while (start <= normalized.Text.Length - query.Length) {
            int found = normalized.Text.IndexOf(query, start, comparison);
            if (found < 0) break;
            ranges.Add((normalized.GetSourceStart(found), normalized.GetSourceEnd(found + query.Length - 1)));
            start = found + 1;
        }
    }

    private static bool[] GetLineBreaks(string text) {
        var lineBreaks = new bool[text.Length];
        for (int index = 0; index < text.Length; index++) lineBreaks[index] = (int)text[index] is 0x000A or 0x000D or 0x2028 or 0x2029;
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
    /// another letter is dropped; with <paramref name="removeLineEndHyphens"/> the hyphen is dropped too, so the word reads joined.
    /// </summary>
    internal static PdfNormalizedSearchText Create(string source, bool[] lineBreaks, bool removeLineEndHyphens, out bool hasHyphenJunction) {
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
            bool containsLineBreak = false;
            while (runEnd < source.Length && char.IsWhiteSpace(source[runEnd])) {
                containsLineBreak |= lineBreaks[runEnd];
                runEnd++;
            }
            if (containsLineBreak && IsLineEndHyphenJunction(source, index, runEnd)) {
                hasHyphenJunction = true;
                if (removeLineEndHyphens) {
                    text.Length--;
                    starts.RemoveAt(starts.Count - 1);
                    ends.RemoveAt(ends.Count - 1);
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

    private static bool IsLineEndHyphenJunction(string source, int runStart, int runEnd) =>
        runStart >= 2 &&
        runEnd < source.Length &&
        IsLineEndHyphen(source[runStart - 1]) &&
        char.IsLetter(source, runStart - 2) &&
        char.IsLetter(source, runEnd);

    private static bool IsLineEndHyphen(char value) => (int)value is 0x002D or 0x00AD or 0x2010 or 0x2011;
}
