using System.Text;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Where text replacement lands in a group of adjacent runs, whatever holds the runs: the SDK model and the Open XML
/// layer both apply this plan, so they replace the same matches the same way.
/// </summary>
internal static class PowerPointRunReplace {
    /// <summary>
    /// Finds the matches in the runs' joined text and returns their number; <paramref name="updated"/> receives the
    /// new text of each run a match touches and null for the others. A match takes the replacement in the run where it
    /// starts; later runs lose only their part of it. <paramref name="preceding"/> and <paramref name="following"/>
    /// are the text around the group, such as a field's, and count only for word boundaries.
    /// </summary>
    internal static int Plan(IReadOnlyList<string> runs, string text, string? replacement, StringComparison comparison,
        bool wholeWords, string? preceding, string? following, string?[] updated) {
        Array.Clear(updated, 0, runs.Count);
        if (runs.Count == 1) {
            int occurrences = 0;
            string result = OfficeTextSearch.Replace(runs[0], text, replacement, comparison, wholeWords, ref occurrences, preceding, following);
            if (occurrences > 0) updated[0] = result;
            return occurrences;
        }

        string original = string.Concat(runs);
        List<int>? matches = null;
        int searchStart = 0;
        while (searchStart <= original.Length - text.Length) {
            int found = original.IndexOf(text, searchStart, comparison);
            if (found < 0) break;
            if (wholeWords && !OfficeTextSearch.HasWordBoundaries(original, found, text.Length, preceding, following)) {
                searchStart = found + 1;
                continue;
            }

            (matches ??= new List<int>()).Add(found);
            searchStart = found + text.Length;
        }

        if (matches == null) return 0;

        // Offsets are in the original text, and each touched run is written once.
        int match = 0;
        int runStart = 0;
        for (int index = 0; index < runs.Count; index++) {
            if (match == matches.Count) break;

            string current = runs[index];
            int runEnd = runStart + current.Length;
            if (current.Length > 0 && matches[match] < runEnd) {
                var builder = new StringBuilder(current.Length);
                int copied = 0;
                while (match < matches.Count && matches[match] < runEnd) {
                    int startInRun = matches[match] - runStart;
                    if (startInRun >= 0) {
                        builder.Append(current, copied, startInRun - copied);
                        builder.Append(replacement);
                    }

                    int endInRun = startInRun + text.Length;
                    copied = Math.Min(endInRun, current.Length);
                    if (endInRun > current.Length) break;
                    match++;
                }
                builder.Append(current, copied, current.Length - copied);
                updated[index] = builder.ToString();
            }
            runStart = runEnd;
        }

        return matches.Count;
    }
}
