namespace OfficeIMO.Diff;

/// <summary>
/// Computes edit scripts between sequences: generic items with a comparer, or text split into lines, words or characters.
/// </summary>
/// <remarks>
/// Items are interned to integer ids first, so the comparer is called once per item rather than once per comparison.
/// Every result is complete: when a work or time limit stops the alignment, the remaining regions are reported as
/// whole deletions and insertions and <see cref="OfficeDiffResult{T}.LimitReached"/> is set.
/// </remarks>
public static class OfficeDiff {
    /// <summary>Compares two sequences.</summary>
    /// <param name="oldItems">Old sequence.</param>
    /// <param name="newItems">New sequence.</param>
    /// <param name="comparer">Item comparer. Defaults to <see cref="EqualityComparer{T}.Default"/>.</param>
    /// <param name="options">Diff settings.</param>
    /// <param name="cancellationToken">Cancels the alignment.</param>
    public static OfficeDiffResult<T> Compare<T>(IReadOnlyList<T> oldItems, IReadOnlyList<T> newItems, IEqualityComparer<T>? comparer = null,
        OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        Compare(oldItems, newItems, comparer, options, weight: null, cancellationToken);

    /// <summary>Compares two texts line by line. Line terminators are part of each line.</summary>
    public static OfficeDiffResult<OfficeTextToken> CompareLines(string? oldText, string? newText, OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        CompareTokens(OfficeTextTokenizer.Lines(oldText), OfficeTextTokenizer.Lines(newText), OfficeTextTokenComparer.ForLines(options), options, cancellationToken);

    /// <summary>Compares two texts word by word (see <see cref="OfficeTextTokenizer.Words"/>).</summary>
    public static OfficeDiffResult<OfficeTextToken> CompareWords(string? oldText, string? newText, OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        CompareTokens(OfficeTextTokenizer.Words(oldText), OfficeTextTokenizer.Words(newText), OfficeTextTokenComparer.ForWords(options), options, cancellationToken);

    /// <summary>Compares two texts by user-perceived character (grapheme cluster).</summary>
    public static OfficeDiffResult<OfficeTextToken> CompareCharacters(string? oldText, string? newText, OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        CompareTokens(OfficeTextTokenizer.Characters(oldText), OfficeTextTokenizer.Characters(newText), OfficeTextTokenComparer.ForCharacters(options), options, cancellationToken);

    /// <summary>Compares two token sequences produced by <see cref="OfficeTextTokenizer"/> or by a caller, honouring the text settings in <paramref name="options"/>.</summary>
    public static OfficeDiffResult<OfficeTextToken> CompareTokens(IReadOnlyList<OfficeTextToken> oldTokens, IReadOnlyList<OfficeTextToken> newTokens,
        OfficeDiffOptions? options = null, CancellationToken cancellationToken = default) =>
        CompareTokens(oldTokens, newTokens, OfficeTextTokenComparer.ForWords(options), options, cancellationToken);

    private static OfficeDiffResult<OfficeTextToken> CompareTokens(IReadOnlyList<OfficeTextToken> oldTokens, IReadOnlyList<OfficeTextToken> newTokens,
        OfficeTextTokenComparer comparer, OfficeDiffOptions? options, CancellationToken cancellationToken) =>
        Compare(oldTokens, newTokens, comparer, options, static token => Math.Max(1, token.Length), cancellationToken);

    internal static OfficeDiffResult<T> Compare<T>(IReadOnlyList<T> oldItems, IReadOnlyList<T> newItems, IEqualityComparer<T>? comparer,
        OfficeDiffOptions? options, Func<T, int>? weight, CancellationToken cancellationToken) {
        if (oldItems is null) throw new ArgumentNullException(nameof(oldItems));
        if (newItems is null) throw new ArgumentNullException(nameof(newItems));
        comparer ??= EqualityComparer<T>.Default;
        options ??= new OfficeDiffOptions();

        int[] oldIds = new int[oldItems.Count];
        int[] newIds = new int[newItems.Count];
        Intern(oldItems, newItems, comparer, oldIds, newIds);
        OfficeDiffEngine engine = OfficeDiffEngine.Run(oldIds, newIds, options, cancellationToken);
        IReadOnlyList<OfficeDiffMatch> matches = engine.Matches;
        if (options.SemanticCleanup && matches.Count > 0) {
            Func<T, int> measure = weight ?? (static _ => 1);
            matches = CleanupSemantic(matches, oldItems, newItems, measure);
        }

        IReadOnlyList<OfficeDiffEdit> edits = BuildEdits(matches, oldItems.Count, newItems.Count);
        return new OfficeDiffResult<T>(oldItems, newItems, edits, comparer, options.Algorithm, engine.LimitReached, engine.WorkUnits);
    }

    private static void Intern<T>(IReadOnlyList<T> oldItems, IReadOnlyList<T> newItems, IEqualityComparer<T> comparer, int[] oldIds, int[] newIds) {
#pragma warning disable CS8714 // Null items never reach the dictionary; they get their own id below.
        var ids = new Dictionary<T, int>(comparer);
#pragma warning restore CS8714
        int nullId = -1;
        int next = 0;
        int IdOf(T item) {
            if (item is null) {
                if (nullId < 0) nullId = next++;
                return nullId;
            }

            if (!ids.TryGetValue(item, out int id)) {
                id = next++;
                ids.Add(item, id);
            }

            return id;
        }

        for (int index = 0; index < oldItems.Count; index++) oldIds[index] = IdOf(oldItems[index]);
        for (int index = 0; index < newItems.Count; index++) newIds[index] = IdOf(newItems[index]);
    }

    private static IReadOnlyList<OfficeDiffEdit> BuildEdits(IReadOnlyList<OfficeDiffMatch> matches, int oldCount, int newCount) {
        var edits = new List<OfficeDiffEdit>(matches.Count * 3 + 2);
        int oldCursor = 0;
        int newCursor = 0;
        foreach (OfficeDiffMatch match in matches) {
            AddGap(edits, oldCursor, match.OldStart, newCursor, match.NewStart);
            edits.Add(new OfficeDiffEdit(OfficeDiffOperation.Equal, match.OldStart, match.Length, match.NewStart, match.Length));
            oldCursor = match.OldStart + match.Length;
            newCursor = match.NewStart + match.Length;
        }

        AddGap(edits, oldCursor, oldCount, newCursor, newCount);
        return edits;
    }

    private static void AddGap(List<OfficeDiffEdit> edits, int oldStart, int oldEnd, int newStart, int newEnd) {
        if (oldEnd > oldStart) edits.Add(new OfficeDiffEdit(OfficeDiffOperation.Delete, oldStart, oldEnd - oldStart, newStart, 0));
        if (newEnd > newStart) edits.Add(new OfficeDiffEdit(OfficeDiffOperation.Insert, oldEnd, 0, newStart, newEnd - newStart));
    }

    // Removes an unchanged run when it is no heavier than the changes on either side of it, so
    // "the quick fox" -> "a slow fox" reads as one replacement rather than word, space, word fragments.
    private static IReadOnlyList<OfficeDiffMatch> CleanupSemantic<T>(IReadOnlyList<OfficeDiffMatch> source, IReadOnlyList<T> oldItems,
        IReadOnlyList<T> newItems, Func<T, int> weight) {
        var matches = new List<OfficeDiffMatch>(source);
        long[] oldPrefix = PrefixWeights(oldItems, weight);
        long[] newPrefix = PrefixWeights(newItems, weight);
        int index = 0;
        while (index < matches.Count) {
            OfficeDiffMatch match = matches[index];
            int oldBefore = index == 0 ? 0 : matches[index - 1].OldStart + matches[index - 1].Length;
            int newBefore = index == 0 ? 0 : matches[index - 1].NewStart + matches[index - 1].Length;
            int oldAfter = index == matches.Count - 1 ? oldItems.Count : matches[index + 1].OldStart;
            int newAfter = index == matches.Count - 1 ? newItems.Count : matches[index + 1].NewStart;
            long before = Math.Max(oldPrefix[match.OldStart] - oldPrefix[oldBefore], newPrefix[match.NewStart] - newPrefix[newBefore]);
            long after = Math.Max(oldPrefix[oldAfter] - oldPrefix[match.OldStart + match.Length], newPrefix[newAfter] - newPrefix[match.NewStart + match.Length]);
            long own = oldPrefix[match.OldStart + match.Length] - oldPrefix[match.OldStart];
            if (before > 0 && after > 0 && own <= before && own <= after) {
                matches.RemoveAt(index);
                index = Math.Max(0, index - 1);
                continue;
            }

            index++;
        }

        return matches;
    }

    private static long[] PrefixWeights<T>(IReadOnlyList<T> items, Func<T, int> weight) {
        long[] prefix = new long[items.Count + 1];
        for (int index = 0; index < items.Count; index++) prefix[index + 1] = prefix[index] + Math.Max(0, weight(items[index]));
        return prefix;
    }
}

/// <summary>Token comparer used by the text helpers; equality and hashing work on a normalized key.</summary>
internal sealed class OfficeTextTokenComparer : IEqualityComparer<OfficeTextToken> {
    private readonly bool _ignoreCase;
    private readonly bool _ignoreWhitespace;
    private readonly bool _lines;

    private OfficeTextTokenComparer(bool ignoreCase, bool ignoreWhitespace, bool lines) {
        _ignoreCase = ignoreCase;
        _ignoreWhitespace = ignoreWhitespace;
        _lines = lines;
    }

    internal static OfficeTextTokenComparer ForLines(OfficeDiffOptions? options) => new OfficeTextTokenComparer(options?.IgnoreCase == true, options?.IgnoreWhitespace == true, lines: true);

    internal static OfficeTextTokenComparer ForWords(OfficeDiffOptions? options) => new OfficeTextTokenComparer(options?.IgnoreCase == true, options?.IgnoreWhitespace == true, lines: false);

    internal static OfficeTextTokenComparer ForCharacters(OfficeDiffOptions? options) => new OfficeTextTokenComparer(options?.IgnoreCase == true, ignoreWhitespace: false, lines: false);

    internal string Key(OfficeTextToken token) {
        string text = token.Text ?? string.Empty;
        if (_ignoreWhitespace) {
            if (_lines) text = string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            else if (token.Kind == OfficeTextTokenKind.Whitespace) text = " ";
        }

        return _ignoreCase ? text.ToUpperInvariant() : text;
    }

    public bool Equals(OfficeTextToken x, OfficeTextToken y) => string.Equals(Key(x), Key(y), StringComparison.Ordinal);

    public int GetHashCode(OfficeTextToken obj) => StringComparer.Ordinal.GetHashCode(Key(obj));
}
