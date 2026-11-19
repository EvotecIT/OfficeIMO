using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

/// <summary>
/// Positions of each boundary's entries in an immutable, offset-ordered forced-break list. A long document folds
/// its next boundary once per page, so the fold reads that boundary's entries instead of every later break.
/// </summary>
internal static class HtmlRenderForcedBreakBoundaryIndex {
    // Short lists are scanned directly; an index costs more than it saves there.
    private const int IndexedCount = 16;
    private static readonly ConditionalWeakTable<IReadOnlyList<HtmlRenderForcedBreak>, Dictionary<object, List<int>>> Indexes = new();

    /// <summary>The ascending positions of the entries whose boundary is <paramref name="boundary"/>.</summary>
    internal static IReadOnlyList<int> Of(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, object boundary) {
        if (forcedBreaks.Count <= IndexedCount) return Scan(forcedBreaks, boundary);
        Dictionary<object, List<int>> index = Indexes.GetValue(forcedBreaks, Build);
        return index.TryGetValue(boundary, out List<int>? positions) ? positions : Array.Empty<int>();
    }

    /// <summary>The first position in <paramref name="positions"/> that is at least <paramref name="value"/>.</summary>
    internal static int FirstAtOrAfter(IReadOnlyList<int> positions, int value) {
        int low = 0;
        int high = positions.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (positions[middle] < value) low = middle + 1;
            else high = middle;
        }
        return low;
    }

    private static List<int> Scan(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, object boundary) {
        var positions = new List<int>();
        for (int index = 0; index < forcedBreaks.Count; index++) {
            if (ReferenceEquals(forcedBreaks[index].Boundary, boundary)) positions.Add(index);
        }
        return positions;
    }

    private static Dictionary<object, List<int>> Build(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks) {
        var index = new Dictionary<object, List<int>>(ReferenceComparer.Instance);
        for (int position = 0; position < forcedBreaks.Count; position++) {
            object boundary = forcedBreaks[position].Boundary;
            if (!index.TryGetValue(boundary, out List<int>? positions)) {
                positions = new List<int>();
                index.Add(boundary, positions);
            }
            positions.Add(position);
        }
        return index;
    }

    private sealed class ReferenceComparer : IEqualityComparer<object> {
        internal static readonly ReferenceComparer Instance = new();

        public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);

        public int GetHashCode(object obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
