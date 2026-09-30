using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Identifies every intersecting rectangle without enumerating overlap pairs or dense table cells.</summary>
internal sealed class IWorkMergeOverlapIndex {
    private readonly IReadOnlyList<IWorkTableMergeRange> _ranges;
    private readonly TopTwo[] _cover;
    private readonly TopTwo[] _maximum;

    private IWorkMergeOverlapIndex(IReadOnlyList<IWorkTableMergeRange> ranges, int columnCount) {
        _ranges = ranges;
        _cover = new TopTwo[checked(columnCount * 4)];
        _maximum = new TopTwo[_cover.Length];
    }

    internal static IReadOnlyList<int> FindOverlapIndexes(IReadOnlyList<IWorkTableMergeRange> ranges,
        int columnCount, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (ranges.Count < 2 || columnCount == 0) return Array.Empty<int>();
        int[] byStart = Enumerable.Range(0, ranges.Count).OrderBy(index => ranges[index].FirstRow).ToArray();
        int[] byEnd = Enumerable.Range(0, ranges.Count).OrderBy(index => ranges[index].LastRow).ToArray();
        var tree = new IWorkMergeOverlapIndex(ranges, columnCount);
        var overlaps = new bool[ranges.Count];
        int added = 0;
        // At each query, all rectangles that start no later than its last row are indexed.
        // A column intersection with another rectangle ending at/after its first row is an overlap.
        foreach (int index in byEnd) {
            cancellationToken.ThrowIfCancellationRequested();
            IWorkTableMergeRange range = ranges[index];
            while (added < byStart.Length && ranges[byStart[added]].FirstRow <= range.LastRow) {
                cancellationToken.ThrowIfCancellationRequested();
                int candidate = byStart[added++];
                tree.Add(1, 0, columnCount - 1, ranges[candidate].FirstColumn - 1,
                    ranges[candidate].LastColumn - 1, candidate + 1);
            }
            TopTwo maximum = tree.Query(1, 0, columnCount - 1, range.FirstColumn - 1, range.LastColumn - 1);
            int other = maximum.First == index + 1 ? maximum.Second : maximum.First;
            overlaps[index] = other != 0 && ranges[other - 1].LastRow >= range.FirstRow;
        }
        return Array.AsReadOnly(Enumerable.Range(0, ranges.Count).Where(index => overlaps[index]).ToArray());
    }

    private void Add(int node, int first, int last, int start, int end, int encodedIndex) {
        if (start <= first && last <= end) _cover[node] = Include(_cover[node], encodedIndex);
        else {
            int middle = first + (last - first) / 2;
            if (start <= middle) Add(node * 2, first, middle, start, end, encodedIndex);
            if (end > middle) Add(node * 2 + 1, middle + 1, last, start, end, encodedIndex);
        }
        _maximum[node] = first == last ? _cover[node]
            : Merge(_cover[node], Merge(_maximum[node * 2], _maximum[node * 2 + 1]));
    }

    private TopTwo Query(int node, int first, int last, int start, int end) {
        if (start <= first && last <= end) return _maximum[node];
        int middle = first + (last - first) / 2;
        TopTwo result = _cover[node];
        if (start <= middle) result = Merge(result, Query(node * 2, first, middle, start, end));
        if (end > middle) result = Merge(result, Query(node * 2 + 1, middle + 1, last, start, end));
        return result;
    }

    private TopTwo Merge(TopTwo left, TopTwo right) => Include(Include(left, right.First), right.Second);

    private TopTwo Include(TopTwo pair, int encodedIndex) {
        if (encodedIndex == 0 || encodedIndex == pair.First || encodedIndex == pair.Second) return pair;
        if (pair.First == 0 || _ranges[encodedIndex - 1].LastRow > _ranges[pair.First - 1].LastRow)
            return new TopTwo(encodedIndex, pair.First);
        if (pair.Second == 0 || _ranges[encodedIndex - 1].LastRow > _ranges[pair.Second - 1].LastRow)
            return new TopTwo(pair.First, encodedIndex);
        return pair;
    }

    // Two distinct identities are needed because each query must exclude its own rectangle.
    // Identity+1 reserves zero for an empty slot, including in freshly allocated arrays.
    private readonly struct TopTwo(int first, int second) {
        internal int First { get; } = first;
        internal int Second { get; } = second;
    }
}
