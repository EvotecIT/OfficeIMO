namespace OfficeIMO.IWork.Internal;

internal static class IWorkMergeRangeValidator {
    internal static bool HasOverlaps(IReadOnlyList<IWorkTableMergeRange> ranges, int columnCount) {
        if (ranges.Count < 2 || columnCount == 0) return false;
        IWorkTableMergeRange[] ordered = ranges.OrderBy(range => range.FirstRow)
            .ThenBy(range => range.FirstColumn).ToArray();
        var active = new SortedSet<ActiveMerge>(ActiveMergeComparer.Instance);
        var columns = new RangeMaximumCounter(columnCount);
        int identifier = 0;
        foreach (IWorkTableMergeRange range in ordered) {
            while (active.Count > 0 && active.Min!.LastRow < range.FirstRow) {
                ActiveMerge expired = active.Min!;
                active.Remove(expired);
                columns.Add(expired.FirstColumn - 1, expired.LastColumn - 1, -1);
            }
            if (columns.Maximum(range.FirstColumn - 1, range.LastColumn - 1) > 0) return true;
            columns.Add(range.FirstColumn - 1, range.LastColumn - 1, 1);
            active.Add(new ActiveMerge(range.LastRow, identifier++, range.FirstColumn, range.LastColumn));
        }
        return false;
    }

    private sealed class ActiveMerge {
        internal ActiveMerge(int lastRow, int identifier, int firstColumn, int lastColumn) {
            LastRow = lastRow;
            Identifier = identifier;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
        }
        internal int LastRow { get; }
        internal int Identifier { get; }
        internal int FirstColumn { get; }
        internal int LastColumn { get; }
    }

    private sealed class ActiveMergeComparer : IComparer<ActiveMerge> {
        internal static ActiveMergeComparer Instance { get; } = new();
        public int Compare(ActiveMerge? left, ActiveMerge? right) {
            if (ReferenceEquals(left, right)) return 0;
            if (left == null) return -1;
            if (right == null) return 1;
            int byRow = left.LastRow.CompareTo(right.LastRow);
            return byRow != 0 ? byRow : left.Identifier.CompareTo(right.Identifier);
        }
    }

    private sealed class RangeMaximumCounter {
        private readonly int[] _maximum;
        private readonly int[] _lazy;
        private readonly int _length;

        internal RangeMaximumCounter(int length) {
            _length = length;
            _maximum = new int[checked(length * 4)];
            _lazy = new int[_maximum.Length];
        }

        internal void Add(int first, int last, int value) => Add(1, 0, _length - 1, first, last, value);

        internal int Maximum(int first, int last) => Maximum(1, 0, _length - 1, first, last);

        private void Add(int node, int nodeFirst, int nodeLast, int first, int last, int value) {
            if (first <= nodeFirst && nodeLast <= last) {
                _maximum[node] += value;
                _lazy[node] += value;
                return;
            }
            int middle = nodeFirst + (nodeLast - nodeFirst) / 2;
            if (first <= middle) Add(node * 2, nodeFirst, middle, first, last, value);
            if (last > middle) Add(node * 2 + 1, middle + 1, nodeLast, first, last, value);
            _maximum[node] = _lazy[node] + Math.Max(_maximum[node * 2], _maximum[node * 2 + 1]);
        }

        private int Maximum(int node, int nodeFirst, int nodeLast, int first, int last) {
            if (first <= nodeFirst && nodeLast <= last) return _maximum[node];
            int middle = nodeFirst + (nodeLast - nodeFirst) / 2;
            int result = 0;
            if (first <= middle) result = Maximum(node * 2, nodeFirst, middle, first, last);
            if (last > middle) result = Math.Max(result, Maximum(node * 2 + 1, middle + 1, nodeLast, first, last));
            return _lazy[node] + result;
        }
    }
}
