namespace OfficeIMO.Excel {
    /// <summary>Indexes rectangular cell ranges without expanding their covered worksheet area.</summary>
    internal sealed class ExcelCellRangeIndex<T> {
        private readonly Entry[] _entries;
        private readonly Node? _root;

        internal ExcelCellRangeIndex(IEnumerable<(int Top, int Left, int Bottom, int Right, T Value)> ranges) {
            _entries = ranges.Select(range => new Entry(range.Top, range.Left, range.Bottom, range.Right, range.Value)).ToArray();
            _root = Build(0, _entries.Length);
        }

        internal IEnumerable<T> At(int row, int column) => Intersect(row, column, row, column);

        internal IEnumerable<T> Intersect(int top, int left, int bottom, int right) => Visit(_root, top, left, bottom, right);

        private IEnumerable<T> Visit(Node? node, int top, int left, int bottom, int right) {
            if (node == null || node.Top > bottom || node.Bottom < top || node.Left > right || node.Right < left) yield break;
            if (node.First != null) {
                foreach (T value in Visit(node.First, top, left, bottom, right)) yield return value;
                foreach (T value in Visit(node.Second, top, left, bottom, right)) yield return value;
            } else {
                for (int index = node.Start; index < node.End; index++) {
                    Entry entry = _entries[index];
                    if (entry.Top <= bottom && entry.Bottom >= top && entry.Left <= right && entry.Right >= left) yield return entry.Value;
                }
            }
        }

        private Node? Build(int start, int count) {
            if (count == 0) return null;
            int end = start + count;
            int top = int.MaxValue, left = int.MaxValue, bottom = 0, right = 0;
            int minRow = int.MaxValue, maxRow = 0, minColumn = int.MaxValue, maxColumn = 0;
            for (int index = start; index < end; index++) {
                Entry entry = _entries[index];
                top = Math.Min(top, entry.Top); bottom = Math.Max(bottom, entry.Bottom);
                left = Math.Min(left, entry.Left); right = Math.Max(right, entry.Right);
                int row = entry.Top + entry.Bottom, column = entry.Left + entry.Right;
                minRow = Math.Min(minRow, row); maxRow = Math.Max(maxRow, row);
                minColumn = Math.Min(minColumn, column); maxColumn = Math.Max(maxColumn, column);
            }
            Node node = new Node(top, left, bottom, right, start, end);
            if (count > 8) {
                bool splitRows = maxRow - minRow >= maxColumn - minColumn;
                Array.Sort(_entries, start, count, Comparer<Entry>.Create((a, b) => splitRows
                    ? (a.Top + a.Bottom).CompareTo(b.Top + b.Bottom)
                    : (a.Left + a.Right).CompareTo(b.Left + b.Right)));
                int firstCount = count / 2;
                node.First = Build(start, firstCount); node.Second = Build(start + firstCount, count - firstCount);
            }
            return node;
        }

        private sealed record Entry(int Top, int Left, int Bottom, int Right, T Value);
        private sealed record Node(int Top, int Left, int Bottom, int Right, int Start, int End) {
            internal Node? First { get; set; }
            internal Node? Second { get; set; }
        }
    }
}
