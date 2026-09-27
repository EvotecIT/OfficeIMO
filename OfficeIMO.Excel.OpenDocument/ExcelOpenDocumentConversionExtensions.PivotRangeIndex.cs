namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    // A balanced bounding-box tree keeps disjoint range queries independent of pivot count.
    // Query work also has a shared budget; pathological overlapping boxes fail closed.
    private sealed class PivotRangeIndex {
        internal sealed class Entry {
            internal long FirstRow, LastRow, FirstColumn, LastColumn;
            internal OfficeIMO.OpenDocument.OdsDataPilotTable? Owner;
            internal Entry(long firstRow, long lastRow, long firstColumn, long lastColumn,
                OfficeIMO.OpenDocument.OdsDataPilotTable? owner = null) {
                FirstRow = firstRow; LastRow = lastRow; FirstColumn = firstColumn; LastColumn = lastColumn; Owner = owner;
            }
        }
        private sealed class Node {
            internal long FirstRow, LastRow, FirstColumn, LastColumn;
            internal Node? Left, Right;
            internal Entry[]? Entries;
        }
        private readonly Node? _root;
        private readonly HashSet<OfficeIMO.OpenDocument.OdsDataPilotTable> _enabled = new();
        internal PivotRangeIndex(IEnumerable<Entry> entries) {
            Entry[] items = entries.ToArray();
            _root = items.Length == 0 ? null : Build(items, 0, items.Length);
        }
        internal void Enable(OfficeIMO.OpenDocument.OdsDataPilotTable owner) => _enabled.Add(owner);
        private static Node Build(Entry[] items, int start, int count) {
            var node = new Node { FirstRow = long.MaxValue, FirstColumn = long.MaxValue };
            for (int i = start; i < start + count; i++) {
                node.FirstRow = Math.Min(node.FirstRow, items[i].FirstRow);
                node.LastRow = Math.Max(node.LastRow, items[i].LastRow);
                node.FirstColumn = Math.Min(node.FirstColumn, items[i].FirstColumn);
                node.LastColumn = Math.Max(node.LastColumn, items[i].LastColumn);
            }
            if (count <= 8) { node.Entries = items.Skip(start).Take(count).ToArray(); return node; }
            bool rows = node.LastRow - node.FirstRow >= node.LastColumn - node.FirstColumn;
            Array.Sort(items, start, count, Comparer<Entry>.Create((a, b) => rows
                ? a.FirstRow.CompareTo(b.FirstRow) : a.FirstColumn.CompareTo(b.FirstColumn)));
            int half = count / 2;
            node.Left = Build(items, start, half);
            node.Right = Build(items, start + half, count - half);
            return node;
        }
        internal bool Intersects(long firstRow, long lastRow, long firstColumn, long lastColumn, ref long budget) =>
            Intersects(_root, firstRow, lastRow, firstColumn, lastColumn, ref budget);
        private bool Intersects(Node? node, long firstRow, long lastRow, long firstColumn, long lastColumn, ref long budget) {
            if (node == null) return false;
            if (budget <= 0) return true;
            budget--;
            if (!PivotFootprintsOverlap(node.FirstRow, node.LastRow, node.FirstColumn, node.LastColumn,
                    firstRow, lastRow, firstColumn, lastColumn)) return false;
            if (node.Entries != null) {
                foreach (Entry item in node.Entries) {
                    if (budget <= 0) return true;
                    budget--;
                    if ((item.Owner == null || _enabled.Contains(item.Owner))
                        && PivotFootprintsOverlap(item.FirstRow, item.LastRow, item.FirstColumn, item.LastColumn,
                            firstRow, lastRow, firstColumn, lastColumn)) return true;
                }
                return false;
            }
            return Intersects(node.Left, firstRow, lastRow, firstColumn, lastColumn, ref budget)
                || Intersects(node.Right, firstRow, lastRow, firstColumn, lastColumn, ref budget);
        }
    }
}
