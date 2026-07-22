using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    private static string ExpandUsedRangeForMerges(
        ExcelSheet sheet,
        string usedRange,
        IReadOnlyList<ExcelMergedRangeSnapshot> mergedRanges,
        int maximumRows,
        int maximumColumns,
        int maximumCells) {
        ParseUsedRange(usedRange, out int firstRow, out int firstColumn, out int rowCount, out int columnCount);
        int lastRow = firstRow + Math.Max(1, rowCount) - 1;
        int lastColumn = firstColumn + Math.Max(1, columnCount) - 1;

        int scanColumns = Math.Min(Math.Min(columnCount, maximumColumns), maximumCells);
        int scanRows = scanColumns == 0
            ? 0
            : Math.Min(Math.Min(rowCount, maximumRows), Math.Max(1, maximumCells / scanColumns));
        if (mergedRanges.Count > 0 &&
            !SheetHasUsedCells(sheet, firstRow, firstColumn, scanRows, scanColumns)) {
            ExcelMergedRangeSnapshot firstMerge = mergedRanges[0];
            firstRow = firstMerge.StartRow;
            firstColumn = firstMerge.StartColumn;
            lastRow = firstMerge.EndRow;
            lastColumn = firstMerge.EndColumn;
        }

        foreach (ExcelMergedRangeSnapshot merge in mergedRanges) {
            firstRow = Math.Min(firstRow, merge.StartRow);
            firstColumn = Math.Min(firstColumn, merge.StartColumn);
            lastRow = Math.Max(lastRow, merge.EndRow);
            lastColumn = Math.Max(lastColumn, merge.EndColumn);
        }

        lastRow = Math.Min(lastRow, AddBounded(firstRow, maximumRows - 1));
        lastColumn = Math.Min(lastColumn, AddBounded(firstColumn, maximumColumns - 1));

        return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
    }

    private static int AddBounded(int value, int offset) =>
        value > int.MaxValue - offset ? int.MaxValue : value + offset;

    private static ExcelMergeExportMap BuildMergeExportMap(
        IReadOnlyList<ExcelMergedRangeSnapshot> mergedRanges,
        int firstRow,
        int firstColumn,
        int rowCount,
        int columnCount) {
        int lastRow = firstRow + rowCount - 1;
        int lastColumn = firstColumn + columnCount - 1;
        var candidates = new List<ExcelMergeExportRange>();

        foreach (ExcelMergedRangeSnapshot merge in mergedRanges) {
            int startRow = Math.Max(firstRow, merge.StartRow);
            int startColumn = Math.Max(firstColumn, merge.StartColumn);
            int endRow = Math.Min(lastRow, merge.EndRow);
            int endColumn = Math.Min(lastColumn, merge.EndColumn);
            if (startRow > endRow || startColumn > endColumn) {
                continue;
            }

            candidates.Add(new ExcelMergeExportRange(startRow, startColumn, endRow, endColumn));
        }

        candidates.Sort(static (left, right) => {
            int row = left.StartRow.CompareTo(right.StartRow);
            return row != 0 ? row : left.StartColumn.CompareTo(right.StartColumn);
        });
        var accepted = new List<ExcelMergeExportRange>(candidates.Count);
        var active = new List<ExcelMergeExportRange>();
        foreach (ExcelMergeExportRange candidate in candidates) {
            active.RemoveAll(range => range.EndRow < candidate.StartRow);
            if (active.Any(range => range.Intersects(candidate))) continue;
            accepted.Add(candidate);
            active.Add(candidate);
        }
        return new ExcelMergeExportMap(accepted);
    }

    private static void AppendMergeAttributes(StringBuilder body, ExcelMergeExportRange merge) {
        if (merge.RowSpan > 1) {
            body.Append(" rowspan=\"")
                .Append(merge.RowSpan.ToString(CultureInfo.InvariantCulture))
                .Append('"');
        }

        if (merge.ColumnSpan > 1) {
            body.Append(" colspan=\"")
                .Append(merge.ColumnSpan.ToString(CultureInfo.InvariantCulture))
                .Append('"');
        }

        body.Append(" data-officeimo-merge=\"")
            .Append(OfficeHtmlText.EscapeAttribute(merge.A1Range))
            .Append('"');
    }

    private sealed class ExcelMergeExportMap {
        private readonly IReadOnlyList<ExcelMergeExportRange> _ranges;

        internal ExcelMergeExportMap(IReadOnlyList<ExcelMergeExportRange> ranges) {
            _ranges = ranges;
        }

        internal int Count => _ranges.Count;

        internal ExcelMergeExportRowCursor CreateRowCursor() => new(_ranges);
    }

    private sealed class ExcelMergeExportRowCursor {
        private readonly IReadOnlyList<ExcelMergeExportRange> _ranges;
        private readonly List<ExcelMergeExportRange> _active = new();
        private int _nextRange;
        private int _row;

        internal ExcelMergeExportRowCursor(IReadOnlyList<ExcelMergeExportRange> ranges) {
            _ranges = ranges;
        }

        internal void MoveToRow(int row) {
            if (_row != 0 && row < _row) throw new InvalidOperationException("Merge row cursors must move forward.");
            _row = row;
            _active.RemoveAll(range => range.EndRow < row);
            while (_nextRange < _ranges.Count && _ranges[_nextRange].StartRow <= row) {
                ExcelMergeExportRange range = _ranges[_nextRange++];
                if (range.EndRow >= row) _active.Add(range);
            }
            _active.Sort(static (left, right) => left.StartColumn.CompareTo(right.StartColumn));
        }

        internal bool IsCoveredCell(int column) {
            if (!TryFind(column, out ExcelMergeExportRange range)) return false;
            return _row != range.StartRow || column != range.StartColumn;
        }

        internal bool TryGetOrigin(int column, out ExcelMergeExportRange merge) {
            if (TryFind(column, out merge) && _row == merge.StartRow && column == merge.StartColumn) return true;
            merge = default;
            return false;
        }

        private bool TryFind(int column, out ExcelMergeExportRange range) {
            int low = 0;
            int high = _active.Count - 1;
            while (low <= high) {
                int middle = low + ((high - low) / 2);
                ExcelMergeExportRange candidate = _active[middle];
                if (column < candidate.StartColumn) high = middle - 1;
                else if (column > candidate.EndColumn) low = middle + 1;
                else {
                    range = candidate;
                    return true;
                }
            }
            range = default;
            return false;
        }
    }

    private readonly struct ExcelMergeExportRange {
        internal ExcelMergeExportRange(int startRow, int startColumn, int endRow, int endColumn) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
        }

        internal int StartRow { get; }

        internal int StartColumn { get; }

        internal int EndRow { get; }

        internal int EndColumn { get; }

        internal int RowSpan => EndRow - StartRow + 1;

        internal int ColumnSpan => EndColumn - StartColumn + 1;

        internal string A1Range => A1.CellReference(StartRow, StartColumn) + ":" + A1.CellReference(EndRow, EndColumn);

        internal bool Intersects(ExcelMergeExportRange other) =>
            StartRow <= other.EndRow && EndRow >= other.StartRow
            && StartColumn <= other.EndColumn && EndColumn >= other.StartColumn;
    }

}
