using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    private static string ExpandUsedRangeForMerges(ExcelSheet sheet, string usedRange, IReadOnlyList<ExcelMergedRangeSnapshot> mergedRanges) {
        ParseUsedRange(usedRange, out int firstRow, out int firstColumn, out int rowCount, out int columnCount);
        int lastRow = firstRow + Math.Max(1, rowCount) - 1;
        int lastColumn = firstColumn + Math.Max(1, columnCount) - 1;

        if (mergedRanges.Count > 0 && !SheetHasUsedCells(sheet, firstRow, firstColumn, rowCount, columnCount)) {
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

        return A1.CellReference(firstRow, firstColumn) + ":" + A1.CellReference(lastRow, lastColumn);
    }

    private static ExcelMergeExportMap BuildMergeExportMap(
        IReadOnlyList<ExcelMergedRangeSnapshot> mergedRanges,
        int firstRow,
        int firstColumn,
        int rowCount,
        int columnCount) {
        int lastRow = firstRow + rowCount - 1;
        int lastColumn = firstColumn + columnCount - 1;
        var map = new ExcelMergeExportMap();

        foreach (ExcelMergedRangeSnapshot merge in mergedRanges) {
            int startRow = Math.Max(firstRow, merge.StartRow);
            int startColumn = Math.Max(firstColumn, merge.StartColumn);
            int endRow = Math.Min(lastRow, merge.EndRow);
            int endColumn = Math.Min(lastColumn, merge.EndColumn);
            if (startRow > endRow || startColumn > endColumn) {
                continue;
            }

            map.TryAdd(new ExcelMergeExportRange(startRow, startColumn, endRow, endColumn));
        }

        return map;
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
        private readonly Dictionary<long, ExcelMergeExportRange> _origins = new();
        private readonly HashSet<long> _coveredCells = new();

        internal int Count => _origins.Count;

        internal bool TryAdd(ExcelMergeExportRange merge) {
            for (int row = merge.StartRow; row <= merge.EndRow; row++) {
                for (int column = merge.StartColumn; column <= merge.EndColumn; column++) {
                    long key = GetCellKey(row, column);
                    if (_coveredCells.Contains(key) || _origins.ContainsKey(key)) {
                        return false;
                    }
                }
            }

            _origins.Add(GetCellKey(merge.StartRow, merge.StartColumn), merge);
            for (int row = merge.StartRow; row <= merge.EndRow; row++) {
                for (int column = merge.StartColumn; column <= merge.EndColumn; column++) {
                    if (row != merge.StartRow || column != merge.StartColumn) {
                        _coveredCells.Add(GetCellKey(row, column));
                    }
                }
            }

            return true;
        }

        internal bool IsCoveredCell(int row, int column) => _coveredCells.Contains(GetCellKey(row, column));

        internal bool TryGetOrigin(int row, int column, out ExcelMergeExportRange merge) =>
            _origins.TryGetValue(GetCellKey(row, column), out merge);
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
    }

    private static long GetCellKey(int row, int column) => ((long)row << 32) | (uint)column;
}
