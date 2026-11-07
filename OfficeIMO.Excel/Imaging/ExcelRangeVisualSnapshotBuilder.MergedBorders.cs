using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Excel;

internal static partial class ExcelRangeVisualSnapshotBuilder {
    /// <summary>Updates source perimeter edges without moving merged geometry or creating borders at viewport cuts.</summary>
    private static void ApplyMergedCellBorders(List<ExcelVisualCell> cells, IReadOnlyList<ExcelMergedRangeInfo> merges,
        ISet<string> boundedMergeGeometry, int firstRow, int firstColumn, int lastRow, int lastColumn,
        ExcelSheet sheet, IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
        IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
        IList<OfficeImageExportDiagnostic> diagnostics, CancellationToken cancellationToken) {
        if (merges.Count == 0) return;
        var indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < cells.Count; index++) indexes[Key(cells[index].Row, cells[index].Column)] = index;
        var rowCoordinates = new SortedSet<int>(rowsByIndex.Keys);
        var columnCoordinates = new SortedSet<int>(columnsByIndex.Keys);
        foreach (ExcelMergedRangeInfo merge in merges) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!boundedMergeGeometry.Contains(MergeGeometryKey(merge))
                || !indexes.TryGetValue(Key(merge.StartRow, merge.StartColumn), out int index)) continue;
            ExcelVisualCell origin = cells[index];
            bool topVisible = merge.StartRow >= firstRow;
            bool rightVisible = merge.EndColumn <= lastColumn;
            bool bottomVisible = merge.EndRow <= lastRow;
            bool leftVisible = merge.StartColumn >= firstColumn;
            ExcelCellBorderInfo? border = ExcelMergedCellBorderResolver.Resolve(
                merge.StartRow, merge.StartColumn, merge.EndRow, merge.EndColumn, VisibleEdges(origin.Style.Border),
                BorderAt, sheet.Name + "!" + merge.A1Range, diagnostics.Add,
                rowCoordinates.GetViewBetween(merge.StartRow, merge.EndRow),
                columnCoordinates.GetViewBetween(merge.StartColumn, merge.EndColumn));
            if (ReferenceEquals(border, origin.Style.Border)) continue;
            cells[index] = new ExcelVisualCell(origin.Row, origin.Column, origin.X, origin.Y, origin.Width, origin.Height,
                origin.Text, origin.Style.CopyWithBorder(border), origin.CoveredByMerge, origin.Hyperlink,
                origin.RichTextRuns, origin.ValueKind);

            ExcelCellBorderInfo? BorderAt(int row, int column) {
                if (row < firstRow || row > lastRow || column < firstColumn || column > lastColumn) return null;
                // Hidden source boundary cells still carry perimeter metadata, even without a visual cell.
                return VisibleEdges(indexes.TryGetValue(Key(row, column), out int cellIndex)
                    ? cells[cellIndex].Style.Border : sheet.GetCellStyle(row, column).Border);
            }

            ExcelCellBorderInfo? VisibleEdges(ExcelCellBorderInfo? source) {
                if (source == null || (topVisible && rightVisible && bottomVisible && leftVisible)) return source;
                return new ExcelCellBorderInfo(leftVisible ? source.Left : null, rightVisible ? source.Right : null,
                    topVisible ? source.Top : null, bottomVisible ? source.Bottom : null, source.Diagonal,
                    source.DiagonalUp, source.DiagonalDown);
            }
        }
    }
}
