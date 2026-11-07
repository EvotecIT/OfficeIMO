using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Excel;

internal static partial class ExcelRangeVisualSnapshotBuilder {
    /// <summary>Updates border snapshots for fully included merges without changing text, formatting or viewport geometry.</summary>
    private static void ApplyMergedCellBorders(List<ExcelVisualCell> cells, IReadOnlyList<ExcelMergedRangeInfo> merges,
        ISet<string> boundedMergeGeometry, int firstRow, int firstColumn, int lastRow, int lastColumn,
        string sheetName, IList<OfficeImageExportDiagnostic> diagnostics, CancellationToken cancellationToken) {
        if (merges.Count == 0) return;
        var indexes = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int index = 0; index < cells.Count; index++) indexes[Key(cells[index].Row, cells[index].Column)] = index;
        foreach (ExcelMergedRangeInfo merge in merges) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!boundedMergeGeometry.Contains(MergeGeometryKey(merge))
                || merge.StartRow < firstRow || merge.StartColumn < firstColumn
                || merge.EndRow > lastRow || merge.EndColumn > lastColumn
                || !indexes.TryGetValue(Key(merge.StartRow, merge.StartColumn), out int index)) continue;
            ExcelVisualCell origin = cells[index];
            ExcelCellBorderInfo? border = ExcelMergedCellBorderResolver.Resolve(
                merge.StartRow, merge.StartColumn, merge.EndRow, merge.EndColumn, origin.Style.Border,
                (row, column) => indexes.TryGetValue(Key(row, column), out int cellIndex) ? cells[cellIndex].Style.Border : null,
                sheetName + "!" + merge.A1Range, diagnostics.Add);
            if (ReferenceEquals(border, origin.Style.Border)) continue;
            cells[index] = new ExcelVisualCell(origin.Row, origin.Column, origin.X, origin.Y, origin.Width, origin.Height,
                origin.Text, origin.Style.CopyWithBorder(border), origin.CoveredByMerge, origin.Hyperlink,
                origin.RichTextRuns, origin.ValueKind);
        }
    }
}
