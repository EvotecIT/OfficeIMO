using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        // Project the selected page grid once. Continuation fragments retain the
        // original merge's value, text formatting and full text-box geometry.
        private static SheetExportData SelectPageCells(WorksheetPdfExportPlan plan, TableChunk chunk, ExcelToPdfOptions options) {
            SheetExportData source = plan.ExportData;
            IReadOnlyList<int> rows = chunk.RowIndexes;
            IReadOnlyList<int> columns = chunk.ColumnIndexes;
            object?[,] values = SelectCells(source.Values, rows, columns)!;
            ExcelCellStyleSnapshot?[,]? styles = SelectCells(source.Styles, rows, columns);
            ExcelHyperlinkSnapshot?[,]? hyperlinks = SelectCells(source.Hyperlinks, rows, columns);
            string?[,]? references = SelectCells(source.CellReferences, rows, columns);
            Dictionary<(int Row, int Column), string>? conditionalFills = source.ConditionalFills == null ? null : SelectCellVisuals(source.ConditionalFills.FillColors, rows, columns);
            Dictionary<(int Row, int Column), ConditionalDataBarCell>? dataBars = source.ConditionalFills == null ? null : SelectCellVisuals(source.ConditionalFills.DataBars, rows, columns);
            Dictionary<(int Row, int Column), ConditionalIconCell>? icons = source.ConditionalFills == null ? null : SelectCellVisuals(source.ConditionalFills.Icons, rows, columns);
            MergeLayoutData? merged = source.MergedCells == null ? null : new MergeLayoutData(rows.Count, columns.Count);
            if (merged != null) {
                for (int localRow = 0; localRow < rows.Count; localRow++) {
                    int row = rows[localRow];
                    for (int localColumn = 0; localColumn < columns.Count; localColumn++) {
                        int column = columns[localColumn];
                        MergeRegion? region = source.MergedCells!.GetRegion(row, column);
                        if (region == null) continue;
                        if (localRow > 0 && localRow != chunk.HeaderRowCount && rows[localRow - 1] == row - 1 && row > region.Row) continue;
                        if (localColumn > 0 && columns[localColumn - 1] == column - 1 && column > region.Column) continue;
                        int rowSpan = CountMergeFragmentAxis(rows, localRow, region.Row + region.Span.RowSpan,
                            localRow < chunk.HeaderRowCount ? chunk.HeaderRowCount : rows.Count);
                        int columnSpan = CountMergeFragmentAxis(columns, localColumn, region.Column + region.Span.ColumnSpan, columns.Count);
                        PdfCore.PdfTableCellViewport? viewport = CreateMergeViewport(plan, region, row, column, rowSpan, columnSpan);
                        merged.SetSpan(localRow, localColumn, rowSpan, columnSpan, viewport);
                        values[localRow, localColumn] = source.Values[region.Row, region.Column];
                        if (source.ConditionalFills != null) {
                            CopyMergeAnchorVisual(source.ConditionalFills.FillColors, conditionalFills!, region, localRow, localColumn);
                            CopyMergeAnchorVisual(source.ConditionalFills.DataBars, dataBars!, region, localRow, localColumn);
                            CopyMergeAnchorVisual(source.ConditionalFills.Icons, icons!, region, localRow, localColumn);
                        }
                        if (hyperlinks != null) hyperlinks[localRow, localColumn] = source.Hyperlinks![region.Row, region.Column];
                        if (styles != null || viewport != null && options.UseWorksheetCellStyles) {
                            styles ??= new ExcelCellStyleSnapshot?[rows.Count, columns.Count];
                            ExcelCellStyleSnapshot originalStyle = source.Styles?[region.Row, region.Column] ?? new ExcelCellStyleSnapshot();
                            ExcelCellStyleSnapshot fragmentStyle = originalStyle.CopyWithBorder(
                                CreateMergeFragmentBorder(source.Styles, region, row, column, rowSpan, columnSpan));
                            if (viewport != null && string.IsNullOrWhiteSpace(fragmentStyle.VerticalAlignment)) fragmentStyle.VerticalAlignment = "bottom";
                            styles[localRow, localColumn] = fragmentStyle;
                        }
                    }
                }
            }
            ColumnLayoutData? widths = source.ColumnWidths == null ? null : new ColumnLayoutData(
                columns.Select(column => source.ColumnWidths.WidthWeights[column]).ToList(),
                source.ColumnWidths.ApproximateWidthPoints * columns.Sum(column => source.ColumnWidths.WidthWeights[column])
                    / Math.Max(1D, source.ColumnWidths.WidthWeights.Sum()));
            RowLayoutData? heights = source.RowHeights == null ? null : new RowLayoutData(
                rows.Select(row => source.RowHeights.MinHeights[row]).ToList());
            ConditionalFillData? conditional = source.ConditionalFills == null ? null : new ConditionalFillData(
                conditionalFills!, dataBars!, icons!);
            return new SheetExportData(values, styles, hyperlinks, references, merged, widths, heights,
                chunk.HeaderRowCount, source.FirstBodyRowNumber, source.StructuredTables, conditional);
        }

        private static int CountMergeFragmentAxis(IReadOnlyList<int> indexes, int start, int mergeEnd, int pageEnd) {
            int count = 1;
            while (start + count < pageEnd && indexes[start + count] == indexes[start] + count && indexes[start + count] < mergeEnd) count++;
            return count;
        }

        private static PdfCore.PdfTableCellViewport? CreateMergeViewport(
            WorksheetPdfExportPlan plan, MergeRegion region, int row, int column, int rowSpan, int columnSpan) {
            if (row == region.Row && column == region.Column && rowSpan == region.Span.RowSpan && columnSpan == region.Span.ColumnSpan) return null;
            double width = Enumerable.Range(column, columnSpan).Sum(index => GetExportedColumnWidthPoints(plan, index));
            double height = Enumerable.Range(row, rowSpan).Sum(index => GetExportedRowHeightPoints(plan, index));
            double offsetX = Enumerable.Range(region.Column, column - region.Column).Sum(index => GetExportedColumnWidthPoints(plan, index));
            double offsetY = Enumerable.Range(region.Row, row - region.Row).Sum(index => GetExportedRowHeightPoints(plan, index));
            double fullWidth = Enumerable.Range(region.Column, region.Span.ColumnSpan).Sum(index => GetExportedColumnWidthPoints(plan, index));
            double fullHeight = Enumerable.Range(region.Row, region.Span.RowSpan).Sum(index => GetExportedRowHeightPoints(plan, index));
            // Accumulating a fragment separately can round a few ulps above the full
            // sum. Keep the viewport contained without changing meaningful geometry.
            return new PdfCore.PdfTableCellViewport(Math.Max(fullWidth, offsetX + width), Math.Max(fullHeight, offsetY + height), width, height, offsetX, offsetY);
        }

        private static ExcelCellBorderSnapshot? CreateMergeFragmentBorder(
            ExcelCellStyleSnapshot?[,]? styles, MergeRegion region, int row, int column, int rowSpan, int columnSpan) {
            if (styles == null) return null;
            ExcelCellBorderSnapshot? original = styles[region.Row, region.Column]?.Border;
            int lastRow = region.Row + region.Span.RowSpan - 1;
            int lastColumn = region.Column + region.Span.ColumnSpan - 1;
            ExcelCellBorderSnapshot? right = styles[region.Row, lastColumn]?.Border;
            ExcelCellBorderSnapshot? bottom = styles[lastRow, region.Column]?.Border;
            bool includesAnchor = row == region.Row && column == region.Column;
            return new ExcelCellBorderSnapshot(
                left: column == region.Column ? original?.Left : null,
                right: column + columnSpan - 1 == lastColumn ? right?.Right ?? original?.Right : null,
                top: row == region.Row ? original?.Top : null,
                bottom: row + rowSpan - 1 == lastRow ? bottom?.Bottom ?? original?.Bottom : null,
                diagonal: includesAnchor ? original?.Diagonal : null,
                diagonalUp: includesAnchor && original?.DiagonalUp == true,
                diagonalDown: includesAnchor && original?.DiagonalDown == true);
        }

        private static T[,]? SelectCells<T>(T[,]? source, IReadOnlyList<int> rows, IReadOnlyList<int> columns) {
            if (source == null) return null;
            var result = new T[rows.Count, columns.Count];
            for (int row = 0; row < rows.Count; row++)
                for (int column = 0; column < columns.Count; column++) result[row, column] = source[rows[row], columns[column]];
            return result;
        }

        private static void CopyMergeAnchorVisual<T>(IReadOnlyDictionary<(int Row, int Column), T> source,
            Dictionary<(int Row, int Column), T> target, MergeRegion region, int row, int column) {
            if (source.TryGetValue((region.Row, region.Column), out T? value)) target[(row, column)] = value;
            else target.Remove((row, column));
        }

        private static Dictionary<(int Row, int Column), T> SelectCellVisuals<T>(
            IReadOnlyDictionary<(int Row, int Column), T> source, IReadOnlyList<int> rows, IReadOnlyList<int> columns) {
            var result = new Dictionary<(int Row, int Column), T>();
            if (source.Count == 0) return result;
            for (int row = 0; row < rows.Count; row++)
                for (int column = 0; column < columns.Count; column++)
                    if (source.TryGetValue((rows[row], columns[column]), out T? value)) result[(row, column)] = value;
            return result;
        }
    }
}
