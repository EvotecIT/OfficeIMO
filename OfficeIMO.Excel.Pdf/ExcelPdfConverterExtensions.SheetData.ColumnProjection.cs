namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        // Project a page's columns once so cells, formatting, merges and conditional visuals
        // use the same coordinates even when title columns precede a distant body segment.
        private static SheetExportData SelectPageColumns(SheetExportData source, IReadOnlyList<int> columns) {
            if (columns.Count == source.Values.GetLength(1) && columns.Select((column, index) => column == index).All(matches => matches)) return source;
            int rows = source.Values.GetLength(0);
            MergeLayoutData? merged = null;
            if (source.MergedCells != null) {
                merged = new MergeLayoutData(rows, columns.Count);
                for (int row = 0; row < rows; row++) {
                    for (int localColumn = 0; localColumn < columns.Count; localColumn++) {
                        int column = columns[localColumn];
                        MergeSpan? span = source.MergedCells.GetSpan(row, column);
                        if (span == null) continue;
                        int count = 1;
                        while (count < span.ColumnSpan && localColumn + count < columns.Count && columns[localColumn + count] == column + count) count++;
                        merged.SetSpan(row, localColumn, span.RowSpan, count);
                    }
                }
            }
            ColumnLayoutData? widths = source.ColumnWidths == null ? null : new ColumnLayoutData(
                columns.Select(column => source.ColumnWidths.WidthWeights[column]).ToList(),
                source.ColumnWidths.ApproximateWidthPoints * columns.Sum(column => source.ColumnWidths.WidthWeights[column])
                    / Math.Max(1D, source.ColumnWidths.WidthWeights.Sum()));
            ConditionalFillData? conditional = source.ConditionalFills == null ? null : new ConditionalFillData(
                SelectColumnVisuals(source.ConditionalFills.FillColors, columns),
                SelectColumnVisuals(source.ConditionalFills.DataBars, columns),
                SelectColumnVisuals(source.ConditionalFills.Icons, columns));
            return new SheetExportData(
                SelectColumns(source.Values, columns)!, SelectColumns(source.Styles, columns),
                SelectColumns(source.Hyperlinks, columns), SelectColumns(source.CellReferences, columns),
                merged, widths, source.RowHeights, source.HeaderRowCount, source.FirstBodyRowNumber,
                source.StructuredTables, conditional);
        }

        private static T[,]? SelectColumns<T>(T[,]? source, IReadOnlyList<int> columns) {
            if (source == null) return null;
            var result = new T[source.GetLength(0), columns.Count];
            for (int row = 0; row < source.GetLength(0); row++) {
                for (int column = 0; column < columns.Count; column++) result[row, column] = source[row, columns[column]];
            }
            return result;
        }

        private static IReadOnlyDictionary<(int Row, int Column), T> SelectColumnVisuals<T>(
            IReadOnlyDictionary<(int Row, int Column), T> source, IReadOnlyList<int> columns) {
            var localColumns = columns.Select((column, index) => (column, index)).ToDictionary(item => item.column, item => item.index);
            var result = new Dictionary<(int Row, int Column), T>();
            foreach (var item in source) {
                if (localColumns.TryGetValue(item.Key.Column, out int column)) result[(item.Key.Row, column)] = item.Value;
            }
            return result;
        }
    }
}
