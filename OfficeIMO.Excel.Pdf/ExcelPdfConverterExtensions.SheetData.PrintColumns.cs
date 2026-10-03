using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static RangeExportData ReadRangeWithPrintTitleColumns(
            ExcelSheetReader reader, ExcelSheet? metadataSheet, ExcelSheet? styleSheet,
            string range, ExcelToPdfOptions options, PdfCore.PdfStandardFont defaultFont, bool boundedRead) {
            RangeExportData body = ReadRangeExportData(reader, metadataSheet, styleSheet, range, options, defaultFont, boundedRead);
            ExcelPrintTitles? titles = metadataSheet?.GetPrintTitles();
            if (!options.UseWorksheetPrintTitleColumns || titles?.HasColumns != true ||
                !A1.TryParseRange(range, out int firstRow, out int firstColumn, out int lastRow, out _)
                || titles.FirstColumn!.Value >= firstColumn) return body;

            string titleRange = ToA1Range(firstRow, titles.FirstColumn!.Value, lastRow,
                Math.Min(titles.LastColumn!.Value, firstColumn - 1));
            RangeExportData left = ReadRangeExportData(reader, metadataSheet, styleSheet, titleRange, options, defaultFont, boundedRead);
            int rows = body.Values.GetLength(0);
            int leftColumns = left.Values.GetLength(1);
            int bodyColumns = body.Values.GetLength(1);
            MergeLayoutData? merged = null;
            if (left.MergedCells != null || body.MergedCells != null) {
                merged = new MergeLayoutData(rows, leftColumns + bodyColumns);
                left.MergedCells?.CopyTo(merged, 0);
                body.MergedCells?.CopyTo(merged, 0, leftColumns);
            }
            ColumnLayoutData? widths = null;
            if (left.ColumnWidths != null || body.ColumnWidths != null) {
                var weights = (left.ColumnWidths?.WidthWeights ?? Enumerable.Repeat(8.43D, leftColumns).ToList())
                    .Concat(body.ColumnWidths?.WidthWeights ?? Enumerable.Repeat(8.43D, bodyColumns).ToList()).ToList();
                widths = new ColumnLayoutData(weights, weights.Sum() * 5.25D);
            }
            return new RangeExportData(
                JoinColumns(left.Values, body.Values, rows, leftColumns, bodyColumns)!,
                JoinColumns(left.Styles, body.Styles, rows, leftColumns, bodyColumns),
                JoinColumns(left.Hyperlinks, body.Hyperlinks, rows, leftColumns, bodyColumns),
                JoinColumns(left.CellReferences, body.CellReferences, rows, leftColumns, bodyColumns),
                merged, widths, body.RowHeights ?? left.RowHeights);
        }

        private static T[,]? JoinColumns<T>(T[,]? left, T[,]? right, int rows, int leftColumns, int rightColumns) {
            if (left == null && right == null) return null;
            var result = new T[rows, leftColumns + rightColumns];
            for (int row = 0; row < rows; row++) {
                for (int column = 0; column < leftColumns; column++) {
                    if (left != null) result[row, column] = left[row, column];
                }
                for (int column = 0; column < rightColumns; column++) {
                    if (right != null) result[row, leftColumns + column] = right[row, column];
                }
            }
            return result;
        }

        private static IReadOnlyList<int> GetPrintTitleColumnIndexes(ExcelSheet? sheet, string?[,]? references, ExcelToPdfOptions options) {
            ExcelPrintTitles? titles = sheet?.GetPrintTitles();
            if (!options.UseWorksheetPrintTitleColumns || titles?.HasColumns != true || references == null) return Array.Empty<int>();
            var indexes = new List<int>();
            for (int column = 0; column < references.GetLength(1); column++) {
                int source = GetOriginalColumnNumber(references, column, references.GetLength(0));
                if (source >= titles.FirstColumn!.Value && source <= titles.LastColumn!.Value) indexes.Add(column);
            }
            return indexes;
        }

        private static IReadOnlyList<int> CreateChunkColumnIndexes(WorksheetPdfExportPlan plan, int start, int count) =>
            plan.ExportData.PrintTitleColumnIndexes.Where(column => column < start || column >= start + count)
                .Concat(Enumerable.Range(start, count)).ToArray();
    }
}
