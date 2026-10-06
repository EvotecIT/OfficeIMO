using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static void ApplyGeneralCellAlignments(PdfCore.PdfTableStyle style, object?[,] values, IReadOnlyList<int> rows) {
            var alignments = style.CellAlignments ?? new Dictionary<(int Row, int Column), PdfCore.PdfColumnAlign>();
            for (int localRow = 0; localRow < rows.Count; localRow++) {
                int row = rows[localRow];
                if (row < 0 || row >= values.GetLength(0)) continue;
                for (int column = 0; column < values.GetLength(1); column++) {
                    if (alignments.ContainsKey((localRow, column))) continue;
                    object? value = values[row, column];
                    if (value is bool) alignments[(localRow, column)] = PdfCore.PdfColumnAlign.Center;
                    else if (value is DateTime || value is DateTimeOffset || value is TimeSpan
                        || value != null && value is not string && TryGetDouble(value, out _)) {
                        alignments[(localRow, column)] = PdfCore.PdfColumnAlign.Right;
                    }
                }
            }
            if (alignments.Count > 0) style.CellAlignments = alignments;
        }
    }
}
