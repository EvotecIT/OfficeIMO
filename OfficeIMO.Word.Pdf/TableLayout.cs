using System.Collections.Generic;

namespace OfficeIMO.Word.Pdf {
    internal sealed class TableLayout {
        internal TableLayout(List<IReadOnlyList<WordTableCell>> rows, float[] columnWidths, int[] rowStartColumns, int[] rowTrailingColumns) {
            Rows = rows;
            ColumnWidths = columnWidths;
            RowStartColumns = rowStartColumns;
            RowTrailingColumns = rowTrailingColumns;
        }

        internal List<IReadOnlyList<WordTableCell>> Rows { get; }

        internal float[] ColumnWidths { get; }

        internal int[] RowStartColumns { get; }

        internal int[] RowTrailingColumns { get; }

        internal int GetRowStartColumn(int rowIndex) =>
            rowIndex >= 0 && rowIndex < RowStartColumns.Length
                ? RowStartColumns[rowIndex]
                : 0;

        internal int GetRowTrailingColumnCount(int rowIndex) =>
            rowIndex >= 0 && rowIndex < RowTrailingColumns.Length
                ? RowTrailingColumns[rowIndex]
                : 0;
    }
}

