using System.Collections.Generic;

namespace OfficeIMO.Word.Pdf {
    internal sealed class TableLayout {
        internal TableLayout(List<IReadOnlyList<WordTableCell>> rows, float[] columnWidths, int[] rowStartColumns) {
            Rows = rows;
            ColumnWidths = columnWidths;
            RowStartColumns = rowStartColumns;
        }

        internal List<IReadOnlyList<WordTableCell>> Rows { get; }

        internal float[] ColumnWidths { get; }

        internal int[] RowStartColumns { get; }

        internal int GetRowStartColumn(int rowIndex) =>
            rowIndex >= 0 && rowIndex < RowStartColumns.Length
                ? RowStartColumns[rowIndex]
                : 0;
    }
}

